"""บันทึกการใช้งานโปรแกรม (ต่องาน 1 แถว) → MySQL เพื่อนำไปคิดค่าใช้จ่าย

หลักการ:
  1. ทุกครั้งที่งานจบ (สำเร็จ/ยกเลิก/error) เขียน 1 บรรทัดลง usage/usage_pending.jsonl ในเครื่องก่อนเสมอ
  2. แล้วค่อยพยายามส่งเข้า MySQL ใน thread แยก (INSERT IGNORE กันซ้ำด้วย event_id)
  3. ส่งสำเร็จ → ย้ายบรรทัดไป usage/usage_sent_YYYYMM.jsonl (สำเนาไว้ตรวจสอบ)
     ส่งไม่ได้ (ไม่มีเน็ต/DB ล่ม) → ค้างไว้ในคิว ส่งใหม่ตอนเปิดโปรแกรมครั้งถัดไปหรืองานถัดไป

ตั้งค่าใน .env (ดู .env.example):
  USAGE_DB_HOST / USAGE_DB_PORT / USAGE_DB_NAME / USAGE_DB_USER / USAGE_DB_PASSWORD
  USAGE_ENABLED=true|false   (ปิดได้ชั่วคราว — ยังเขียนไฟล์คิวในเครื่องแต่ไม่ส่ง)

สร้างตารางด้วย sql/usage_schema.sql — แนะนำให้ user ของ MySQL ที่ใส่ใน .env มีสิทธิ์ INSERT อย่างเดียว
"""
from __future__ import annotations

import getpass
import json
import os
import platform
import socket
import threading
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

try:
    from dotenv import load_dotenv
except Exception:  # pragma: no cover
    load_dotenv = None  # type: ignore

ROOT = Path(__file__).parent
USAGE_DIR = ROOT / "usage"
PENDING_FILE = USAGE_DIR / "usage_pending.jsonl"
INSTALL_ID_FILE = USAGE_DIR / "install_id.txt"

MODE_GROUPS_BY_MODE = {
    "etracking": "report", "permit_report": "report", "aliens": "report",
    "namelist_alien": "report", "ticket": "report",
    "receipts": "download", "results": "download", "bill_payment": "download",
    "payment_receipts": "download", "bt30_ctn": "download", "appointment": "download",
    "register": "submit", "inform": "submit", "bt30": "submit", "bt44": "submit",
    "booking_availability": "booking",
}

_lock = threading.Lock()
_flush_lock = threading.Lock()
LogFn = Callable[[str], None]


# ---------------------------------------------------------------------------
# config
# ---------------------------------------------------------------------------
def _env(name: str, default: str = "") -> str:
    return (os.getenv(name) or default).strip()


def load_db_config() -> dict | None:
    """อ่านค่าเชื่อมต่อจาก .env — คืน None ถ้ายังไม่ได้ตั้งค่า/ปิดไว้"""
    if load_dotenv is not None:
        load_dotenv(ROOT / ".env")
    if _env("USAGE_ENABLED", "true").lower() in ("0", "false", "no", "off"):
        return None
    host = _env("USAGE_DB_HOST")
    name = _env("USAGE_DB_NAME")
    user = _env("USAGE_DB_USER")
    if not (host and name and user):
        return None
    try:
        port = int(_env("USAGE_DB_PORT", "3306"))
    except ValueError:
        port = 3306
    return {
        "host": host, "port": port, "database": name, "user": user,
        "password": os.getenv("USAGE_DB_PASSWORD") or "",
        "connect_timeout": 8,
    }


def is_configured() -> bool:
    return load_db_config() is not None


def pending_count() -> int:
    try:
        if not PENDING_FILE.exists():
            return 0
        return sum(1 for ln in PENDING_FILE.read_text(encoding="utf-8").splitlines() if ln.strip())
    except Exception:
        return 0


# ---------------------------------------------------------------------------
# identity
# ---------------------------------------------------------------------------
def install_id() -> str:
    """UUID ประจำเครื่อง/การติดตั้ง — สร้างครั้งแรกแล้วเก็บไว้ในไฟล์"""
    try:
        if INSTALL_ID_FILE.exists():
            v = INSTALL_ID_FILE.read_text(encoding="utf-8").strip()
            if len(v) >= 32:
                return v
        USAGE_DIR.mkdir(parents=True, exist_ok=True)
        v = str(uuid.uuid4())
        INSTALL_ID_FILE.write_text(v, encoding="utf-8")
        return v
    except Exception:
        return "unknown"


def default_operator() -> str:
    try:
        return getpass.getuser()
    except Exception:
        return ""


def _machine() -> str:
    try:
        return socket.gethostname()
    except Exception:
        return platform.node() or ""


# ---------------------------------------------------------------------------
# job lifecycle
# ---------------------------------------------------------------------------
def begin_job(mode: str, operator: str, ewp_username: str = "",
              login_file: str | Path | None = None, out_path: str | Path | None = None,
              app_version: str = "") -> dict[str, Any]:
    """เรียกตอนกดเริ่มงาน — คืน meta ที่ต้องส่งกลับมาตอน end_job"""
    accounts: list[str] = []
    if login_file:
        try:
            from scrape_wa import _read_login_accounts
            accts = _read_login_accounts(Path(login_file))
            accounts = sorted({str(a.get("username", "")).strip() for a in accts.values()
                               if a.get("username")})
        except Exception:
            accounts = []
    if ewp_username and ewp_username not in accounts:
        accounts.insert(0, ewp_username)
    return {
        "event_id": str(uuid.uuid4()),
        "install_id": install_id(),
        "machine": _machine()[:64],
        "windows_user": default_operator()[:64],
        "operator": (operator or default_operator())[:100],
        "app_version": str(app_version)[:20],
        "mode": mode[:40],
        "mode_group": MODE_GROUPS_BY_MODE.get(mode, ""),
        "ewp_username": (ewp_username or (accounts[0] if accounts else ""))[:190],
        "accounts": accounts,
        "login_file": (Path(login_file).name if login_file else "")[:255],
        "out_file": (str(out_path) if out_path else "")[:255],
        "started_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "_t0": datetime.now().timestamp(),
    }


def end_job(meta: dict[str, Any] | None, status: str, count: int = 0,
            out_path: str | Path | None = None, progress: tuple[int, int] = (0, 0),
            error: str = "", log: LogFn | None = None) -> None:
    """เรียกตอนงานจบ (status: ok / cancelled / error) — เขียนคิว + ส่งแบบไม่บล็อก"""
    if not meta:
        return
    now = datetime.now()
    row = {k: v for k, v in meta.items() if not k.startswith("_")}
    row.update({
        "finished_at": now.strftime("%Y-%m-%d %H:%M:%S"),
        "duration_sec": int(max(0, now.timestamp() - float(meta.get("_t0", now.timestamp())))),
        "status": status if status in ("ok", "cancelled", "error") else "error",
        "record_count": int(count or 0),
        "progress_done": int(progress[0] or 0) if progress else 0,
        "progress_total": int(progress[1] or 0) if progress else 0,
        "error_msg": (error or "")[:2000],
    })
    if out_path:
        row["out_file"] = str(out_path)[:255]
    _append_pending(row)
    flush_async(log=log)


def _append_pending(row: dict[str, Any]) -> None:
    with _lock:
        USAGE_DIR.mkdir(parents=True, exist_ok=True)
        with PENDING_FILE.open("a", encoding="utf-8") as f:
            f.write(json.dumps(row, ensure_ascii=False) + "\n")


# ---------------------------------------------------------------------------
# flush → MySQL
# ---------------------------------------------------------------------------
_INSERT_SQL = """
INSERT IGNORE INTO usage_runs
  (event_id, install_id, machine, windows_user, operator, app_version,
   mode, mode_group, ewp_username, accounts, login_file,
   started_at, finished_at, duration_sec, status,
   record_count, progress_done, progress_total, out_file, error_msg)
VALUES
  (%(event_id)s, %(install_id)s, %(machine)s, %(windows_user)s, %(operator)s, %(app_version)s,
   %(mode)s, %(mode_group)s, %(ewp_username)s, %(accounts)s, %(login_file)s,
   %(started_at)s, %(finished_at)s, %(duration_sec)s, %(status)s,
   %(record_count)s, %(progress_done)s, %(progress_total)s, %(out_file)s, %(error_msg)s)
"""


def _read_pending() -> list[dict[str, Any]]:
    if not PENDING_FILE.exists():
        return []
    rows: list[dict[str, Any]] = []
    for ln in PENDING_FILE.read_text(encoding="utf-8").splitlines():
        ln = ln.strip()
        if not ln:
            continue
        try:
            rows.append(json.loads(ln))
        except Exception:
            continue  # บรรทัดเสีย — ทิ้ง
    return rows


def _write_pending(rows: list[dict[str, Any]]) -> None:
    USAGE_DIR.mkdir(parents=True, exist_ok=True)
    tmp = PENDING_FILE.with_suffix(".tmp")
    with tmp.open("w", encoding="utf-8") as f:
        for r in rows:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")
    tmp.replace(PENDING_FILE)


def _archive_sent(rows: list[dict[str, Any]]) -> None:
    if not rows:
        return
    try:
        p = USAGE_DIR / f"usage_sent_{datetime.now():%Y%m}.jsonl"
        with p.open("a", encoding="utf-8") as f:
            for r in rows:
                f.write(json.dumps(r, ensure_ascii=False) + "\n")
    except Exception:
        pass


def flush(log: LogFn | None = None) -> tuple[int, int]:
    """ส่งคิวทั้งหมดเข้า MySQL — คืน (จำนวนที่ส่งได้, จำนวนที่ยังค้าง)"""
    cfg = load_db_config()
    with _flush_lock:
        with _lock:
            rows = _read_pending()
        if not rows:
            return 0, 0
        if cfg is None:
            return 0, len(rows)
        try:
            import pymysql  # type: ignore
        except Exception:
            if log:
                log("[usage] ✗ ยังไม่ได้ติดตั้ง PyMySQL — รัน: python -m pip install -r requirements.txt")
            return 0, len(rows)
        try:
            conn = pymysql.connect(
                host=cfg["host"], port=cfg["port"], user=cfg["user"],
                password=cfg["password"], database=cfg["database"],
                charset="utf8mb4", connect_timeout=cfg["connect_timeout"],
                read_timeout=15, write_timeout=15, autocommit=False,
            )
        except Exception as e:
            if log:
                log(f"[usage] ✗ ต่อ MySQL ไม่ได้ ({e.__class__.__name__}) — เก็บไว้ส่งรอบหน้า ({len(rows)} รายการ)")
            return 0, len(rows)
        sent: list[dict[str, Any]] = []
        failed: list[dict[str, Any]] = []
        try:
            with conn.cursor() as cur:
                for r in rows:
                    params = {
                        "event_id": r.get("event_id", str(uuid.uuid4())),
                        "install_id": r.get("install_id", ""),
                        "machine": r.get("machine", ""),
                        "windows_user": r.get("windows_user", ""),
                        "operator": r.get("operator", ""),
                        "app_version": r.get("app_version", ""),
                        "mode": r.get("mode", ""),
                        "mode_group": r.get("mode_group", ""),
                        "ewp_username": r.get("ewp_username", ""),
                        "accounts": json.dumps(r.get("accounts") or [], ensure_ascii=False),
                        "login_file": r.get("login_file", ""),
                        "started_at": r.get("started_at"),
                        "finished_at": r.get("finished_at"),
                        "duration_sec": int(r.get("duration_sec") or 0),
                        "status": r.get("status", "error"),
                        "record_count": int(r.get("record_count") or 0),
                        "progress_done": int(r.get("progress_done") or 0),
                        "progress_total": int(r.get("progress_total") or 0),
                        "out_file": r.get("out_file", ""),
                        "error_msg": r.get("error_msg", ""),
                    }
                    try:
                        cur.execute(_INSERT_SQL, params)
                        sent.append(r)
                    except Exception as e:
                        # ข้อมูลผิดรูป (เช่น ค่าเกินความยาว) — เก็บ error ไว้ในแถวแล้วลองใหม่รอบหน้า
                        r["_last_error"] = f"{e.__class__.__name__}: {e}"[:300]
                        failed.append(r)
            conn.commit()
        except Exception as e:
            try:
                conn.rollback()
            except Exception:
                pass
            if log:
                log(f"[usage] ✗ ส่งข้อมูลไม่สำเร็จ ({e.__class__.__name__}: {e}) — เก็บไว้ส่งรอบหน้า")
            return 0, len(rows)
        finally:
            try:
                conn.close()
            except Exception:
                pass
        with _lock:
            # เผื่อมีแถวใหม่ถูกเพิ่มระหว่างส่ง → อ่านคิวปัจจุบันแล้วตัดเฉพาะที่ส่งแล้วออก
            sent_ids = {r.get("event_id") for r in sent}
            current = _read_pending()
            remain = [r for r in current if r.get("event_id") not in sent_ids]
            # แถวที่ fail ให้ใช้เวอร์ชันที่มี _last_error
            failed_by_id = {r.get("event_id"): r for r in failed}
            remain = [failed_by_id.get(r.get("event_id"), r) for r in remain]
            _write_pending(remain)
        _archive_sent(sent)
        if log:
            if sent:
                log(f"[usage] ✓ บันทึกการใช้งานเข้า MySQL {len(sent)} รายการ"
                    + (f" (ค้าง {len(remain)})" if remain else ""))
            elif remain:
                log(f"[usage] ✗ ส่งไม่ได้ {len(remain)} รายการ — {remain[0].get('_last_error', '')}")
        return len(sent), len(remain)


def flush_async(log: LogFn | None = None) -> threading.Thread | None:
    """ส่งคิวใน thread แยก (ไม่บล็อก UI) — ข้ามถ้าไม่มีอะไรค้าง"""
    if pending_count() == 0:
        return None
    t = threading.Thread(target=flush, kwargs={"log": log}, daemon=True, name="usage-flush")
    t.start()
    return t


def test_connection() -> tuple[bool, str]:
    """ทดสอบเชื่อมต่อ + ตรวจว่ามีตาราง usage_runs — ใช้จากปุ่มในหน้าจอหรือ CLI"""
    cfg = load_db_config()
    if cfg is None:
        return False, "ยังไม่ได้ตั้งค่า USAGE_DB_* ใน .env (หรือ USAGE_ENABLED=false)"
    try:
        import pymysql  # type: ignore
    except Exception:
        return False, "ยังไม่ได้ติดตั้ง PyMySQL — รัน: python -m pip install -r requirements.txt"
    try:
        conn = pymysql.connect(
            host=cfg["host"], port=cfg["port"], user=cfg["user"], password=cfg["password"],
            database=cfg["database"], charset="utf8mb4", connect_timeout=cfg["connect_timeout"],
        )
        try:
            with conn.cursor() as cur:
                cur.execute("SHOW COLUMNS FROM usage_runs")  # ใช้ได้แม้มีสิทธิ์แค่ INSERT
        finally:
            conn.close()
        return True, f"เชื่อมต่อสำเร็จ: {cfg['user']}@{cfg['host']}:{cfg['port']}/{cfg['database']}"
    except Exception as e:
        return False, f"{e.__class__.__name__}: {e}"


if __name__ == "__main__":
    import sys
    ok, msg = test_connection()
    print(("✓ " if ok else "✗ ") + msg)
    if ok:
        s, r = flush(log=print)
        print(f"ส่งแล้ว {s} รายการ, ค้าง {r} รายการ")
    sys.exit(0 if ok else 1)
