"""Test: บต.25 + Receipt download (run_bt30_ctn) — MT_59, สถานะ SS, 1 record, บัญชี .env.
รัน: python _probe_bt25_receipt_test.py
(เปิดเบราว์เซอร์จริง — ต้องแก้ reCAPTCHA เอง ~30 วิ)
"""
from __future__ import annotations
import os
import sys
import tempfile
from pathlib import Path
from openpyxl import Workbook
from dotenv import load_dotenv

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()

from scrape_wa import run_bt30_ctn  # noqa: E402


def build_temp_login_excel(username: str, password: str, user_type: str, method: str) -> Path:
    """สร้างไฟล์ UsernameLogin ชั่วคราวที่มี 1 บัญชีจาก .env"""
    tmp = Path(tempfile.mktemp(suffix="_UsernameLogin.xlsx"))
    wb = Workbook()
    ws = wb.active
    ws.append(["Username", "Password", "Type", "ระบบ"])
    ws.append([username, password, user_type, method])
    wb.save(tmp)
    return tmp


def main() -> int:
    username = os.getenv("EWP_USERNAME", "").strip()
    password = os.getenv("EWP_PASSWORD", "").strip()
    if not username or not password:
        print("ERR: EWP_USERNAME/EWP_PASSWORD ไม่ถูกตั้งใน .env")
        return 1

    user_type = os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน")
    method = os.getenv("EWP_LOGIN_METHOD", "E-Workpermit")

    login_excel = build_temp_login_excel(username, password, user_type, method)
    print(f"[i] บัญชี: {username} ({user_type} / {method})")
    print("[i] กรอง: request_type=MT_59 | สถานะ=SS | row=1")
    print("[i] เอกสาร: บต.25 + Receipt (รันรอบเดียว)")

    cfg = {
        "headless": False,
        "hide_window": False,
        "request_types": ["MT_59"],       # บต.25 ต้องใช้ MT_59
        "filter_status_ids": ["SS", "AP"],  # ดำเนินการเสร็จสิ้น + รอนัดหมาย
        "date_from": "",
        "date_to": "",
        "method": method,
        "login_timeout_ms": 120_000,      # ให้เวลาแก้ reCAPTCHA นานขึ้น (120 วิ)
    }

    out_path = Path("reports/_test_bt25_receipt.xlsx")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        count, path = run_bt30_ctn(
            cfg, login_excel, out_path,
            row_range="1",
            make_subfolder=False,
            do_ctn=False,
            do_appointment=False,
            do_bt25=True,
            do_receipt=True,
            log=print,
        )
    except Exception as e:
        import traceback
        print(f"[!] เกิด exception: {e}")
        traceback.print_exc()
        return 2
    finally:
        try:
            login_excel.unlink()
        except Exception:
            pass

    print(f"\n[✓] เสร็จ — success = {count}, report = {path}")
    save_dir = out_path.parent / "bt30_ctn"
    if save_dir.exists():
        pdfs = sorted(save_dir.glob("*.pdf"))
        print(f"[i] PDF ในโฟลเดอร์ bt30_ctn: {len(pdfs)} ไฟล์ (แสดง 10 ล่าสุด)")
        for p in pdfs[-10:]:
            print(f"    - {p.name} ({p.stat().st_size:,} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
