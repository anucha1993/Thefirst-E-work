"""Test: run_bt30_ctn with .env credentials on 1 request per account.
รัน: python _probe_bt30_ctn_test.py
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
    print(f"[i] ใช้บัญชี: {username} ({user_type} / {method})")
    print(f"[i] temp login file: {login_excel}")

    cfg = {
        "headless": False,
        "hide_window": False,
        # ไม่กรอง request_types/date → ดึงทุก request ในสถานะ WP
        "request_types": [],
        "date_from": "",
        "date_to": "",
        "method": method,
    }

    out_path = Path("reports/_test_bt30_ctn.xlsx")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    def _log(msg):
        print(msg)

    # จำกัดเป็นแค่ 1 row (แถวแรก) เพื่อทดสอบเร็ว
    try:
        count, path = run_bt30_ctn(
            cfg, login_excel, out_path,
            row_range="1",
            make_subfolder=False,
            log=_log,
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
    # list PDF ที่สร้าง
    save_dir = out_path.parent / "bt30_ctn"
    if save_dir.exists():
        pdfs = list(save_dir.glob("*.pdf"))
        print(f"\n[PDF ที่ดาวน์โหลด] {len(pdfs)} ไฟล์:")
        for p in pdfs:
            print(f"  - {p.name}  ({p.stat().st_size:,} bytes)")
    else:
        print(f"[!] ไม่พบโฟลเดอร์ {save_dir}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
