"""ทดสอบฟีเจอร์ใหม่: โหมด e-Tracking ดึงตามเลขคำขอ (cfg["req_no_list"]) ผ่าน run_scrape จริง
ใช้: python _probe_etracking_reqno.py <เลขคำขอ1> <เลขคำขอ2> ... (ไม่ใส่ = ใช้ค่า default 4 เลขจากผู้ใช้)
login จาก .env, แก้ reCAPTCHA เอง ~120s, ผลลัพธ์ไปที่ reports/_probe_etracking_reqno.xlsx
"""
import os
import sys
from pathlib import Path
from dotenv import load_dotenv

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import run_scrape, REPORTS_DIR  # noqa: E402

DEFAULT_REQ_NOS = [
    "69123000065113",
    "69123000065021",
    "69123000065017",
    "69123000065010",
]


def main() -> int:
    req_nos = [a.strip() for a in sys.argv[1:] if a.strip()] or DEFAULT_REQ_NOS
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", ""),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
        "headless": False,
        "login_timeout_ms": 120_000,
        "req_no_list": "\n".join(req_nos),
    }
    if not cfg["username"] or not cfg["password"]:
        print("ERR: EWP_USERNAME/EWP_PASSWORD ไม่ถูกตั้งใน .env")
        return 1
    print(f"[i] บัญชี: {cfg['username']} ({cfg['user_type']})")
    print(f"[i] เลขคำขอที่จะทดสอบ ({len(req_nos)}): {req_nos}")

    out_path = REPORTS_DIR / "_probe_etracking_reqno.xlsx"
    count, path = run_scrape(cfg, out_path, log=print)
    print(f"\n[เสร็จสิ้น] ดึงได้ {count} แถว → {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
