"""Smoke test: run_booking_availability
- 2 เดือน จาก .env credentials
"""
from __future__ import annotations
import os, sys
from pathlib import Path
from dotenv import load_dotenv

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()

from scrape_wa import run_booking_availability

cfg = {
    "username": os.getenv("EWP_USERNAME", "").strip(),
    "password": os.getenv("EWP_PASSWORD", "").strip(),
    "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
    "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    "headless": False,
}

out_path = Path("reports/_test_booking_availability.xlsx")
out_path.parent.mkdir(parents=True, exist_ok=True)

try:
    count, path = run_booking_availability(
        cfg, login_excel=None,   # ใช้ .env credentials
        out_path=out_path,
        months_ahead=3,          # 3 เดือน
        branch_filter=None,      # ทุกสาขา
        log=print,
    )
except Exception as e:
    import traceback; traceback.print_exc()
    sys.exit(2)

print(f"\n[✓] rows = {count}, report = {path}")

# แสดง top 5 rows ที่ยังว่างเยอะ
from openpyxl import load_workbook
wb = load_workbook(path)
ws = wb["ตารางว่างจอง"]
print(f"\nrows in report: {ws.max_row - 1}")
print("\nsummary sheet (5 first):")
ws2 = wb["สรุปตามสาขา"]
for i, row in enumerate(ws2.iter_rows(min_row=1, max_row=6, values_only=True), 1):
    print("  " + " | ".join(str(c)[:35] if c is not None else "-" for c in row))
