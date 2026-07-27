"""Smoke test: run_bt30_ctn โหมด appointment only 1 record
- do_ctn=False, do_appointment=True
- filter สถานะ AP+SS (ผ่าน cfg.filter_status_ids)
- limit 1 row
"""
from __future__ import annotations
import os
import sys
import tempfile
from pathlib import Path
from openpyxl import Workbook, load_workbook
from dotenv import load_dotenv

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()

from scrape_wa import run_bt30_ctn  # noqa: E402

username = os.getenv("EWP_USERNAME", "").strip()
password = os.getenv("EWP_PASSWORD", "").strip()
user_type = os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน")
method = os.getenv("EWP_LOGIN_METHOD", "E-Workpermit")

# สร้าง temp UsernameLogin เพื่อ pass เป็น 1 account
tmp = Path(tempfile.mktemp(suffix="_UsernameLogin.xlsx"))
wb = Workbook()
ws = wb.active
ws.append(["Username", "Password", "Type", "ระบบ"])
ws.append([username, password, user_type, method])
wb.save(tmp)

cfg = {
    "headless": False,
    "hide_window": False,
    "request_types": [],
    "date_from": "",
    "date_to": "",
    "method": method,
    # filter สถานะ = AP + SS (สำหรับดึงใบนัดหมาย ไม่ล็อก WP)
    "filter_status_ids": ["AP", "SS"],
}

out_path = Path("reports/_test_appt_only.xlsx")
out_path.parent.mkdir(parents=True, exist_ok=True)

print(f"[i] test: run_bt30_ctn — do_ctn=False, do_appointment=True, row=1, status=AP+SS")
try:
    count, path = run_bt30_ctn(
        cfg, tmp, out_path,
        row_range="23-25",  # AP=22 rows แรก แล้ว APSS ตามมา
        make_subfolder=False,
        do_ctn=False,
        do_appointment=True,
        log=print,
    )
except Exception as e:
    import traceback; traceback.print_exc()
    sys.exit(2)
finally:
    try: tmp.unlink()
    except Exception: pass

print(f"\n[✓] success = {count}, report = {path}")
save_dir = out_path.parent / "bt30_ctn"
if save_dir.exists():
    pdfs = list(save_dir.glob("*.pdf"))
    print(f"\n[PDF] {len(pdfs)} ไฟล์:")
    for p in pdfs:
        print(f"  · {p.name}  ({p.stat().st_size:,} bytes)")

wbr = load_workbook(path, data_only=True)
wsr = wbr.active
print(f"\n[Report {wsr.max_row - 1} rows]")
for r in wsr.iter_rows(min_row=1, values_only=True):
    print("  " + " | ".join(str(x)[:40] if x is not None else "-" for x in r))
