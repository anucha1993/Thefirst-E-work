"""ทดสอบ run_bt30_ctn กับบัญชี Kanyaphat0 (บัญชีที่ 2 ใน UsernameLogin.xlsx)
ที่น่าจะมีคำขอในสถานะ WP จริง"""
from __future__ import annotations
import os, sys, tempfile
from pathlib import Path
from openpyxl import Workbook, load_workbook
from dotenv import load_dotenv

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()

from scrape_wa import run_bt30_ctn  # noqa: E402

# อ่านบัญชีที่ 2 จาก UsernameLogin.xlsx
wb0 = load_workbook("UsernameLogin.xlsx", data_only=True)
ws0 = wb0.active
hdr = [c.value for c in ws0[1]]
row2 = [c.value for c in ws0[2]]
print(f"[i] ใช้บัญชี: {row2[0]} ({row2[2]} / {row2[3]})")
if not row2[1]:
    print("[!] Password ว่าง")
    sys.exit(1)

tmp = Path(tempfile.mktemp(suffix="_UsernameLogin.xlsx"))
wb = Workbook()
ws = wb.active
ws.append(["Username","Password","Type","ระบบ"])
ws.append([str(row2[0]), str(row2[1]), str(row2[2]), str(row2[3])])
wb.save(tmp)

cfg = {"headless": False, "hide_window": False,
       "request_types": [], "date_from": "", "date_to": "",
       "method": str(row2[3])}

out_path = Path("reports/_test_bt30_ctn_acct2.xlsx")
out_path.parent.mkdir(parents=True, exist_ok=True)

try:
    count, path = run_bt30_ctn(cfg, tmp, out_path, row_range="1-3",
                                 make_subfolder=False, log=print)
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

# แสดงเนื้อหา report
wbr = load_workbook(path, data_only=True)
wsr = wbr.active
print("\n[Report]")
for r in wsr.iter_rows(min_row=1, values_only=True):
    print("  " + " | ".join(str(x)[:40] if x is not None else "-" for x in r))
