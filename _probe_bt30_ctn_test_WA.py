"""ทดสอบ download BT30 CTN กับ 1 row จริง (ใช้ WA แทน WP เพราะบัญชีนี้ WP=0)
ปรับ status_ids ชั่วคราวเพื่อพิสูจน์ว่าโฟลว์ download + rename ทำงาน"""
from __future__ import annotations
import os, sys, tempfile
from pathlib import Path
from openpyxl import Workbook
from dotenv import load_dotenv

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()

import scrape_wa
from scrape_wa import run_bt30_ctn  # noqa: E402

# monkey-patch run_bt30_ctn's apply_wa_filter call to use WA instead of WP (for testing)
# — วิธีที่ตรงกว่าคือแก้ status ใน run_bt30_ctn เป็น argument, แต่แค่ทดสอบพอ
_orig_apply = scrape_wa.apply_wa_filter
def _patched_apply(page, request_type="", status_ids=None, date_from="", date_to=""):
    if status_ids == ["WP"]:
        print("  [patch] เปลี่ยน status_ids จาก ['WP'] → ['WA'] เพื่อทดสอบ (WP=0)")
        status_ids = ["WA"]
    return _orig_apply(page, request_type=request_type, status_ids=status_ids,
                       date_from=date_from, date_to=date_to)
scrape_wa.apply_wa_filter = _patched_apply

# ── ต่อจากเดิม ──
username = os.getenv("EWP_USERNAME","").strip()
password = os.getenv("EWP_PASSWORD","").strip()
user_type = os.getenv("EWP_USER_TYPE","ผู้กระทำการแทน")
method = os.getenv("EWP_LOGIN_METHOD","E-Workpermit")

tmp = Path(tempfile.mktemp(suffix="_UsernameLogin.xlsx"))
wb = Workbook()
ws = wb.active
ws.append(["Username","Password","Type","ระบบ"])
ws.append([username, password, user_type, method])
wb.save(tmp)

cfg = {"headless": False, "hide_window": False,
       "request_types": [], "date_from": "", "date_to": "",
       "method": method}

out_path = Path("reports/_test_bt30_ctn_WA.xlsx")
out_path.parent.mkdir(parents=True, exist_ok=True)

try:
    count, path = run_bt30_ctn(cfg, tmp, out_path, row_range="1",
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
    for p in save_dir.glob("*.pdf"):
        print(f"  · {p.name}  ({p.stat().st_size:,} bytes)")
