"""ดึงเลขคำขอจริง 2 เลข (สถานะ AP หรือ SS) มาใส่ Ref_number.xlsx เพื่อรัน DOM demo"""
from pathlib import Path
from playwright.sync_api import sync_playwright
from openpyxl import load_workbook
import sys
sys.path.insert(0, str(Path(__file__).parent))
from scrape_wa import (  # noqa: E402
    load_config, login, goto_tracking, apply_wa_filter, collect_all_wa_rows, _launch_chromium,
)

cfg = load_config()
with sync_playwright() as pw:
    br = _launch_chromium(pw, cfg, ["--ignore-certificate-errors", "--start-maximized"])
    ctx = br.new_context(locale="th-TH", ignore_https_errors=True,
                         viewport={"width": 1920, "height": 1080}, accept_downloads=True)
    page = ctx.new_page()
    try:
        login(page, {
            "username": "woraphitcha.opall@gmail.com",
            "password": "Opwrpch13",
            "user_type": "ผู้กระทำการแทน",
            "method": "E-Workpermit",
        })
        goto_tracking(page)
        apply_wa_filter(page, "", status_ids=["AP"])
        rows = collect_all_wa_rows(page)
        ap = [(r.get("reqNo"), r.get("statusText")) for r in rows if "รอนัดหมาย" in str(r.get("statusText", ""))]
        print(f"AP rows: {len(ap)}")
        for r in ap[:10]:
            print("  ", r)
        picks = [r[0] for r in ap[:2]]
        print("PICKS:", picks)
        if len(picks) >= 2:
            wb = load_workbook("Ref_number.xlsx"); ws = wb.active
            ws.cell(row=2, column=1, value=str(picks[0]))
            ws.cell(row=2, column=2, value="woraphitcha.opall@gmail.com")
            ws.cell(row=3, column=1, value=str(picks[1]))
            ws.cell(row=3, column=2, value="woraphitcha.opall@gmail.com")
            wb.save("Ref_number.xlsx")
            print("Ref_number.xlsx updated with:", picks)
    finally:
        ctx.close(); br.close()
