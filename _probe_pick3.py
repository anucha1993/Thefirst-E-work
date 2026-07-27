"""ดึงเลขคำขอสุ่ม 3 เลข (สถานะ AP=รอนัดหมาย) มาใส่ Ref_number.xlsx เพื่อทดสอบ"""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from playwright.sync_api import sync_playwright  # noqa: E402
from openpyxl import load_workbook  # noqa: E402
from scrape_wa import (  # noqa: E402
    load_config, login, goto_tracking, apply_wa_filter, collect_all_wa_rows, _launch_chromium,
)

# เลขคำขอเดิมที่เคยรัน — ต้องข้ามไม่ให้ซ้ำ
EXCLUDE = {"68115000962703", "68115000272962"}

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
        picks: list[str] = []
        # เลือกจากกลางลิสต์เพื่อเปลี่ยนความหลากหลาย (ไม่ใช่ 3 ตัวแรก)
        start = min(500, max(0, len(rows) // 3))
        for r in rows[start:]:
            rn = str(r.get("reqNo", "") or "").strip()
            if not rn or rn in EXCLUDE:
                continue
            picks.append(rn)
            if len(picks) >= 3:
                break
        print("PICKS:", picks)
        if len(picks) >= 3:
            wb = load_workbook("Ref_number.xlsx"); ws = wb.active
            # ล้างแถวเก่า (ยกเว้น header)
            max_row = ws.max_row
            for row_idx in range(2, max_row + 1):
                ws.cell(row=row_idx, column=1).value = None
                ws.cell(row=row_idx, column=2).value = None
            for i, rn in enumerate(picks, start=2):
                ws.cell(row=i, column=1, value=str(rn))
                ws.cell(row=i, column=2, value="woraphitcha.opall@gmail.com")
            wb.save("Ref_number.xlsx")
            print("Ref_number.xlsx updated")
            for r in ws.iter_rows(values_only=True):
                print("  ", r)
    finally:
        ctx.close(); br.close()
