"""ตรวจ: ในบัญชี 2 มี group_status_id ค่าอะไรบ้าง (ครอบคลุมทุก checkbox)"""
import sys
from pathlib import Path
from collections import Counter
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
from scrape_wa import (login, goto_tracking, apply_wa_filter,
                        _wait_datatable_idle, collect_all_wa_rows)

wb0 = load_workbook("UsernameLogin.xlsx", data_only=True)
ws0 = wb0.active
row2 = [c.value for c in ws0[2]]
cfg = {"username": str(row2[0]), "password": str(row2[1]),
       "user_type": str(row2[2]), "method": str(row2[3]),
       "headless": False}

with sync_playwright() as pw:
    b = pw.chromium.launch(headless=False,
                           args=["--ignore-certificate-errors", "--start-maximized"])
    ctx = b.new_context(locale="th-TH", ignore_https_errors=True,
                        viewport={"width":1600,"height":1000})
    page = ctx.new_page()
    login(page, cfg)
    goto_tracking(page)
    page.wait_for_timeout(3500)

    # ทุก checkbox
    apply_wa_filter(page, "", status_ids=["WP","WCOSNA","WA","AP","SS"])
    _wait_datatable_idle(page, max_wait_ms=30000)
    rows = collect_all_wa_rows(page)
    print(f"\ntotal rows collected: {len(rows)}")

    counts = Counter(str(r.get("status") or "?").upper() for r in rows)
    print("group_status_id distribution:")
    for k, v in counts.most_common():
        print(f"  {k:10s} → {v}")

    # ตัวอย่าง row ที่ WP2 (ถ้ามี)
    wp2 = [r for r in rows if str(r.get("status") or "").upper() == "WP2"]
    print(f"\nWP2 rows: {len(wp2)}")
    for r in wp2[:5]:
        print(f"  reqNo={r.get('reqNo')} form_type={r.get('form_type')} statusText={r.get('statusText','')[:50]!r}")

    page.wait_for_timeout(5000)
    ctx.close(); b.close()
