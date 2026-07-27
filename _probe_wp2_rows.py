"""ตรวจ: ในบัญชี 2 (966 rows) มีสถานะย่อย WP2 กี่แถว?"""
import sys
from pathlib import Path
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

    # filter WP checkbox → collect
    print("=== filter=WP checkbox ===")
    apply_wa_filter(page, "", status_ids=["WP"])
    _wait_datatable_idle(page, max_wait_ms=15000)
    rows = collect_all_wa_rows(page)
    print(f"  collected: {len(rows)} rows")

    # นับ sub-status
    from collections import Counter
    counts = Counter(str(r.get("status") or "?").upper() for r in rows)
    print(f"  sub-status distribution:")
    for k, v in counts.most_common():
        print(f"    {k:10s} → {v}")

    # ตัวอย่าง 5 rows แรกที่ status=WP2 (ถ้ามี)
    wp2 = [r for r in rows if str(r.get("status") or "").upper() == "WP2"]
    print(f"\n  WP2 rows: {len(wp2)}")
    for r in wp2[:3]:
        print(f"    reqNo={r.get('reqNo')} form={r.get('form_type')} statusText={r.get('statusText','')[:60]!r}")

    page.wait_for_timeout(5000)
    ctx.close(); b.close()
