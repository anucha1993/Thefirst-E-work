"""probe e-Tracking — ทดสอบ filter ต่างๆ เพื่อหาว่าทำไมเก็บได้ 0 แถว"""
import os, sys
from dotenv import load_dotenv
sys.path.insert(0, ".")
load_dotenv()

from playwright.sync_api import sync_playwright
from scrape_wa import (
    login, goto_tracking, apply_wa_filter, set_all_page_length,
    collect_wa_rows, _is_logged_out,
)

cfg = {
    "username": os.getenv("EWP_USERNAME", ""),
    "password": os.getenv("EWP_PASSWORD", ""),
    "user_type": os.getenv("EWP_USER_TYPE", ""),
    "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    "headless": False,
}

scenarios = [
    ("AP only, no request_type", "", ["AP"]),
    ("AP + MT_63_2_1302_RENEWAL", "MT_63_2_1302_RENEWAL", ["AP"]),
    ("AP + MT_41_4_59", "MT_41_4_59", ["AP"]),
    ("All status, no request_type", "", ["WP","WCOSNA","WA","AP","SS"]),
]

with sync_playwright() as pw:
    b = pw.chromium.launch(headless=False, args=["--ignore-certificate-errors","--start-maximized"])
    ctx = b.new_context(locale="th-TH", ignore_https_errors=True, viewport={"width":1920,"height":1080})
    page = ctx.new_page()
    print(f"login as {cfg['username']!r}")
    login(page, cfg)
    print("login url:", page.url)
    for label, rt, ids in scenarios:
        print(f"\n=== {label} ===")
        try:
            goto_tracking(page)
            apply_wa_filter(page, rt, status_ids=ids)
            set_all_page_length(page)
            page.wait_for_timeout(2500)
            rows = collect_wa_rows(page)
            print(f"  rows = {len(rows)}")
            # show first 3 statusText / desc
            for i, r in enumerate(rows[:3], 1):
                print(f"  [{i}] reqNo={r.get('reqNo')!r} status={r.get('status')!r} statusText={(r.get('statusText') or '')[:80]!r}")
                print(f"      form_type={r.get('form_type')!r} desc={(r.get('desc') or '')[:60]!r}")
            # count by status text
            stat_counter = {}
            for r in rows:
                s = (r.get("statusText") or "").splitlines()[0][:40]
                stat_counter[s] = stat_counter.get(s, 0) + 1
            print("  status distribution:")
            for s, c in sorted(stat_counter.items(), key=lambda x: -x[1])[:8]:
                print(f"    {c:>4} × {s!r}")
        except Exception as e:
            print(f"  ERR: {e}")
    page.wait_for_timeout(2000)
    ctx.close(); b.close()
