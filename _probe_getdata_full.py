"""Dump full GetDataRequestFormEtrackingForAlien source"""
import os, sys, json
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()

from scrape_wa import login, goto_tracking  # noqa: E402

with sync_playwright() as pw:
    b = pw.chromium.launch(headless=False,
                           args=["--ignore-certificate-errors", "--start-maximized"])
    ctx = b.new_context(locale="th-TH", ignore_https_errors=True,
                        viewport={"width":1600,"height":1000})
    page = ctx.new_page()
    login(page, {
        "username": os.getenv("EWP_USERNAME",""), "password": os.getenv("EWP_PASSWORD",""),
        "user_type": os.getenv("EWP_USER_TYPE","ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD","E-Workpermit"),
    })
    goto_tracking(page)
    page.wait_for_timeout(5000)

    full = page.evaluate("() => GetDataRequestFormEtrackingForAlien.toString()")
    Path("_probe_getdata_full.js").write_text(full, encoding="utf-8")
    print(f"[✓] wrote _probe_getdata_full.js ({len(full)} chars)")

    # หาส่วน 'data:' หรือ 'ajax:' ในฟังก์ชัน
    import re
    for m in re.finditer(r"(form_status_id|status_id|WP2|checked|isChecked|\.val\(\))", full):
        s = max(0, m.start()-150); e = min(len(full), m.end()+250)
        snippet = full[s:e].replace('\n', ' ')
        print(f"\n--- @{m.start()} match={m.group()!r} ---")
        print(snippet)

    ctx.close(); b.close()
