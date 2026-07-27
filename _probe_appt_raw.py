"""Dump raw innerHTML ของ #tab_default_5 (การนัดหมาย) เพื่อดู hidden content
+ ทุก item ใน #tab_default_4 (การชำระเงิน) แบบละเอียด
"""
from __future__ import annotations
import os
import sys
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import login  # noqa: E402

TARGET_URL = os.getenv(
    "PROBE_APPT_URL",
    "https://eworkpermit.doe.go.th/RenewMOU/DetailFormRenewMOU?user_id=303816&status=APSS&group_id=69123000059294&form_type=MT_59_MOU_RENEWAL",
)


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }
    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, args=["--start-maximized"])
        ctx = b.new_context(locale="th-TH", viewport={"width": 1600, "height": 1000})
        page = ctx.new_page()
        try:
            login(page, cfg)
            page.wait_for_timeout(1500)
            page.goto(TARGET_URL, wait_until="domcontentloaded", timeout=45_000)
            page.wait_for_timeout(4000)

            # dump raw innerHTML ของ tab_default_5
            html5 = page.evaluate("() => (document.querySelector('#tab_default_5') || {}).innerHTML || ''")
            print("=" * 60)
            print("RAW innerHTML of #tab_default_5:")
            print("=" * 60)
            print(html5)

            print("\n" + "=" * 60)
            print("RAW innerHTML of #tab_default_4:")
            print("=" * 60)
            html4 = page.evaluate("() => (document.querySelector('#tab_default_4') || {}).innerHTML || ''")
            print(html4)

            input("\n Enter to close...")
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
