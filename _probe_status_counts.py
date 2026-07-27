"""Quick check: statuses available for accounts (to find WP for BT30 CTN test)"""
from __future__ import annotations
import os
import sys
from pathlib import Path
from dotenv import load_dotenv
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()

from scrape_wa import (  # noqa: E402
    login, goto_tracking, apply_wa_filter,
    _wait_datatable_idle, _read_login_accounts,
)


def check_account(page, username: str, password: str, user_type: str, method: str) -> None:
    """login + count rows per status"""
    cfg = {"username": username, "password": password,
           "user_type": user_type, "method": method}
    print(f"\n=== {username} ({user_type} / {method}) ===")
    try:
        login(page, cfg)
    except Exception as e:
        print(f"  login FAIL: {e}")
        return
    goto_tracking(page)
    page.wait_for_timeout(2000)

    for status in ["WP", "WCOSNA", "WA", "AP", "SS"]:
        try:
            apply_wa_filter(page, "", status_ids=[status])
            _wait_datatable_idle(page, max_wait_ms=15000)
            total = page.evaluate(
                "() => { try { return window.jQuery('#datatableE_Tracking').DataTable().page.info().recordsTotal; } catch(e) { return -1; } }"
            )
            print(f"  {status:8s} → {total} รายการ")
        except Exception as e:
            print(f"  {status:8s} → ERROR: {str(e).splitlines()[0][:100]}")


def main() -> int:
    # ทดสอบทั้ง .env account + UsernameLogin.xlsx
    accts = []
    env_user = os.getenv("EWP_USERNAME", "").strip()
    if env_user:
        accts.append({
            "username": env_user,
            "password": os.getenv("EWP_PASSWORD", "").strip(),
            "type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
            "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
        })
    try:
        for _u, a in _read_login_accounts(Path("UsernameLogin.xlsx")).items():
            # skip duplicate with .env
            if a["username"] == env_user:
                continue
            accts.append(a)
    except Exception as e:
        print(f"[!] read UsernameLogin.xlsx: {e}")

    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False,
                               args=["--ignore-certificate-errors", "--start-maximized"])
        ctx = b.new_context(locale="th-TH", ignore_https_errors=True,
                            viewport={"width": 1600, "height": 1000})
        page = ctx.new_page()
        try:
            for a in accts:
                check_account(page, a["username"], a["password"], a["type"],
                              a.get("method") or "E-Workpermit")
                # logout for next
                try:
                    page.goto("https://eworkpermit.doe.go.th/Home/Logout",
                              wait_until="domcontentloaded", timeout=15000)
                    page.wait_for_timeout(1500)
                except Exception:
                    pass
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
