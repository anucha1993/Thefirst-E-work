"""Test /v1/round/get-rounds/ endpoint with actual token."""
from __future__ import annotations
import json, os, sys
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import login, ROOT, _booking_get_token

BE = "https://queue-be-uat.doe.go.th/doe-booking/api"


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }
    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=True)
        ctx = b.new_context(locale="th-TH")
        page = ctx.new_page()
        try:
            login(page, cfg)
            token = _booking_get_token(page, log=print)
            if not token:
                print("[!] no token"); return 1
            print(f"[+] token: {token[:60]}...")

            candidates = [
                ("MDH-OB-M-001", "2026-08-06"),
                ("MDH-OB-M-001", "2026-08-13"),
                ("CBI-SC-M-001", "2026-08-06"),
                ("CBI-SC-M-001", "2026-08-14"),
                ("NKI-OB-L-001", "2026-08-06"),
                ("TAK-OB-L-001", "2026-08-06"),
                ("RNG-SC-S-001", "2026-08-06"),
            ]
            for branch, date in candidates:
                url = f"{BE}/v1/round/get-rounds/?b_id={branch}&date_string={date}"
                try:
                    r = page.context.request.get(url, headers={
                        "Authorization": f"Bearer {token}",
                        "Accept": "application/json, text/plain, */*",
                    }, timeout=15_000)
                    try: body = r.text()
                    except Exception: body = ""
                    print(f"\n[{r.status}] {branch} / {date}")
                    if r.status == 200:
                        try:
                            j = json.loads(body)
                            print(json.dumps(j, ensure_ascii=False, indent=2)[:1500])
                        except Exception:
                            print(body[:1200])
                    else:
                        print(body[:400])
                except Exception as e:
                    print(f"  ERR: {e}")
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
