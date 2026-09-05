"""Probe 2: log raw response body + headers ของ calendar API เพื่อเข้าใจ auth
"""
from __future__ import annotations
import json
import os
import sys
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import login, ROOT  # noqa: E402

TARGET_URL = "https://eworkpermit.doe.go.th/RenewMOU/DetailFormRenewMOU?user_id=303816&status=APSS&group_id=69123000059307&form_type=MT_59_MOU_RENEWAL"
OUT_JSON = ROOT / "_probe_calendar_headers_out.json"


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }

    captured = []  # network requests + response headers

    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, args=["--start-maximized"])
        ctx = b.new_context(locale="th-TH", viewport={"width": 1600, "height": 1000})

        def on_request(req):
            u = req.url or ""
            if "queue-be" in u:
                captured.append({
                    "kind": "request",
                    "method": req.method,
                    "url": u[:250],
                    "headers": dict(req.headers),
                    "post": req.post_data or "",
                })

        def on_response(res):
            u = res.url or ""
            if "queue-be" in u:
                try:
                    text = res.text()
                except Exception:
                    text = ""
                captured.append({
                    "kind": "response",
                    "status": res.status,
                    "url": u[:250],
                    "headers": dict(res.headers),
                    "body_first": text[:1500],
                })

        ctx.on("request", on_request)
        ctx.on("response", on_response)

        page = ctx.new_page()
        try:
            login(page, cfg)
            page.wait_for_timeout(1500)
            page.goto(TARGET_URL, wait_until="domcontentloaded", timeout=45_000)
            page.wait_for_timeout(4000)

            page.evaluate("() => { if(window.jQuery) jQuery('a[href=\"#tab_default_5\"]').tab('show'); }")
            page.wait_for_timeout(3500)

            iframe_src = page.evaluate("() => document.querySelector('#link_appointment').src")

            qp = ctx.new_page()
            qp.goto(iframe_src, wait_until="networkidle", timeout=30_000)
            qp.wait_for_timeout(6000)  # รอ FE call APIs ทั้งหมด

            print(f"[+] captured {len(captured)} events")

            # หาข้อมูล calendar API call ที่ FE เรียกเอง — ดู headers ที่จริง
            calls = [c for c in captured if "calendar/get-data" in c.get("url", "")]
            print(f"[+] found {len(calls)} calendar/get-data events:")
            for c in calls:
                if c["kind"] == "request":
                    print(f"\n  REQ {c['method']} {c['url']}")
                    for h, v in (c["headers"] or {}).items():
                        if h.lower() in ("authorization", "cookie", "x-api-key", "x-auth-token", "x-token", "referer"):
                            print(f"    {h}: {v[:150]}")
                elif c["kind"] == "response":
                    print(f"  RES {c['status']}")
                    print(f"    body: {c['body_first'][:500]}")

            # ลอง call ตรง ๆ เพื่อดู response
            print("\n=== manual call: CCO-SC-S-001 2026-12 ===")
            res = qp.evaluate(r"""async () => {
                try {
                    const r = await fetch('/doe-booking/api/v2/calendar/get-data/CCO-SC-S-001/2026-12-01?month=12&year=2026', {
                        credentials: 'include'
                    });
                    return {status: r.status, body: await r.text()};
                } catch(e) { return {error: e.message}; }
            }""")
            print(f"  status={res.get('status')}")
            print(f"  body: {res.get('body', '')[:800]}")

            OUT_JSON.write_text(
                json.dumps(captured, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            print(f"\n[done] → {OUT_JSON}")

            input("\n Enter to close ...")
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
