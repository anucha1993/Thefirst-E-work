"""Save FE JS bundle to disk + grep for URL patterns (calendar/round/slot/appointment)."""
from __future__ import annotations
import os, re, sys
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import login, ROOT, goto_tracking, apply_wa_filter, collect_all_wa_rows, build_detail_url

BUNDLE_PATH = ROOT / "_fe_bundle.js"


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }
    bundle_text = None
    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=True)
        ctx = b.new_context(locale="th-TH")

        def on_res(res):
            nonlocal bundle_text
            u = res.url or ""
            if "queue-fe" in u and u.endswith(".js") and "index-" in u:
                try:
                    bundle_text = res.text()
                except Exception:
                    pass

        ctx.on("response", on_res)
        page = ctx.new_page()
        try:
            login(page, cfg); page.wait_for_timeout(1200)
            goto_tracking(page); page.wait_for_timeout(1500)
            apply_wa_filter(page, "", status_ids=["AP"]); page.wait_for_timeout(1500)
            rows = collect_all_wa_rows(page, log=lambda *a: None)
            ap = [r for r in rows if str(r.get("status") or "").upper() == "AP"]
            if not ap: return 1
            tgt = ap[0]
            page.goto(build_detail_url(tgt), wait_until="domcontentloaded", timeout=25_000)
            page.wait_for_timeout(2500)
            page.evaluate("() => { if(window.jQuery) jQuery('a[href=\"#tab_default_5\"]').tab('show'); }")
            page.wait_for_timeout(3000)
            iframe_src = page.evaluate("() => (document.querySelector('#link_appointment') || {}).src || ''")
            qp = ctx.new_page()
            qp.goto(iframe_src, wait_until="networkidle", timeout=30_000)
            qp.wait_for_timeout(4000)
        finally:
            ctx.close(); b.close()

    if not bundle_text:
        print("[!] bundle not captured"); return 1
    BUNDLE_PATH.write_text(bundle_text, encoding="utf-8")
    print(f"[+] saved {len(bundle_text):,} bytes → {BUNDLE_PATH}")

    print("\n=== keyword scan ===")
    for kw in ["round", "slot", "time-slot", "timeslot", "period",
               "appointment", "calendar", "book-time", "get-time", "make-booking"]:
        hits = [m.start() for m in re.finditer(kw, bundle_text, re.IGNORECASE)]
        if hits:
            print(f"\n'{kw}' — {len(hits)} hits:")
            for pos in hits[:5]:
                snip = bundle_text[max(0,pos-70):pos+110].replace("\n", " ")
                print(f"   …{snip}…")

    # Broader: any /vN/... path fragment
    print("\n=== all API path fragments (v1/v2 endpoints) ===")
    for pat in [r'"(/v\d/[a-zA-Z0-9\-_/]+)"',
                r'"(/api/v\d/[a-zA-Z0-9\-_/]+)"',
                r'"(/[a-zA-Z\-]+/api/v\d/[a-zA-Z0-9\-_/]+)"',
                r"'(/v\d/[a-zA-Z0-9\-_/]+)'"]:
        for m in set(re.findall(pat, bundle_text)):
            print(f"  · {m}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
