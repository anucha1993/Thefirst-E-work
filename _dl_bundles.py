"""Load booking iframe via Playwright, capture ALL JS bundles, save each to disk."""
from __future__ import annotations
import os, sys, re
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import login, ROOT, goto_tracking, apply_wa_filter, collect_all_wa_rows, build_detail_url

BUNDLE_DIR = ROOT / "_probe_bundles"
BUNDLE_DIR.mkdir(exist_ok=True)


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }
    saved = []

    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=True)  # can be headless — we just want files
        ctx = b.new_context(locale="th-TH")

        def on_res(res):
            u = res.url or ""
            if "queue-fe" in u and ("index-" in u and ".js" in u or "/_next/static/" in u or u.endswith(".js")):
                try:
                    body = res.text()
                    fname = re.sub(r"[^A-Za-z0-9._\-]", "_", u.split("/")[-1].split("?")[0])
                    (BUNDLE_DIR / fname).write_text(body, encoding="utf-8")
                    saved.append((u, len(body)))
                except Exception as e:
                    print(f"  save err: {e}")

        ctx.on("response", on_res)

        page = ctx.new_page()
        try:
            login(page, cfg); page.wait_for_timeout(1200)
            goto_tracking(page); page.wait_for_timeout(1500)
            apply_wa_filter(page, "", status_ids=["AP"]); page.wait_for_timeout(1500)
            rows = collect_all_wa_rows(page, log=lambda *a: None)
            ap = [r for r in rows if str(r.get("status") or "").upper() == "AP"]
            tgt = ap[0]
            page.goto(build_detail_url(tgt), wait_until="domcontentloaded", timeout=25_000)
            page.wait_for_timeout(2500)
            page.evaluate("() => { if(window.jQuery) jQuery('a[href=\"#tab_default_5\"]').tab('show'); }")
            page.wait_for_timeout(3000)
            iframe_src = page.evaluate("() => (document.querySelector('#link_appointment') || {}).src || ''")

            qp = ctx.new_page()
            qp.goto(iframe_src, wait_until="networkidle", timeout=30_000)
            qp.wait_for_timeout(4000)

            print(f"\nsaved {len(saved)} bundles → {BUNDLE_DIR}")
            for u, n in saved:
                print(f"  {n:>10} B  {u}")
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
