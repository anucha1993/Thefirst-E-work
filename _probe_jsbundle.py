"""Load iframe URL, capture all JS bundles, grep for time-slot API URL pattern."""
from __future__ import annotations
import json, os, re, sys
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import login, ROOT, goto_tracking, apply_wa_filter, collect_all_wa_rows, build_detail_url

OUT = ROOT / "_probe_jsbundle_out.json"


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }
    js_bundles = {}  # url -> body

    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, args=["--start-maximized"])
        ctx = b.new_context(locale="th-TH", viewport={"width": 1600, "height": 1000})

        def on_res(res):
            u = res.url or ""
            if "queue-fe" in u and (u.endswith(".js") or "/_next/static/" in u):
                try:
                    body = res.text()
                    js_bundles[u] = body
                except Exception: pass

        ctx.on("response", on_res)

        page = ctx.new_page()
        try:
            login(page, cfg); page.wait_for_timeout(1200)
            goto_tracking(page); page.wait_for_timeout(1500)
            apply_wa_filter(page, "", status_ids=["AP"]); page.wait_for_timeout(1500)
            rows = collect_all_wa_rows(page, log=lambda *a: None)
            ap = [r for r in rows if str(r.get("status") or "").upper() == "AP"]
            if not ap: print("[!] no AP"); return 1
            tgt = ap[0]
            print(f"target: {tgt.get('reqNo')}")
            page.goto(build_detail_url(tgt), wait_until="domcontentloaded", timeout=25_000)
            page.wait_for_timeout(2500)
            page.evaluate("() => { if(window.jQuery) jQuery('a[href=\"#tab_default_5\"]').tab('show'); }")
            page.wait_for_timeout(3000)
            iframe_src = page.evaluate("() => (document.querySelector('#link_appointment') || {}).src || ''")
            print(f"iframe: {iframe_src[:100]}...")

            qp = ctx.new_page()
            qp.goto(iframe_src, wait_until="networkidle", timeout=30_000)
            qp.wait_for_timeout(5000)

            print(f"\n=== captured {len(js_bundles)} JS bundles ===")
            for u in list(js_bundles.keys())[:20]:
                print(f"  · {u[:120]}  ({len(js_bundles[u])} bytes)")

            # scan for API URL patterns
            print("\n=== scanning bundles for API URL patterns ===")
            # combined text
            combined = "\n".join(js_bundles.values())
            print(f"combined size: {len(combined):,} bytes")

            # look for /api/... paths
            patterns = [
                r'"(/doe-booking/api/v\d/[^"\s]+)"',
                r"'(/doe-booking/api/v\d/[^'\s]+)'",
                r'`(/doe-booking/api/v\d/[^`\s]+)`',
            ]
            found = set()
            for p in patterns:
                for m in re.finditer(p, combined):
                    found.add(m.group(1))
            print(f"\nFound {len(found)} distinct API paths in bundles:")
            for path in sorted(found):
                # highlight if related to time
                marker = "🎯" if any(k in path.lower() for k in
                    ("time","slot","period","hour","schedule","session","reserve","book/create","period","hour")) else "  "
                print(f"  {marker} {path}")

            # Additional: search for keywords in JS
            print("\n=== keyword hits in combined bundles ===")
            for kw in ["timeslot","time_slot","time-slot","period","hour","session","booking-time","get-time","slot"]:
                matches = [m.start() for m in re.finditer(kw, combined, re.IGNORECASE)]
                if matches:
                    print(f"  · '{kw}': {len(matches)} hits")
                    # show context of first 3 hits
                    for pos in matches[:3]:
                        snip = combined[max(0,pos-60):pos+80].replace("\n"," ")
                        print(f"     …{snip}…")

            OUT.write_text(json.dumps({
                "iframe": iframe_src,
                "js_urls": list(js_bundles.keys()),
                "api_paths_in_bundles": sorted(found),
            }, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"\n[done] → {OUT}")
            try: input("\n Enter to close ...")
            except Exception: pass
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
