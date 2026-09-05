"""Walk with SS status record (not yet booked) → find enabled date → capture time-slot API."""
from __future__ import annotations
import json, os, sys
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import login, ROOT, goto_tracking, apply_wa_filter, collect_all_wa_rows, build_detail_url

OUT = ROOT / "_probe_walk4_out.json"


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }
    api_calls = []

    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, args=["--start-maximized"])
        ctx = b.new_context(locale="th-TH", viewport={"width": 1600, "height": 1000})

        def on_res(res):
            u = res.url or ""
            if "queue-be" in u:
                try: body = res.text()
                except Exception: body = ""
                api_calls.append({"status": res.status, "method": res.request.method,
                                  "url": u[:250], "body": body[:3000]})
                mark = "🎯" if any(k in u.lower() for k in
                    ("timeslot","time-slot","period","slot","session","hour","schedule","reserve","book-time","get-time","booking-time")) else "  "
                print(f"  {mark} [{res.status}] {res.request.method} {u[u.find('/api/'):][:200]}")
        ctx.on("response", on_res)

        page = ctx.new_page()
        try:
            login(page, cfg); page.wait_for_timeout(1200)
            goto_tracking(page); page.wait_for_timeout(1500)
            apply_wa_filter(page, "", status_ids=["SS"]); page.wait_for_timeout(1500)
            rows = collect_all_wa_rows(page, log=lambda *a: None)
            ss = [r for r in rows if str(r.get("status") or "").upper() == "SS"]
            print(f"SS rows: {len(ss)}")
            if not ss:
                # try all status
                apply_wa_filter(page, ""); page.wait_for_timeout(1500)
                rows = collect_all_wa_rows(page, log=lambda *a: None)
                print(f"all rows: {len(rows)}, statuses: {sorted(set(str(r.get('status') or '').upper() for r in rows))}")
                return 1
            tgt = ss[0]
            print(f"target SS: {tgt.get('reqNo')}")
            page.goto(build_detail_url(tgt), wait_until="domcontentloaded", timeout=25_000)
            page.wait_for_timeout(2500)
            page.evaluate("() => { if(window.jQuery) jQuery('a[href=\"#tab_default_5\"]').tab('show'); }")
            page.wait_for_timeout(3000)
            iframe_src = page.evaluate("() => (document.querySelector('#link_appointment') || {}).src || ''")
            print(f"iframe: {iframe_src[:100]}...")

            if not iframe_src:
                print("[!] no iframe on SS record — this might be the wrong tab")
                # dump tabs
                tabs = page.evaluate("() => [...document.querySelectorAll('a[href*=tab]')].map(a=>({href:a.getAttribute('href'), text:a.textContent.trim()}))")
                print(f"tabs: {tabs}")
                return 1

            qp = ctx.new_page()
            qp.goto(iframe_src, wait_until="networkidle", timeout=30_000)
            qp.wait_for_timeout(3500)

            print("\n=== dropdown → first branch ===")
            qp.get_by_role("button", name="กรุณาเลือกสถานที่").first.click(timeout=10_000)
            qp.wait_for_timeout(2000)
            first_li = qp.locator("li").first
            print(f"  → {first_li.text_content()}")
            first_li.click(timeout=10_000)
            qp.wait_for_timeout(3500)

            print("\n=== open modal ===")
            qp.get_by_role("button", name="เพิ่มวันนัดหมาย").first.click(timeout=10_000)
            qp.wait_for_timeout(4000)

            # dump ALL cells with full cls and cursor
            print("\n=== full cell dump ===")
            all_cells = qp.evaluate(r"""
            () => [...document.querySelectorAll('.grid.grid-cols-7 > div')].map(el => {
                const s = getComputedStyle(el);
                const r = el.getBoundingClientRect();
                return {
                    text: (el.textContent||'').replace(/\s+/g,' ').trim().slice(0,10),
                    cls: (el.className||'').toString(),
                    cursor: s.cursor,
                    bg: s.backgroundColor,
                    cx: r.x + r.width/2, cy: r.y + r.height/2,
                };
            })
            """)
            print(f"total cells: {len(all_cells)}")
            enabled = [c for c in all_cells if c["cursor"] == "pointer" and c["text"].strip().isdigit()]
            print(f"enabled: {len(enabled)}")
            # sample first 3
            for c in all_cells[:6]:
                print(f"  '{c['text']}' cursor={c['cursor']} bg={c['bg']} cls={c['cls'][:150]}")
            if enabled:
                print(f"\nFIRST ENABLED: '{enabled[0]['text']}' cls={enabled[0]['cls'][:200]}")

            if enabled:
                tgt_cell = enabled[0]
                print(f"\n=== click date '{tgt_cell['text']}' ===")
                qp.mouse.click(tgt_cell["cx"], tgt_cell["cy"])
                print("  waiting 15s for API...")
                for _ in range(15): qp.wait_for_timeout(1000)
                qp.screenshot(path=str(ROOT / "_probe_walk4_after.png"), full_page=True)

            uniq = set()
            for c in api_calls:
                if "/doe-booking/" in c["url"]:
                    ep = c["url"].split("/doe-booking/",1)[1].split("?",1)[0]
                    uniq.add(ep)
            print(f"\n=== {len(api_calls)} calls, {len(uniq)} unique endpoints ===")
            for ep in sorted(uniq): print(f"  · /{ep}")

            OUT.write_text(json.dumps({
                "target": {"reqNo": tgt.get("reqNo"), "status": tgt.get("status")},
                "all_cells_sample": all_cells[:20],
                "enabled_count": len(enabled),
                "api_calls": api_calls,
                "unique_endpoints": sorted(uniq),
            }, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"\n[done] → {OUT}")
            try: input("\n Enter to close ...")
            except Exception: pass
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
