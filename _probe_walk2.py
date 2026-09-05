"""Walk with CCO (Chonburi = Normal-supporting branch) to trigger time-slot API."""
from __future__ import annotations
import json, os, sys
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import login, ROOT, goto_tracking, apply_wa_filter, collect_all_wa_rows, build_detail_url

OUT = ROOT / "_probe_walk2_out.json"

BRANCH_KEYWORD = "ชลบุรี"  # try Chonburi, known to have Normal capacity


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
                                  "url": u[:250], "body": body[:2500]})
                mark = "🎯" if any(k in u.lower() for k in
                    ("timeslot","time-slot","period","slot","session","hour","schedule","reserve","book-time","get-time")) else "  "
                print(f"  {mark} [{res.status}] {res.request.method} {u[u.find('/api/'):][:180]}")
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

            qp = ctx.new_page()
            qp.goto(iframe_src, wait_until="networkidle", timeout=30_000)
            qp.wait_for_timeout(3500)

            # STEP 1: dropdown → pick Chonburi
            print(f"\n=== select branch (keyword: {BRANCH_KEYWORD}) ===")
            qp.get_by_role("button", name="กรุณาเลือกสถานที่").first.click(timeout=10_000)
            qp.wait_for_timeout(2000)
            # find LI containing keyword
            li_match = qp.locator("li", has_text=BRANCH_KEYWORD).first
            branch_text = li_match.text_content()
            print(f"  found: {branch_text}")
            li_match.click(timeout=10_000)
            qp.wait_for_timeout(3500)

            # STEP 2: open modal
            print("\n=== open calendar modal ===")
            qp.get_by_role("button", name="เพิ่มวันนัดหมาย").first.click(timeout=10_000)
            qp.wait_for_timeout(4000)

            # STEP 3: enumerate cells, find ENABLED (cursor=pointer or NOT border-hint)
            print("\n=== find enabled date cell ===")
            enabled_dates = qp.evaluate(r"""
            () => {
                const cells = [...document.querySelectorAll('.grid.grid-cols-7 > div')];
                return cells.filter(el => {
                    const s = getComputedStyle(el);
                    const cls = (el.className||'').toString();
                    const txt = (el.textContent||'').replace(/\s+/g,' ').trim();
                    return s.cursor === 'pointer'
                        && /^\d{1,2}/.test(txt)
                        && !cls.includes('border-hint');
                }).map(el => ({
                    text: (el.textContent||'').replace(/\s+/g,' ').trim(),
                    cls: (el.className||'').toString().slice(0, 150),
                    rect: (r => ({x: r.x, y: r.y, w: r.width, h: r.height}))(el.getBoundingClientRect()),
                }));
            }
            """)
            print(f"  enabled cells: {len(enabled_dates)}")
            for c in enabled_dates[:5]:
                print(f"    · '{c['text']}' rect={c['rect']} cls={c['cls'][:100]}")

            if not enabled_dates:
                # try navigating to next month
                print("\n[!] no enabled dates in current month — trying next month via > chevron")
                # dump all buttons in modal
                btns = qp.evaluate(r"""
                    () => [...document.querySelectorAll('button, [role="button"]')]
                        .filter(el => el.offsetParent)
                        .map(el => ({
                            text: (el.textContent||'').replace(/\s+/g,' ').trim().slice(0,50),
                            aria: el.getAttribute('aria-label') || '',
                            cls: (el.className||'').toString().slice(0,100),
                        }))
                """)
                print(f"  visible buttons ({len(btns)}):")
                for b_ in btns[:20]:
                    print(f"    · '{b_['text']}' aria='{b_['aria']}' cls={b_['cls'][:80]}")
            else:
                # STEP 4: click at exact rect center via mouse
                target = enabled_dates[0]
                cx = target["rect"]["x"] + target["rect"]["w"] / 2
                cy = target["rect"]["y"] + target["rect"]["h"] / 2
                print(f"\n=== click enabled date '{target['text']}' at ({cx:.0f},{cy:.0f}) ===")
                qp.mouse.click(cx, cy)
                print("  [+] clicked. Waiting 15s for API...")
                for i in range(15):
                    qp.wait_for_timeout(1000)
                qp.screenshot(path=str(ROOT / "_probe_walk2_after_date.png"), full_page=True)

            uniq = set()
            for c in api_calls:
                if "/doe-booking/" in c["url"]:
                    ep = c["url"].split("/doe-booking/",1)[1].split("?",1)[0]
                    uniq.add(ep)
            print(f"\n=== {len(api_calls)} calls, {len(uniq)} unique endpoints ===")
            for ep in sorted(uniq): print(f"  · /{ep}")

            OUT.write_text(json.dumps({
                "branch": branch_text,
                "enabled_dates": enabled_dates,
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
