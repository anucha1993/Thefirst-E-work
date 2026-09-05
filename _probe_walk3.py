"""Walk with CBI Chonburi, change month to Sep/Oct via SELECT, then click enabled date."""
from __future__ import annotations
import json, os, sys
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import login, ROOT, goto_tracking, apply_wa_filter, collect_all_wa_rows, build_detail_url

OUT = ROOT / "_probe_walk3_out.json"
BRANCH_KEYWORD = "ชลบุรี"


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

            print(f"\n=== select {BRANCH_KEYWORD} ===")
            qp.get_by_role("button", name="กรุณาเลือกสถานที่").first.click(timeout=10_000)
            qp.wait_for_timeout(2000)
            li_match = qp.locator("li", has_text=BRANCH_KEYWORD).first
            print(f"  → {li_match.text_content()}")
            li_match.click(timeout=10_000)
            qp.wait_for_timeout(3500)

            print("\n=== open modal ===")
            qp.get_by_role("button", name="เพิ่มวันนัดหมาย").first.click(timeout=10_000)
            qp.wait_for_timeout(4000)

            # inspect SELECTs
            print("\n=== inspect SELECTs (month/year) ===")
            selects_info = qp.evaluate(r"""
                () => [...document.querySelectorAll('select')].filter(el=>el.offsetParent).map(el => ({
                    name: el.name||'',
                    value: el.value,
                    options: [...el.options].map(o => ({v: o.value, t: o.textContent.trim()})),
                }))
            """)
            for i, s in enumerate(selects_info):
                print(f"  select#{i} name='{s['name']}' value={s['value']}")
                for o in s["options"][:14]:
                    print(f"    · {o['v']} = '{o['t']}'")

            # Try to change month/year to find an available month
            # We'll iterate months
            for try_month in ["9", "10", "11", "12"]:
                print(f"\n=== try month={try_month} ===")
                # find month select by matching options containing '9','10','11','12'
                changed = qp.evaluate(f"""
                (target) => {{
                    const sels = [...document.querySelectorAll('select')].filter(el=>el.offsetParent);
                    for (const s of sels) {{
                        const vals = [...s.options].map(o => o.value);
                        if (vals.includes(target) && vals.length <= 13) {{
                            // month has 12 options
                            s.value = target;
                            s.dispatchEvent(new Event('change', {{bubbles:true}}));
                            return {{ok:true, sel:s.name||'unnamed'}};
                        }}
                    }}
                    return {{ok:false}};
                }}
                """, try_month)
                print(f"  month change: {changed}")
                if not changed.get("ok"):
                    continue
                qp.wait_for_timeout(3500)  # wait for API + rerender

                enabled = qp.evaluate(r"""
                () => {
                    const cells = [...document.querySelectorAll('.grid.grid-cols-7 > div')];
                    return cells.filter(el => {
                        const s = getComputedStyle(el);
                        const cls = (el.className||'').toString();
                        const txt = (el.textContent||'').replace(/\s+/g,' ').trim();
                        return s.cursor === 'pointer'
                            && /^\d{1,2}$/.test(txt)
                            && !cls.includes('border-hint');
                    }).map(el => {
                        const r = el.getBoundingClientRect();
                        return {
                            text: (el.textContent||'').replace(/\s+/g,' ').trim(),
                            cx: r.x + r.width/2, cy: r.y + r.height/2,
                            cls: (el.className||'').toString().slice(0, 120),
                        };
                    });
                }
                """)
                print(f"  enabled cells: {len(enabled)}")
                if enabled:
                    print(f"  first: '{enabled[0]['text']}' at ({enabled[0]['cx']:.0f},{enabled[0]['cy']:.0f})")
                    print(f"         cls={enabled[0]['cls']}")
                    # scroll into view then click
                    qp.mouse.click(enabled[0]["cx"], enabled[0]["cy"])
                    print(f"  [+] clicked date {enabled[0]['text']} in month {try_month}")
                    print("  waiting 15s for API...")
                    for _ in range(15):
                        qp.wait_for_timeout(1000)
                    qp.screenshot(path=str(ROOT / "_probe_walk3_after_click.png"), full_page=True)
                    break
            else:
                print("[!] no month found with enabled dates")

            uniq = set()
            for c in api_calls:
                if "/doe-booking/" in c["url"]:
                    ep = c["url"].split("/doe-booking/",1)[1].split("?",1)[0]
                    uniq.add(ep)
            print(f"\n=== {len(api_calls)} calls, {len(uniq)} unique endpoints ===")
            for ep in sorted(uniq): print(f"  · /{ep}")

            OUT.write_text(json.dumps({
                "selects_info": selects_info,
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
