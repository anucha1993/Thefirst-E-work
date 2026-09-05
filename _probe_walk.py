"""Auto-walk booking flow to discover time-slot API.
Steps:
 1. Click 'กรุณาเลือกสถานที่' button (dropdown)
 2. Wait for option list → click first option
 3. Click 'เพิ่มวันนัดหมาย' → calendar popup
 4. Click first enabled date cell in calendar
 5. Watch for time-slot API
All via Playwright locator.click() (fires proper events)
"""
from __future__ import annotations
import json, os, sys, time
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import login, ROOT, goto_tracking, apply_wa_filter, collect_all_wa_rows, build_detail_url

OUT = ROOT / "_probe_walk_out.json"


def snap(page, name):
    try: page.screenshot(path=str(ROOT / f"_probe_walk_{name}.png"), full_page=True)
    except Exception: pass


def dump_state(page):
    """Return brief state: visible buttons + LI + inputs"""
    return page.evaluate(r"""
    () => {
        const cl = t => (t||'').replace(/\s+/g,' ').trim().slice(0,80);
        return {
            buttons: [...document.querySelectorAll('button')].filter(el=>el.offsetParent)
                .map(el=>({text:cl(el.textContent), cls:(el.className||'').slice(0,80)})),
            list: [...document.querySelectorAll('li, [role=option], [role=menuitem]')].filter(el=>el.offsetParent)
                .map(el=>({tag:el.tagName, text:cl(el.textContent), cls:(el.className||'').slice(0,80)})).slice(0,30),
            calendar_cells: [...document.querySelectorAll('[class*="calendar"] button, [class*="date"] button, td button, .rdrDay, .react-datepicker__day')]
                .filter(el=>el.offsetParent).map(el=>({
                  text:cl(el.textContent),
                  cls:(el.className||'').toString().slice(0,120),
                  disabled: el.disabled || el.getAttribute('aria-disabled')==='true',
                })).slice(0,50),
        };
    }
    """)


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }
    api_calls = []
    events = []

    def logevt(msg):
        print(msg)
        events.append({"t": time.time(), "msg": msg})

    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, args=["--start-maximized"])
        ctx = b.new_context(locale="th-TH", viewport={"width": 1600, "height": 1000})

        def on_res(res):
            u = res.url or ""
            if "queue-be" in u:
                try: body = res.text()
                except Exception: body = ""
                rec = {"status": res.status, "method": res.request.method,
                       "url": u[:250], "body": body[:1500], "t": time.time()}
                api_calls.append(rec)
                # live print
                mark = "🎯" if any(k in u.lower() for k in
                                   ("timeslot","time-slot","period","slot","session","hour","schedule","reserve","book/create")) else "  "
                print(f"  {mark} [{res.status}] {res.request.method} {u[u.find('/api/'):][:180]}")
        ctx.on("response", on_res)

        page = ctx.new_page()
        try:
            login(page, cfg); page.wait_for_timeout(1200)
            goto_tracking(page); page.wait_for_timeout(1500)
            apply_wa_filter(page, "", status_ids=["AP"]); page.wait_for_timeout(1500)
            rows = collect_all_wa_rows(page, log=lambda *a: None)
            ap = [r for r in rows if str(r.get("status") or "").upper() == "AP"]
            if not ap:
                logevt("[!] no AP"); return 1
            tgt = ap[0]
            logevt(f"target: {tgt.get('reqNo')}")
            page.goto(build_detail_url(tgt), wait_until="domcontentloaded", timeout=25_000)
            page.wait_for_timeout(2500)
            page.evaluate("() => { if(window.jQuery) jQuery('a[href=\"#tab_default_5\"]').tab('show'); }")
            page.wait_for_timeout(3000)
            iframe_src = page.evaluate("() => (document.querySelector('#link_appointment') || {}).src || ''")
            logevt(f"iframe: {iframe_src[:100]}...")

            qp = ctx.new_page()
            qp.goto(iframe_src, wait_until="networkidle", timeout=30_000)
            qp.wait_for_timeout(3500)
            snap(qp, "01_loaded")

            # STEP 1: click dropdown button (button with text 'กรุณาเลือกสถานที่')
            logevt("\n=== STEP 1: click dropdown 'กรุณาเลือกสถานที่' ===")
            dropdown = qp.get_by_role("button", name="กรุณาเลือกสถานที่").first
            dropdown.click(timeout=10_000)
            qp.wait_for_timeout(2000)
            snap(qp, "02_dropdown_open")

            st1 = dump_state(qp)
            logevt(f"  visible LI/option: {len(st1['list'])}")
            for x in st1["list"][:8]:
                logevt(f"    · {x['tag']} {x['text']}")

            # STEP 2: click first option
            if not st1["list"]:
                logevt("[!] no options after dropdown click — try different selector")
                # try P-Select style
                opts = qp.locator("li, [role=option]").all()
                logevt(f"  fallback opts: {len(opts)}")
                return 1

            logevt("\n=== STEP 2: click first option ===")
            first_opt_text = st1["list"][0]["text"]
            logevt(f"  → clicking '{first_opt_text}'")
            # use locator with exact text
            opt = qp.locator("li, [role=option]").filter(has_text=first_opt_text).first
            opt.click(timeout=10_000)
            qp.wait_for_timeout(3000)
            snap(qp, "03_option_selected")

            # STEP 3: click 'เพิ่มวันนัดหมาย'
            logevt("\n=== STEP 3: click 'เพิ่มวันนัดหมาย' ===")
            add_btn = qp.get_by_role("button", name="เพิ่มวันนัดหมาย").first
            add_btn.click(timeout=10_000)
            qp.wait_for_timeout(4000)
            snap(qp, "04_calendar_open")

            st3 = dump_state(qp)
            logevt(f"  visible calendar cells: {len(st3['calendar_cells'])}")
            enabled = [c for c in st3["calendar_cells"] if not c.get("disabled")]
            logevt(f"  enabled cells: {len(enabled)}")
            for c in enabled[:15]:
                logevt(f"    · '{c['text']}' cls={c['cls'][:80]}")

            # STEP 4: click first ENABLED date
            if not enabled:
                logevt("[!] no enabled date cells — DOM may differ")
            else:
                logevt("\n=== STEP 4: click first enabled date cell ===")
                target_text = enabled[0]["text"]
                target_cls_snip = enabled[0]["cls"].split()[0] if enabled[0]["cls"] else ""
                logevt(f"  → click date '{target_text}'")
                # try clicking by text within button
                try:
                    day_btn = qp.locator("button:not([disabled])").filter(has_text=target_text).first
                    day_btn.click(timeout=8000)
                    logevt("  [+] clicked")
                except Exception as e:
                    logevt(f"  [!] click err: {e}")

                # wait long for time-slot API
                logevt("\n=== waiting 15s for time-slot API ===")
                for i in range(15):
                    qp.wait_for_timeout(1000)
                snap(qp, "05_after_date_click")

            # STEP 5: dump final state
            logevt("\n=== FINAL DOM STATE ===")
            st_final = dump_state(qp)
            logevt(f"  visible buttons: {len(st_final['buttons'])}")
            for x in st_final["buttons"][:20]:
                logevt(f"    · '{x['text']}'")

            # Summarize APIs
            logevt(f"\n=== TOTAL API CALLS: {len(api_calls)} ===")
            uniq_ep = set()
            for c in api_calls:
                if "/doe-booking/" in c["url"]:
                    ep = c["url"].split("/doe-booking/",1)[1].split("?",1)[0]
                    uniq_ep.add(ep)
            logevt(f"unique endpoints ({len(uniq_ep)}):")
            for ep in sorted(uniq_ep):
                logevt(f"  · /{ep}")

            OUT.write_text(json.dumps({
                "events": events,
                "api_calls": api_calls,
                "unique_endpoints": sorted(uniq_ep),
                "final_state": st_final,
            }, ensure_ascii=False, indent=2), encoding="utf-8")
            logevt(f"\n[done] → {OUT}")

            try: input("\n Enter to close ...")
            except Exception: pass
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
