"""Dump DOM structure of booking iframe to figure out selectors.
- Opens iframe URL
- After load: capture all buttons, dropdown-like elements, and their text
- Then click 'เพิ่มวันนัดหมาย' via locator + wait & re-dump
"""
from __future__ import annotations
import json, os, sys
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import login, ROOT, goto_tracking, apply_wa_filter, collect_all_wa_rows, build_detail_url

OUT = ROOT / "_probe_dom_booking_out.json"


def dump(page, stage):
    js = r"""
    () => {
        const clean = t => (t||'').replace(/\s+/g,' ').trim().slice(0, 100);
        const btns = [...document.querySelectorAll('button')].map((el,i) => ({
            i, tag: el.tagName, text: clean(el.textContent),
            id: el.id || '', cls: (el.className || '').toString().slice(0, 120),
            disabled: el.disabled, visible: !!el.offsetParent,
        })).filter(x => x.visible);
        const inputs = [...document.querySelectorAll('input, select, [role="combobox"], [role="listbox"], [role="button"], .p-dropdown, .p-select')]
          .slice(0, 40).map((el,i) => ({
            i, tag: el.tagName, role: el.getAttribute('role') || '',
            id: el.id || '', name: el.name || '',
            placeholder: el.placeholder || '',
            text: clean(el.textContent),
            cls: (el.className || '').toString().slice(0, 200),
            visible: !!el.offsetParent,
        })).filter(x => x.visible);
        // headings / labels
        const labels = [...document.querySelectorAll('h1,h2,h3,h4,label')]
          .map(el => clean(el.textContent)).filter(t => t.length > 0).slice(0, 30);
        return { url: location.href, buttons: btns, inputs, labels };
    }
    """
    d = page.evaluate(js)
    d["stage"] = stage
    return d


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }
    all_captured = []
    stages = []

    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, args=["--start-maximized"])
        ctx = b.new_context(locale="th-TH", viewport={"width": 1600, "height": 1000})

        def on_res(res):
            u = res.url or ""
            if "queue-be" in u:
                try: body = res.text()
                except Exception: body = ""
                all_captured.append({
                    "status": res.status, "method": res.request.method,
                    "url": u[:280], "body": body[:1500],
                })
        ctx.on("response", on_res)

        page = ctx.new_page()
        try:
            login(page, cfg); page.wait_for_timeout(1200)
            goto_tracking(page); page.wait_for_timeout(1500)
            apply_wa_filter(page, "", status_ids=["AP"]); page.wait_for_timeout(1500)
            rows = collect_all_wa_rows(page, log=lambda *a: None)
            ap = [r for r in rows if str(r.get("status") or "").upper() == "AP"]
            if not ap:
                print("[!] no AP"); return 1
            target = ap[0]
            url = build_detail_url(target)
            print(f"target: {target.get('reqNo')}")
            page.goto(url, wait_until="domcontentloaded", timeout=25_000)
            page.wait_for_timeout(2500)
            page.evaluate("() => { if(window.jQuery) jQuery('a[href=\"#tab_default_5\"]').tab('show'); }")
            page.wait_for_timeout(3000)
            iframe_src = page.evaluate("() => (document.querySelector('#link_appointment') || {}).src || ''")
            print(f"iframe: {iframe_src[:120]}")

            qp = ctx.new_page()
            qp.goto(iframe_src, wait_until="networkidle", timeout=30_000)
            qp.wait_for_timeout(4000)

            # STAGE 1 — initial load
            print("\n=== STAGE 1: initial load ===")
            s1 = dump(qp, "initial")
            stages.append(s1)
            print(f"buttons: {len(s1['buttons'])}, inputs: {len(s1['inputs'])}")
            print("labels:", s1["labels"][:10])
            print("first buttons:")
            for b_ in s1["buttons"][:12]:
                print(f"  · [{b_['i']}] {b_['text'][:60]}  (cls={b_['cls'][:60]})")
            print("first inputs:")
            for i_ in s1["inputs"][:12]:
                print(f"  · [{i_['i']}] {i_['tag']} role={i_['role']} text={i_['text'][:60]}  (cls={i_['cls'][:80]})")

            qp.screenshot(path=str(ROOT / "_probe_dom_booking_s1.png"), full_page=True)

            # STAGE 2 — try to find dropdown & open
            # look for text with 'สถานที่' or 'เลือกสถานที่'
            print("\n=== STAGE 2: try click on 'สถานที่' area ===")
            candidates = qp.locator("text=/สถานที่|เลือกสถานที่|กรุณาเลือกสถานที่/").all()
            print(f"candidates: {len(candidates)}")
            for c in candidates[:5]:
                try:
                    txt = c.text_content()
                    print(f"  · {(txt or '')[:80]}")
                except Exception: pass

            # try clicking last matching (should be the input row)
            try:
                if candidates:
                    candidates[-1].click(timeout=5000)
                    print("[+] clicked candidate")
                    qp.wait_for_timeout(2000)
            except Exception as e:
                print(f"[!] click err: {e}")

            s2 = dump(qp, "after-dropdown-click")
            stages.append(s2)
            qp.screenshot(path=str(ROOT / "_probe_dom_booking_s2.png"), full_page=True)
            # any new listbox/options?
            new_items = [i for i in s2["inputs"] if "listbox" in i["role"].lower() or "option" in i["cls"].lower()]
            print(f"new listbox/option elems: {len(new_items)}")
            for x in new_items[:20]:
                print(f"  · {x['tag']} role={x['role']} text={x['text'][:60]}")

            # dump also LI elements
            li_dump = qp.evaluate(r"""
                () => [...document.querySelectorAll('li, [role=option], .p-dropdown-item')]
                        .filter(el => el.offsetParent)
                        .slice(0, 30)
                        .map(el => ({
                          tag: el.tagName,
                          text: (el.textContent||'').replace(/\s+/g,' ').trim().slice(0,80),
                          cls: (el.className||'').toString().slice(0,120),
                        }))
            """)
            print(f"\nvisible LI/option ({len(li_dump)}):")
            for x in li_dump[:15]:
                print(f"  · {x['tag']} text={x['text']}  (cls={x['cls'][:60]})")

            OUT.write_text(json.dumps({
                "stages": stages,
                "li_after_click": li_dump,
                "api_calls": all_captured,
            }, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"\n[done] → {OUT}")
            print("\n Enter to close ...")
            try: input()
            except Exception: pass
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
