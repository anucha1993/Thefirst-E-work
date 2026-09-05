"""After opening modal, dump EVERY clickable element inside the modal container.
Focus: find how date cells are rendered.
"""
from __future__ import annotations
import json, os, sys, time
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import login, ROOT, goto_tracking, apply_wa_filter, collect_all_wa_rows, build_detail_url

OUT = ROOT / "_probe_modal_out.json"


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
                                  "url": u[:250], "body": body[:2000]})
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

            # STEP 1: dropdown + choose first
            print("\n=== select branch ===")
            qp.get_by_role("button", name="กรุณาเลือกสถานที่").first.click(timeout=10_000)
            qp.wait_for_timeout(2000)
            first_opt = qp.locator("li").first
            first_opt_text = first_opt.text_content()
            print(f"  branch: {first_opt_text}")
            first_opt.click(timeout=10_000)
            qp.wait_for_timeout(3000)

            # STEP 2: open calendar modal
            print("\n=== open calendar modal ===")
            qp.get_by_role("button", name="เพิ่มวันนัดหมาย").first.click(timeout=10_000)
            qp.wait_for_timeout(4000)
            qp.screenshot(path=str(ROOT / "_probe_modal_after_open.png"), full_page=True)

            # DUMP everything inside modal
            print("\n=== dumping modal DOM ===")
            info = qp.evaluate(r"""
            () => {
                const cl = t => (t||'').replace(/\s+/g,' ').trim();
                // find largest overlaid modal (fixed/absolute pos, high z-index)
                const all = [...document.querySelectorAll('div, section')];
                const modals = all.filter(el => {
                    const s = getComputedStyle(el);
                    return (s.position === 'fixed' || s.position === 'absolute')
                        && el.offsetParent
                        && el.offsetWidth > 300 && el.offsetHeight > 200
                        && parseInt(s.zIndex) > 10;
                });
                // sort by area desc; take biggest
                modals.sort((a,b) => (b.offsetWidth*b.offsetHeight) - (a.offsetWidth*a.offsetHeight));
                const modal = modals[0] || document.body;
                // enumerate ALL children with size + text
                const kids = [...modal.querySelectorAll('*')].filter(el =>
                    el.offsetParent
                    && el.offsetWidth > 0 && el.offsetHeight > 0
                );
                // group by tag
                const byTag = {};
                for (const el of kids) {
                    const tag = el.tagName;
                    byTag[tag] = (byTag[tag]||0) + 1;
                }
                // pick 'clickable-looking' small elements (potential date cells)
                const cells = kids.filter(el => {
                    const s = getComputedStyle(el);
                    const txt = cl(el.textContent);
                    const isNumericText = /^\d{1,2}$/.test(txt);
                    return isNumericText && el.offsetWidth < 100 && el.offsetHeight < 100;
                }).map(el => ({
                    tag: el.tagName,
                    text: cl(el.textContent),
                    cls: (el.className || '').toString().slice(0, 150),
                    id: el.id || '',
                    w: el.offsetWidth, h: el.offsetHeight,
                    cursor: getComputedStyle(el).cursor,
                    disabled: el.getAttribute('aria-disabled') === 'true' || el.hasAttribute('disabled'),
                    parentTag: el.parentElement ? el.parentElement.tagName : '',
                    parentCls: el.parentElement ? (el.parentElement.className||'').toString().slice(0,120) : '',
                }));

                // headings/labels in modal
                const heads = [...modal.querySelectorAll('h1,h2,h3,h4,h5,label,span')]
                    .filter(el => el.offsetParent && cl(el.textContent))
                    .map(el => cl(el.textContent))
                    .filter(t => t.length > 1 && t.length < 80)
                    .slice(0, 40);

                return {
                    modalFound: modal !== document.body,
                    modalCls: (modal.className||'').toString().slice(0, 150),
                    modalTag: modal.tagName,
                    modalSize: `${modal.offsetWidth}x${modal.offsetHeight}`,
                    byTag,
                    numericCells: cells,
                    heads,
                };
            }
            """)
            print(f"modal: {info['modalTag']}.{info['modalCls'][:60]} ({info['modalSize']})")
            print(f"tag distribution: {info['byTag']}")
            print(f"heads: {info['heads'][:20]}")
            print(f"\nnumeric cells ({len(info['numericCells'])}):")
            for c in info["numericCells"][:35]:
                print(f"  · {c['tag']} '{c['text']}' {c['w']}x{c['h']} cursor={c['cursor']} disabled={c['disabled']}")
                print(f"      cls={c['cls'][:100]}")
                print(f"      parent={c['parentTag']}.{c['parentCls'][:80]}")

            # try click first non-disabled numeric cell
            if info["numericCells"]:
                enabled = [c for c in info["numericCells"] if not c["disabled"]]
                if enabled:
                    tgt_cell = enabled[0]
                    print(f"\n=== click cell '{tgt_cell['text']}' ({tgt_cell['tag']}) ===")
                    # build precise selector
                    if tgt_cell["cls"]:
                        first_cls = tgt_cell["cls"].split()[0]
                        sel = f'{tgt_cell["tag"].lower()}.{first_cls}'
                    else:
                        sel = tgt_cell["tag"].lower()
                    print(f"  selector: {sel} :text('{tgt_cell['text']}')")
                    try:
                        loc = qp.locator(sel, has_text=tgt_cell["text"]).first
                        loc.click(timeout=8000)
                        print("  [+] clicked!")
                    except Exception as e:
                        print(f"  [!] click err: {e}")
                    print("\n  waiting 12s for API ...")
                    for _ in range(12):
                        qp.wait_for_timeout(1000)
                    qp.screenshot(path=str(ROOT / "_probe_modal_after_date.png"), full_page=True)

            uniq = set()
            for c in api_calls:
                if "/doe-booking/" in c["url"]:
                    ep = c["url"].split("/doe-booking/",1)[1].split("?",1)[0]
                    uniq.add(ep)
            print(f"\n=== TOTAL {len(api_calls)} calls, {len(uniq)} unique endpoints ===")
            for ep in sorted(uniq): print(f"  · /{ep}")

            OUT.write_text(json.dumps({
                "dom_info": info,
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
