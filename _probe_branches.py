"""Focused probe: หา branch code ใน LI ของ dropdown สาขา
+ ลอง /branch/by-codes ด้วย body format ต่าง ๆ
"""
from __future__ import annotations
import json, os, sys
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import login, ROOT, goto_tracking, apply_wa_filter, collect_all_wa_rows, build_detail_url

OUT_JSON = ROOT / "_probe_branches_out.json"


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }

    captured = []

    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, args=["--start-maximized"])
        ctx = b.new_context(locale="th-TH", viewport={"width": 1600, "height": 1000})

        def on_req(req):
            u = req.url or ""
            if "queue-be" in u and "branch" in u:
                captured.append({
                    "kind": "req",
                    "method": req.method,
                    "url": u[:250],
                    "post": req.post_data or "",
                    "headers": {k: v for k, v in dict(req.headers).items()
                                if k.lower() in ("authorization", "content-type", "origin", "referer")},
                })

        def on_res(res):
            u = res.url or ""
            if "queue-be" in u and "branch" in u:
                try: text = res.text()
                except Exception: text = ""
                captured.append({"kind": "res", "status": res.status,
                                 "url": u[:250], "body": text[:6000]})
        ctx.on("request", on_req)
        ctx.on("response", on_res)

        page = ctx.new_page()
        try:
            login(page, cfg)
            page.wait_for_timeout(1500)
            goto_tracking(page)
            apply_wa_filter(page, "", status_ids=["AP"])
            rows = collect_all_wa_rows(page, log=lambda *a: None)
            ap_rows = [r for r in rows if str(r.get("status") or "").upper() == "AP"]
            target = ap_rows[0]
            url = build_detail_url(target)
            page.goto(url, wait_until="domcontentloaded", timeout=25_000)
            page.wait_for_timeout(3000)
            page.evaluate("() => { if(window.jQuery) jQuery('a[href=\"#tab_default_5\"]').tab('show'); }")
            page.wait_for_timeout(3500)
            iframe_src = page.evaluate("() => (document.querySelector('#link_appointment') || {}).src || ''")

            qp = ctx.new_page()
            captured.clear()
            qp.goto(iframe_src, wait_until="networkidle", timeout=30_000)
            qp.wait_for_timeout(6000)

            # ดู branch API ที่ FE ยิงตอน page load
            print(f"[!] captured branch API calls: {len(captured)}")
            for c in captured:
                if c["kind"] == "req":
                    print(f"\n  REQ {c['method']} {c['url']}")
                    print(f"    headers: {c['headers']}")
                    if c["post"]:
                        print(f"    body: {c['post'][:600]}")
                elif c["kind"] == "res":
                    print(f"  RES {c['status']}  body_len={len(c['body'])}")
                    print(f"    body preview: {c['body'][:400]}")

            # ลอง click dropdown → ดู LI attributes
            print("\n[i] click dropdown...")
            captured.clear()
            qp.evaluate(r"""() => {
                const vis = el => !!(el && el.offsetParent !== null);
                const el = Array.from(document.querySelectorAll('div, [role="combobox"]'))
                    .filter(vis)
                    .filter(e => /^กรุณาเลือกสถานที่$/.test((e.textContent||'').trim()))[0];
                if (el) el.click();
            }""")
            qp.wait_for_timeout(2500)

            # dump ALL LI attributes
            li_info = qp.evaluate(r"""() => {
                const vis = el => !!(el && el.offsetParent !== null);
                const lis = Array.from(document.querySelectorAll('li')).filter(vis);
                return lis.map((li, idx) => {
                    const attrs = {};
                    for (const a of li.attributes) attrs[a.name] = a.value;
                    return {
                        idx,
                        text: (li.textContent||'').replace(/\s+/g,' ').trim().slice(0, 100),
                        attrs,
                        // React props via key trick
                        reactProps: Object.keys(li).filter(k => k.startsWith('__react')),
                    };
                });
            }""")
            print(f"\n[LI elements] {len(li_info)}:")
            for li in li_info[:10]:
                print(f"  [{li['idx']}] {li['text'][:60]!r}")
                if li['attrs']:
                    print(f"    attrs: {li['attrs']}")

            # ลอง SELECT ปกติด้วย
            print("\n[i] check <select> and <option>...")
            sel_info = qp.evaluate(r"""() => {
                const opts = Array.from(document.querySelectorAll('option')).map(o => ({
                    text: (o.textContent||'').trim().slice(0, 100),
                    value: o.value,
                }));
                return {optionCount: opts.length, options: opts.slice(0, 20)};
            }""")
            print(f"  options: {sel_info}")

            # Try clicking first branch LI and see what URL /calendar loads
            print("\n[i] click first branch LI + capture calendar URL...")
            captured.clear()
            qp.evaluate(r"""() => {
                const vis = el => !!(el && el.offsetParent !== null);
                const lis = Array.from(document.querySelectorAll('li')).filter(vis)
                    .filter(x => /ศูนย์บริการ|ศูนย์แรกรับ|หน่วยบริการ/.test((x.textContent||'')));
                if (lis[0]) lis[0].click();
            }""")
            qp.wait_for_timeout(3000)
            # click 'เพิ่มวันนัดหมาย' to trigger calendar API
            qp.evaluate(r"""() => {
                const vis = el => !!(el && el.offsetParent !== null);
                const b = Array.from(document.querySelectorAll('button')).filter(vis)
                    .find(x => /^\s*เพิ่มวันนัดหมาย\s*$/.test((x.textContent||'').trim()));
                if (b) b.click();
            }""")
            qp.wait_for_timeout(4000)

            # capture ที่ตามมา — calendar API มี branch_code_id ใน URL
            all_after = []
            def on_req_after(req):
                u = req.url or ""
                if "queue-be" in u:
                    all_after.append(u[:250])
            qp.on("request", on_req_after)
            qp.wait_for_timeout(2000)
            print(f"\n[URLs captured after selecting branch]:")
            for c in captured[:15]:
                if c["kind"] == "req":
                    print(f"  {c['method']} {c['url']}")

            OUT_JSON.write_text(json.dumps({
                "li_info": li_info,
                "select_info": sel_info,
                "captured_all": captured,
            }, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"\n[done] → {OUT_JSON}")

            input("\n Enter to close ...")
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
