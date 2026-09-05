"""Probe: เดินตาม UI flow ครบ — เลือกสถานที่ + วัน → capture time-slot API
"""
from __future__ import annotations
import json, os, sys
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import login, ROOT, goto_tracking, apply_wa_filter, collect_all_wa_rows, build_detail_url

OUT_JSON = ROOT / "_probe_timeslot2_out.json"
SHOTS = ROOT / "screenshots" / "timeslot2"
SHOTS.mkdir(parents=True, exist_ok=True)


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

        def on_response(res):
            u = res.url or ""
            if "queue-be" in u:
                try: text = res.text()
                except Exception: text = ""
                captured.append({"status": res.status, "url": u[:250],
                                 "body_first": text[:2000]})
        ctx.on("response", on_response)

        page = ctx.new_page()
        try:
            login(page, cfg)
            page.wait_for_timeout(1500)
            goto_tracking(page)
            apply_wa_filter(page, "", status_ids=["AP"])
            rows = collect_all_wa_rows(page, log=lambda *a: None)
            ap_rows = [r for r in rows if str(r.get("status") or "").upper() == "AP"]
            if not ap_rows:
                print("[!] no AP"); return 1
            target = ap_rows[0]
            url = build_detail_url(target)
            print(f"target: {target.get('reqNo')}")
            page.goto(url, wait_until="domcontentloaded", timeout=25_000)
            page.wait_for_timeout(3000)
            page.evaluate("() => { if(window.jQuery) jQuery('a[href=\"#tab_default_5\"]').tab('show'); }")
            page.wait_for_timeout(3500)

            iframe_src = page.evaluate("() => (document.querySelector('#link_appointment') || {}).src || ''")

            qp = ctx.new_page()
            captured.clear()
            qp.goto(iframe_src, wait_until="networkidle", timeout=30_000)
            qp.wait_for_timeout(6000)

            qp.screenshot(path=str(SHOTS / "01_form_view.png"), full_page=True)

            # ขั้น 1: dump elements ก่อน — หา dropdown สถานที่ + ปุ่ม 'เพิ่มวันนัดหมาย'
            elements = qp.evaluate(r"""() => {
                const vis = el => !!(el && el.offsetParent !== null);
                const txt = e => (e.textContent||'').replace(/\s+/g,' ').trim();
                return {
                    // dropdown ที่มี text 'กรุณาเลือกสถานที่'
                    dropdowns: Array.from(document.querySelectorAll('button, [role="combobox"], div'))
                        .filter(vis)
                        .filter(e => /กรุณาเลือกสถานที่|สถานที่/.test(txt(e)) && txt(e).length < 100)
                        .slice(0, 5).map(e => ({tag: e.tagName, text: txt(e).slice(0, 80),
                                                cls: (e.className||'').toString().slice(0, 80)})),
                    addDayBtn: Array.from(document.querySelectorAll('button'))
                        .filter(vis)
                        .filter(b => /^\s*เพิ่มวันนัดหมาย\s*$/.test(txt(b)))
                        .map(b => ({text: txt(b), cls: (b.className||'').toString().slice(0, 100)})),
                    inputs: Array.from(document.querySelectorAll('input, select'))
                        .filter(vis)
                        .map(e => ({tag: e.tagName, id: e.id, name: e.name,
                                    placeholder: e.placeholder || ''})),
                };
            }""")
            print("\n=== elements ที่เจอ ===")
            print(json.dumps(elements, ensure_ascii=False, indent=2))

            # ขั้น 2: เปิด dropdown สถานที่ → เลือกสาขาที่มี slot
            print("\n[i] เปิด dropdown สถานที่...")
            qp.evaluate(r"""() => {
                const vis = el => !!(el && el.offsetParent !== null);
                const els = Array.from(document.querySelectorAll('button, div, [role="combobox"]'))
                    .filter(vis)
                    .filter(e => /^กรุณาเลือกสถานที่$/.test((e.textContent||'').trim()));
                if (els[0]) els[0].click();
            }""")
            qp.wait_for_timeout(2000)
            qp.screenshot(path=str(SHOTS / "02_dropdown_open.png"), full_page=True)

            # ดู option ที่มี
            options = qp.evaluate(r"""() => {
                const vis = el => !!(el && el.offsetParent !== null);
                return Array.from(document.querySelectorAll('div, li, button, option'))
                    .filter(vis)
                    .filter(e => /ศูนย์บริการ|ศูนย์แรกรับ/.test((e.textContent||'')))
                    .slice(0, 20).map(e => ({tag: e.tagName,
                                             text: ((e.textContent||'').replace(/\s+/g,' ').trim().slice(0, 100)),
                                             cls: (e.className||'').toString().slice(0, 80)}));
            }""")
            print(f"\n[options] {len(options)}:")
            for i, o in enumerate(options[:10]):
                print(f"  [{i}] {o['tag']} {o['text']!r}")

            # เลือกอันแรก
            if options:
                print(f"\n[i] เลือก: {options[0]['text']!r}")
                captured.clear()
                qp.evaluate(r"""() => {
                    const vis = el => !!(el && el.offsetParent !== null);
                    const el = Array.from(document.querySelectorAll('div, li, button, option'))
                        .filter(vis)
                        .filter(e => /ศูนย์บริการ|ศูนย์แรกรับ/.test((e.textContent||'')))[0];
                    if (el) el.click();
                }""")
                qp.wait_for_timeout(3000)
                qp.screenshot(path=str(SHOTS / "03_place_selected.png"), full_page=True)
                print(f"[APIs after place select]:")
                for c in captured:
                    print(f"  {c['status']} {c['url'][:200]}")

            # ขั้น 3: คลิก 'เพิ่มวันนัดหมาย'
            print("\n[i] คลิก 'เพิ่มวันนัดหมาย' ...")
            captured.clear()
            clicked = qp.evaluate(r"""() => {
                const vis = el => !!(el && el.offsetParent !== null);
                const b = Array.from(document.querySelectorAll('button'))
                    .filter(vis)
                    .find(x => /^\s*เพิ่มวันนัดหมาย\s*$/.test((x.textContent||'').trim()));
                if (b) { b.click(); return true; }
                return false;
            }""")
            print(f"clicked: {clicked}")
            qp.wait_for_timeout(4000)
            qp.screenshot(path=str(SHOTS / "04_calendar_opened.png"), full_page=True)
            print(f"[APIs after 'add day']:")
            for c in captured[:10]:
                print(f"  {c['status']} {c['url'][:200]}")

            # ขั้น 4: หาวันที่ยังว่าง (ปฏิทินโผล่แล้ว) แล้วคลิก
            captured.clear()
            print("\n[i] หาวันในปฏิทิน + คลิก...")
            day_info = qp.evaluate(r"""() => {
                const vis = el => !!(el && el.offsetParent !== null);
                // React calendar อาจใช้ [role='gridcell'] หรือ button ตัวเลข
                const candidates = Array.from(document.querySelectorAll(
                    '[role="gridcell"], button, td, div, span'
                )).filter(vis).filter(e => /^\s*\d{1,2}\s*$/.test((e.textContent||'').trim()));
                // filter: ต้องไม่ disabled + ต้องอยู่ในปฏิทินที่โผล่ล่าสุด
                const active = candidates.filter(e => !e.disabled &&
                    !(e.className||'').toString().includes('disabled') &&
                    !e.getAttribute('aria-disabled'));
                return {
                    total: candidates.length,
                    active: active.length,
                    sample: active.slice(0, 15).map(e => ({
                        text: (e.textContent||'').trim(),
                        cls: (e.className||'').toString().slice(0,100),
                        role: e.getAttribute('role') || '',
                        aria: e.getAttribute('aria-label') || '',
                    })),
                };
            }""")
            print(f"[days] total={day_info['total']}, active={day_info['active']}")
            for d in day_info['sample'][:8]:
                print(f"  · {d['text']!r}  role={d['role']!r}  aria={d['aria'][:50]!r}")

            # คลิกวันในลำดับกลาง ๆ (เลข 15+)
            qp.evaluate(r"""() => {
                const vis = el => !!(el && el.offsetParent !== null);
                const cands = Array.from(document.querySelectorAll('[role="gridcell"], button, td, div, span'))
                    .filter(vis).filter(e => /^\s*\d{1,2}\s*$/.test((e.textContent||'').trim()))
                    .filter(e => !e.disabled && !(e.className||'').toString().includes('disabled'));
                // เลือก index กลาง — น่าจะเป็นวันในเดือน
                const target = cands.find(e => {
                    const n = parseInt((e.textContent||'').trim());
                    return n >= 15 && n <= 28;
                });
                if (target) target.click();
            }""")
            qp.wait_for_timeout(4000)
            qp.screenshot(path=str(SHOTS / "05_after_click_day.png"), full_page=True)

            print(f"\n[APIs after click day]:")
            for c in captured[:20]:
                print(f"\n  {c['status']} {c['url'][:220]}")
                if len(c['body_first']) < 1200:
                    print(f"    body: {c['body_first'][:800]}")
                else:
                    print(f"    body (first 400): {c['body_first'][:400]}")

            # ขั้น 5: ดู body หลังคลิกวัน — มี "เลือกช่วงเวลา" list ไหม
            body_final = qp.evaluate("() => (document.body ? document.body.innerText : '').slice(0, 3500)")
            print("\n=== body FINAL หลังคลิกวัน ===")
            print(body_final)

            OUT_JSON.write_text(json.dumps({
                "elements": elements,
                "options": options,
                "day_info": day_info,
                "body_final": body_final,
                "captured_final": captured,
            }, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"\n[done] → {OUT_JSON}")

            input("\n Enter to close ...")
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
