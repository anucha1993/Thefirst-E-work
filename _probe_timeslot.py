"""Probe: มี time slot ต่อวัน หรือแค่วันเดียว?
- เปิด iframe ของ record ที่ยังไม่จอง (status=AP)
- capture ทุก API call ตอน click date → เพื่อหา time-slot API
"""
from __future__ import annotations
import json, os, sys
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import login, ROOT, goto_tracking, apply_wa_filter, collect_all_wa_rows, build_detail_url

OUT_JSON = ROOT / "_probe_timeslot_out.json"
SHOTS = ROOT / "screenshots" / "timeslot"
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
                try:
                    text = res.text()
                except Exception:
                    text = ""
                captured.append({
                    "status": res.status,
                    "url": u[:250],
                    "body_first": text[:1500],
                })
        ctx.on("response", on_response)

        page = ctx.new_page()
        try:
            login(page, cfg)
            page.wait_for_timeout(1500)

            # หา record status=AP (ยังไม่จอง)
            goto_tracking(page)
            apply_wa_filter(page, "", status_ids=["AP"])
            rows = collect_all_wa_rows(page, log=lambda *a: None)
            ap_rows = [r for r in rows if str(r.get("status") or "").upper() == "AP"]
            print(f"AP rows: {len(ap_rows)}")
            if not ap_rows:
                print("[!] ไม่มี AP record → ทดสอบไม่ได้"); return 1

            target = ap_rows[0]
            url = build_detail_url(target)
            print(f"target: {target.get('reqNo')}  url: {url[:100]}")
            page.goto(url, wait_until="domcontentloaded", timeout=25_000)
            page.wait_for_timeout(3000)

            # trigger tab การนัดหมาย
            page.evaluate("() => { if(window.jQuery) jQuery('a[href=\"#tab_default_5\"]').tab('show'); }")
            page.wait_for_timeout(3500)

            iframe_src = page.evaluate("() => (document.querySelector('#link_appointment') || {}).src || ''")
            print(f"iframe: {iframe_src[:120]}...")

            qp = ctx.new_page()
            captured.clear()  # เริ่มบันทึกใหม่จากตรงนี้
            qp.goto(iframe_src, wait_until="networkidle", timeout=30_000)
            qp.wait_for_timeout(6000)  # รอ FE load

            qp.screenshot(path=str(SHOTS / "01_calendar_view.png"), full_page=True)

            body_txt = qp.evaluate("() => (document.body ? document.body.innerText : '').slice(0, 3000)")
            print("\n=== body ที่โผล่ ===")
            print(body_txt[:1500])

            # เก็บ initial API calls
            initial_apis = list(captured)
            print(f"\n[initial APIs] {len(initial_apis)} calls:")
            for c in initial_apis:
                print(f"  {c['status']} {c['url']}")

            # ลอง click วันเปิด (เขียว/ไม่ full) ในปฏิทิน
            # โครงสร้าง: React calendar อาจเป็น <button> or <div> ที่มี date
            captured.clear()
            print("\n[i] ลอง click วันในปฏิทิน...")
            # หา element ที่ text = ตัวเลขเดือน + ไม่ disabled
            clicked = qp.evaluate(r"""() => {
                const vis = el => !!(el && el.offsetParent !== null);
                const els = Array.from(document.querySelectorAll('button, div, span, td'))
                    .filter(vis)
                    .filter(e => /^\s*\d{1,2}\s*$/.test((e.textContent||'').trim()))
                    .filter(e => {
                        const cls = (e.className||'').toString();
                        return !cls.includes('disabled') && !e.disabled;
                    });
                // เลือกตัวที่กลาง ๆ (skip 1-5 ที่อาจเป็นวันผ่านไปแล้ว)
                const target = els[Math.min(15, els.length - 1)];
                if (target) {
                    target.click();
                    return {
                        clicked: (target.textContent||'').trim(),
                        cls: (target.className||'').toString().slice(0, 100),
                    };
                }
                return null;
            }""")
            print(f"[clicked] {clicked}")
            qp.wait_for_timeout(4000)
            qp.screenshot(path=str(SHOTS / "02_after_click_date.png"), full_page=True)

            # จับ APIs หลัง click
            after_apis = list(captured)
            print(f"\n[after-click APIs] {len(after_apis)} calls:")
            for c in after_apis[:20]:
                url_short = c['url']
                print(f"  {c['status']} {url_short[:180]}")
                # ถ้ามี body สั้น ๆ ที่ดู promising ให้ preview
                if len(c['body_first']) < 800:
                    print(f"    body: {c['body_first'][:400]}")

            # ดู body หลังคลิก
            body_after = qp.evaluate("() => (document.body ? document.body.innerText : '').slice(0, 3000)")
            print("\n=== body หลังคลิกวัน ===")
            print(body_after[:1500])

            OUT_JSON.write_text(
                json.dumps({
                    "initial_apis": initial_apis,
                    "after_click_apis": after_apis,
                    "body_initial": body_txt,
                    "body_after_click": body_after,
                }, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            print(f"\n[done] → {OUT_JSON}")

            input("\n Enter to close ...")
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
