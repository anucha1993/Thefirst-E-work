"""Probe: หน้าจองการนัดหมาย (record status=AP, ยังไม่จอง)
- เปิด iframe ของ record ที่ยังไม่ได้จอง → เห็น calendar UI
- capture calendar API response → ทราบวันที่ว่างให้จอง
"""
from __future__ import annotations
import json
import os
import sys
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import login, ROOT  # noqa: E402

# record status=AP (ยังไม่จอง) — เอาจาก account เดิม
TARGET_URL = "https://eworkpermit.doe.go.th/RenewMOU/DetailFormRenewMOU?user_id=303816&status=AP&group_id=69123000060887&form_type=MT_59_MOU_RENEWAL"
OUT_JSON = ROOT / "_probe_appt_booking_out.json"
SHOTS = ROOT / "screenshots" / "appt_booking"
SHOTS.mkdir(parents=True, exist_ok=True)


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }

    captured_api = []

    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, args=["--start-maximized"])
        ctx = b.new_context(locale="th-TH", viewport={"width": 1600, "height": 1000},
                            accept_downloads=True)

        # จับ response ของ calendar API
        def on_response(res):
            u = res.url or ""
            if "queue-be" in u and ("calendar" in u or "branch" in u or "annotation" in u):
                try:
                    body_text = res.text()
                except Exception:
                    body_text = ""
                captured_api.append({
                    "status": res.status,
                    "url": u[:250],
                    "body_preview": body_text[:2000],
                })

        ctx.on("response", on_response)

        page = ctx.new_page()
        try:
            login(page, cfg)
            page.wait_for_timeout(1500)
            page.goto(TARGET_URL, wait_until="domcontentloaded", timeout=45_000)
            page.wait_for_timeout(4000)

            # trigger tab การนัดหมาย
            page.evaluate("""() => {
                if (window.jQuery) jQuery('a[href="#tab_default_5"]').tab('show');
            }""")
            page.wait_for_timeout(3500)

            # หา iframe
            appt_frame = None
            for _ in range(15):
                for fr in page.frames:
                    if fr.name == "link_appointment":
                        appt_frame = fr; break
                if appt_frame: break
                page.wait_for_timeout(500)

            print(f"[+] iframe: {appt_frame.url[:150]}")
            iframe_src = page.evaluate(
                "() => document.querySelector('#link_appointment').src"
            )

            # เปิด iframe URL ใน new tab เพื่อเห็น booking form ชัดขึ้น
            new_page = ctx.new_page()
            new_page.goto(iframe_src, wait_until="networkidle", timeout=30_000)
            new_page.wait_for_timeout(5000)
            new_page.screenshot(path=str(SHOTS / "01_booking_page.png"), full_page=True)

            # dump body content ของ new_page
            body_txt = new_page.evaluate(
                "() => (document.body ? document.body.innerText : '').slice(0, 3000)"
            )
            print("\n=== body content ===")
            print(body_txt[:2000])

            # DOM structure — หา calendar / date cells / branch dropdown
            dom_info = new_page.evaluate(r"""() => {
                const vis = el => !!(el && el.offsetParent !== null);
                const txt = e => (e.textContent||'').replace(/\s+/g,' ').trim();
                return {
                    // ปุ่มทั้งหมดที่เห็น
                    buttons: Array.from(document.querySelectorAll('button'))
                        .filter(vis).map(b => ({
                            text: txt(b).slice(0, 60),
                            cls: (b.className||'').toString().slice(0,120),
                            disabled: b.disabled,
                            aria: b.getAttribute('aria-label') || '',
                        })).slice(0, 60),
                    // dropdown / select
                    selects: Array.from(document.querySelectorAll('select, [role="combobox"], [role="listbox"]'))
                        .filter(vis).map(s => ({
                            tag: s.tagName,
                            text: txt(s).slice(0, 80),
                            cls: (s.className||'').toString().slice(0,100),
                        })).slice(0, 20),
                    // ปฏิทิน — หา element ที่มี class calendar/date
                    calendarEls: Array.from(document.querySelectorAll(
                        '[class*="calendar"], [class*="Calendar"], [class*="date-picker"], [class*="DatePicker"]'
                    )).filter(vis).map(e => ({
                        tag: e.tagName,
                        cls: (e.className||'').toString().slice(0,100),
                        text: txt(e).slice(0, 100),
                    })).slice(0, 10),
                    // divs ที่มีเลขวันที่ (heuristic: text = 1-31)
                    dateNumberEls: Array.from(document.querySelectorAll('div, td, button, span, a'))
                        .filter(e => vis(e) && /^\s*\d{1,2}\s*$/.test((e.textContent||'').trim()))
                        .slice(0, 40).map(e => ({
                            tag: e.tagName,
                            num: (e.textContent||'').trim(),
                            cls: (e.className||'').toString().slice(0,80),
                            disabled: e.disabled || e.classList.contains('disabled'),
                            aria: e.getAttribute('aria-label') || '',
                        })),
                };
            }""")
            print("\n=== DOM structure ===")
            print(json.dumps(dom_info, ensure_ascii=False, indent=2)[:5000])

            # บันทึกทุกอย่าง
            (OUT_JSON).write_text(
                json.dumps({
                    "target_url": TARGET_URL,
                    "iframe_src": iframe_src,
                    "body_txt": body_txt,
                    "dom_info": dom_info,
                    "captured_api": captured_api,
                }, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            print(f"\n=== API responses captured ({len(captured_api)}) ===")
            for i, c in enumerate(captured_api[:20]):
                print(f"\n[{i}] status={c['status']}")
                print(f"    url: {c['url']}")
                # decode body preview if JSON
                bp = c["body_preview"]
                if bp.startswith("{") or bp.startswith("["):
                    try:
                        d = json.loads(bp)
                        print(f"    body (parsed): {json.dumps(d, ensure_ascii=False, indent=6)[:2000]}")
                    except Exception:
                        print(f"    body: {bp[:800]}")
                else:
                    print(f"    body: {bp[:600]}")

            input("\n Enter to close ...")
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
