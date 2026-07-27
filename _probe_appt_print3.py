"""Test: กดปุ่ม 'พิมพ์แบบฟอร์มนัดหมาย' โดยไม่ hook window.print
- ดูว่า DOM เปลี่ยนไหมหลังกด
- ดูว่า @media print ทำงานอย่างไร
- ลอง page.pdf() หลังกด (ตอน React อาจ swap DOM)
"""
from __future__ import annotations
import os
import sys
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import login, ROOT  # noqa: E402

TARGET_URL = "https://eworkpermit.doe.go.th/RenewMOU/DetailFormRenewMOU?user_id=303816&status=APSS&group_id=69123000059307&form_type=MT_59_MOU_RENEWAL"
SHOTS = ROOT / "screenshots" / "appt_print3"
SHOTS.mkdir(parents=True, exist_ok=True)


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }
    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, args=["--start-maximized"])
        ctx = b.new_context(locale="th-TH", viewport={"width": 1600, "height": 1000},
                            accept_downloads=True)
        page = ctx.new_page()

        try:
            login(page, cfg)
            page.wait_for_timeout(1500)
            page.goto(TARGET_URL, wait_until="domcontentloaded", timeout=45_000)
            page.wait_for_timeout(4000)

            # trigger tab
            page.evaluate("""() => {
                if (window.jQuery) jQuery('a[href="#tab_default_5"]').tab('show');
            }""")
            page.wait_for_timeout(3000)

            # หา frame
            appt_frame = None
            for _ in range(15):
                for fr in page.frames:
                    if fr.name == "link_appointment":
                        appt_frame = fr; break
                if appt_frame: break
                page.wait_for_timeout(500)

            appt_frame.wait_for_load_state("domcontentloaded", timeout=15_000)
            page.wait_for_timeout(3000)

            # เปิด iframe URL ใน tab ใหม่
            iframe_src = page.evaluate(
                "() => document.querySelector('#link_appointment').src"
            )
            new_page = ctx.new_page()
            new_page.goto(iframe_src, wait_until="networkidle", timeout=30_000)
            new_page.wait_for_timeout(4000)

            new_page.screenshot(path=str(SHOTS / "01_before_click.png"), full_page=True)
            html_before = new_page.evaluate("() => document.body.innerHTML.length")
            print(f"[before click] body innerHTML len = {html_before}")

            # dump ปุ่มทั้งหมด (พร้อม onclick / React handlers) ก่อนกด
            btns = new_page.evaluate(r"""() => {
                return Array.from(document.querySelectorAll('button')).map(b => ({
                    text: (b.textContent||'').trim().slice(0, 80),
                    cls: (b.className||'').toString().slice(0, 100),
                    hasReactProps: !!Object.keys(b).find(k => k.startsWith('__react')),
                    outerHTMLLen: b.outerHTML.length,
                }));
            }""")
            print("\n[buttons on page]")
            for i, bt in enumerate(btns):
                print(f"  [{i}] {bt['text']!r:50} react={bt['hasReactProps']}")

            # Hook window.print แต่ให้ callback set flag
            new_page.evaluate("""() => {
                window.__printCalled = false;
                window.__docBeforePrint = '';
                window.__docAtPrint = '';
                window.__docBeforePrint = document.body.innerHTML;
                const _print = window.print;
                window.print = function() {
                    window.__printCalled = true;
                    window.__docAtPrint = document.body.innerHTML;
                    // ไม่เรียก _print เพื่อไม่เปิด print dialog (จะบล็อก playwright)
                };
            }""")

            # กดปุ่ม 'พิมพ์แบบฟอร์มนัดหมาย' (single not 'ทั้งหมด')
            print("\n[i] คลิก 'พิมพ์แบบฟอร์มนัดหมาย' ...")
            new_page.locator(
                "button:has-text('พิมพ์แบบฟอร์มนัดหมาย')"
            ).nth(0).click(timeout=5000)
            new_page.wait_for_timeout(3000)

            # check ว่า DOM เปลี่ยนไหม
            info = new_page.evaluate("""() => ({
                printCalled: window.__printCalled || false,
                bodyLenBefore: (window.__docBeforePrint||'').length,
                bodyLenAt: (window.__docAtPrint||'').length,
                bodyLenNow: document.body.innerHTML.length,
                changed: window.__docBeforePrint !== window.__docAtPrint,
                changedNow: window.__docBeforePrint !== document.body.innerHTML,
            })""")
            print(f"[after click] print={info['printCalled']}  "
                  f"bodyBefore={info['bodyLenBefore']}  bodyAt={info['bodyLenAt']}  "
                  f"bodyNow={info['bodyLenNow']}  changed={info['changed']}  "
                  f"changedNow={info['changedNow']}")

            new_page.screenshot(path=str(SHOTS / "02_after_click.png"), full_page=True)

            # Dump diff — เช็คว่ามี element ใหม่ที่ไม่เคยมี
            new_elems = new_page.evaluate(r"""() => {
                const before = window.__docBeforePrint || '';
                const now = document.body.innerHTML;
                // นับ element โดยรวม
                return {
                    hasBarcode: !!document.querySelector('svg[id*="barcode"], canvas[id*="barcode"], [class*="barcode"], img[src*="barcode"]'),
                    imgCount: document.querySelectorAll('img').length,
                    svgCount: document.querySelectorAll('svg').length,
                    canvasCount: document.querySelectorAll('canvas').length,
                    urls: [location.href],
                    // check for print-only classes
                    printOnly: Array.from(document.querySelectorAll('[class*="print"], [class*="Print"]'))
                        .slice(0, 10).map(e => ({tag: e.tagName, cls: e.className.toString().slice(0,100)})),
                };
            }""")
            print(f"\n[dom after click]")
            print(f"  barcode: {new_elems['hasBarcode']}, img={new_elems['imgCount']}, "
                  f"svg={new_elems['svgCount']}, canvas={new_elems['canvasCount']}")
            print(f"  print-only elements: {new_elems['printOnly']}")

            # ลอง save PDF ตอน DOM หลังคลิก
            try:
                new_page.emulate_media(media="print")
                pdf_path = SHOTS / "after_click.pdf"
                new_page.pdf(path=str(pdf_path), format="A4", print_background=True)
                print(f"\n[pdf after click] {pdf_path.stat().st_size:,} bytes")
            except Exception as e:
                print(f"[!] pdf err: {e}")

            # dump raw HTML รวม body innerHTML ทั้งหมด
            raw = new_page.evaluate("() => document.body.innerHTML")
            (SHOTS / "body_after_click.html").write_text(raw, encoding="utf-8")
            print(f"[body html] {len(raw)} bytes → {SHOTS}/body_after_click.html")

            input("\n Enter to close ...")
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
