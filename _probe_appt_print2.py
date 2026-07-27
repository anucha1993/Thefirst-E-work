"""ลอง 2 วิธี:
1. Locator.click() แบบ native + capture window.print
2. เปิด iframe URL ใน tab ใหม่ตรง ๆ → page.pdf()
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

TARGET_URL = os.getenv(
    "PROBE_APPT_URL",
    "https://eworkpermit.doe.go.th/RenewMOU/DetailFormRenewMOU?user_id=303816&status=APSS&group_id=69123000059294&form_type=MT_59_MOU_RENEWAL",
)
SHOTS = ROOT / "screenshots" / "appt_print2"
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

        # Intercept window.print + window.open on ALL pages
        ctx.add_init_script(r"""
            (function() {
                const _print = window.print;
                window.print = function() {
                    console.warn('[HOOKED] window.print() called');
                    window.__printCalled = (window.__printCalled || 0) + 1;
                    window.__printTime = new Date().toISOString();
                };
                const _open = window.open;
                window.open = function(url, target, features) {
                    console.warn('[HOOKED] window.open() called:', url);
                    window.__openCalls = window.__openCalls || [];
                    window.__openCalls.push({url: url, target: target, time: new Date().toISOString()});
                    return _open.call(window, url, target, features);
                };
            })();
        """)

        try:
            login(page, cfg)
            page.wait_for_timeout(1500)
            page.goto(TARGET_URL, wait_until="domcontentloaded", timeout=45_000)
            page.wait_for_timeout(4000)

            page.evaluate("""() => {
                if (window.jQuery) jQuery('a[href="#tab_default_5"]').tab('show');
                const a = document.querySelector('a[href="#tab_default_5"]');
                if (a) a.click();
            }""")
            page.wait_for_timeout(3500)

            # หา iframe URL
            iframe_src = page.evaluate(
                "() => (document.querySelector('#link_appointment') || {}).src || ''"
            )
            print(f"[iframe src] {iframe_src[:150]}...")

            appt_frame = None
            for _ in range(15):
                for fr in page.frames:
                    if fr.name == "link_appointment" or "bookingdate" in fr.url or "bookingdetail" in fr.url:
                        appt_frame = fr; break
                if appt_frame: break
                page.wait_for_timeout(500)

            if not appt_frame:
                print("[!] no iframe"); return 1

            appt_frame.wait_for_load_state("domcontentloaded", timeout=15_000)
            page.wait_for_timeout(3000)
            page.screenshot(path=str(SHOTS / "01_before.png"), full_page=True)

            # === Approach 1: Locator.click ผ่าน frame_locator ===
            print("\n=== Approach 1: locator click on button 'พิมพ์แบบฟอร์มนัดหมาย' ===")
            try:
                btn = page.frame_locator("#link_appointment").locator(
                    "button:has-text('พิมพ์แบบฟอร์มนัดหมาย')"
                ).nth(0)  # ตัวแรก (single, ไม่ใช่ 'ทั้งหมด')
                # เตรียม catch popup ก่อน
                popup = None
                try:
                    with ctx.expect_page(timeout=8000) as pop_info:
                        btn.click(timeout=5000)
                    popup = pop_info.value
                    print(f"[POPUP] {popup.url[:200]}")
                except Exception as _e:
                    print(f"[i] no popup ({str(_e).splitlines()[0][:80]})")

                page.wait_for_timeout(3000)

                # ตรวจ window.print / window.open hook ใน main page + iframe
                for label, ctx_source in [("main", page), ("iframe", appt_frame)]:
                    try:
                        info = ctx_source.evaluate(
                            "() => ({print: window.__printCalled || 0, "
                            "printTime: window.__printTime || '',"
                            "opens: window.__openCalls || []})"
                        )
                        print(f"  {label}: print={info['print']}, opens={info['opens']}, "
                              f"printTime={info.get('printTime','')}")
                    except Exception as e:
                        print(f"  {label}: error {e}")

                # check for new page
                pages_all = ctx.pages
                print(f"  pages in ctx: {len(pages_all)}")
                for i, p in enumerate(pages_all):
                    print(f"    [{i}] {p.url[:150]}")
                    try:
                        p.screenshot(path=str(SHOTS / f"page_{i}.png"), full_page=True)
                    except Exception:
                        pass
            except Exception as e:
                print(f"[!] approach 1 err: {e}")

            # === Approach 2: เปิด iframe URL ตรง ๆ ใน tab ใหม่ ===
            print("\n=== Approach 2: navigate to iframe URL directly & page.pdf() ===")
            try:
                new_page = ctx.new_page()
                # cookies ต้อง share ผ่าน ctx เดียวกันอยู่แล้ว
                print(f"[i] goto iframe URL...")
                new_page.goto(iframe_src, wait_until="networkidle", timeout=30_000)
                new_page.wait_for_timeout(4000)
                new_page.screenshot(path=str(SHOTS / "02_direct_url.png"), full_page=True)
                print(f"    title: {new_page.title()}")
                bt = new_page.evaluate(
                    "() => (document.body && document.body.innerText || '').slice(0, 400)"
                )
                print(f"    body: {bt[:400]}")

                # ลอง page.pdf() — ต้อง emulate print media
                new_page.emulate_media(media="print")
                pdf_path = SHOTS / "appointment.pdf"
                new_page.pdf(
                    path=str(pdf_path),
                    format="A4",
                    print_background=True,
                    margin={"top": "10mm", "bottom": "10mm", "left": "10mm", "right": "10mm"},
                )
                sz = pdf_path.stat().st_size
                print(f"    ✓ PDF saved: {pdf_path.name} ({sz:,} bytes)")

                # ลองคลิกปุ่ม 'พิมพ์แบบฟอร์มนัดหมาย' ในหน้าใหม่นี้ด้วย
                print("\n    trying button click on direct-URL page...")
                new_page.wait_for_timeout(2000)
                try:
                    with ctx.expect_page(timeout=8000) as pop_info:
                        new_page.locator(
                            "button:has-text('พิมพ์แบบฟอร์มนัดหมาย')"
                        ).nth(0).click(timeout=5000)
                    popup2 = pop_info.value
                    print(f"    [POPUP2] {popup2.url[:200]}")
                    popup2.wait_for_load_state("domcontentloaded", timeout=8000)
                    popup2.screenshot(path=str(SHOTS / "popup2.png"), full_page=True)
                    # save PDF from popup2
                    try:
                        popup2.emulate_media(media="print")
                        pdf2 = SHOTS / "appointment_from_popup.pdf"
                        popup2.pdf(path=str(pdf2), format="A4", print_background=True)
                        print(f"    ✓ popup PDF: {pdf2.name} ({pdf2.stat().st_size:,} bytes)")
                    except Exception as e:
                        print(f"    popup pdf err: {e}")
                except Exception as e:
                    print(f"    no popup: {str(e).splitlines()[0][:100]}")

                print_hook = new_page.evaluate(
                    "() => ({print: window.__printCalled || 0, opens: window.__openCalls || []})"
                )
                print(f"    hooks: {print_hook}")

            except Exception as e:
                print(f"[!] approach 2 err: {e}")

            input("\n Enter to close ...")
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
