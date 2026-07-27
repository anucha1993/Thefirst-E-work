"""กดปุ่ม "พิมพ์แบบฟอร์มนัดหมาย" ใน iframe แล้ว capture:
- network requests (URL ที่ยิงไป, response)
- new page / download event
- PDF bytes ถ้าได้
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
SHOTS = ROOT / "screenshots" / "appt_print"
SHOTS.mkdir(parents=True, exist_ok=True)


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }
    captured_requests = []
    captured_responses = []
    new_pages = []
    downloads = []

    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, args=["--start-maximized"])
        ctx = b.new_context(locale="th-TH", viewport={"width": 1600, "height": 1000},
                            accept_downloads=True)

        def on_req(req):
            u = req.url or ""
            # log ทุก request หลังกดปุ่ม (filter ให้อ่านง่าย)
            if any(k in u for k in ("booking", "print", "queue", "form", "appointment",
                                     "pdf", ".pdf", "GetAppoint", "GetPrint")):
                captured_requests.append({"method": req.method, "url": u[:220]})

        def on_res(res):
            u = res.url or ""
            ct = res.headers.get("content-type", "")
            if "pdf" in ct.lower() or ".pdf" in u.lower():
                captured_responses.append({
                    "status": res.status, "url": u[:200],
                    "content_type": ct[:80],
                    "content_length": res.headers.get("content-length", ""),
                })

        ctx.on("request", on_req)
        ctx.on("response", on_res)

        def on_page(pg):
            new_pages.append(pg)
            print(f"[NEW PAGE] {pg.url[:150]}")

        ctx.on("page", on_page)

        def on_download(dl):
            downloads.append(dl)
            print(f"[DOWNLOAD] suggested={dl.suggested_filename} url={dl.url[:150]}")

        page = ctx.new_page()
        page.on("download", on_download)
        try:
            login(page, cfg)
            page.wait_for_timeout(1500)
            page.goto(TARGET_URL, wait_until="domcontentloaded", timeout=45_000)
            page.wait_for_timeout(4000)

            # trigger tab การนัดหมาย
            page.evaluate("""() => {
                if (window.jQuery) jQuery('a[href="#tab_default_5"]').tab('show');
                const a = document.querySelector('a[href="#tab_default_5"]');
                if (a) a.click();
            }""")
            page.wait_for_timeout(3500)

            # หา iframe frame
            appt_frame = None
            for _ in range(10):
                for fr in page.frames:
                    if "queue" in fr.url or "bookingdate" in fr.url or "bookingdetail" in fr.url:
                        appt_frame = fr; break
                if appt_frame: break
                page.wait_for_timeout(600)

            if not appt_frame:
                print("[!] ไม่พบ iframe queue"); return 1

            print(f"[+] iframe: {appt_frame.url[:150]}")
            appt_frame.wait_for_load_state("domcontentloaded", timeout=15_000)
            page.wait_for_timeout(2500)

            page.screenshot(path=str(SHOTS / "01_before_print.png"), full_page=True)

            # หาปุ่ม + inspect onclick
            btn_info = appt_frame.evaluate(r"""() => {
                const bts = Array.from(document.querySelectorAll('button'));
                const found = bts.filter(b => /^\s*พิมพ์แบบฟอร์มนัดหมาย\s*$/.test((b.textContent||'').trim()));
                return found.map(b => ({
                    text: (b.textContent||'').trim(),
                    id: b.id, cls: (b.className||'').toString().slice(0, 100),
                    onclick: (b.getAttribute('onclick')||'').slice(0,200),
                    outerHTML: b.outerHTML.slice(0, 500),
                }));
            }""")
            print(f"\n[buttons 'พิมพ์แบบฟอร์มนัดหมาย'] {btn_info}")

            # กดปุ่ม "พิมพ์แบบฟอร์มนัดหมาย" (single, ไม่ใช่ 'ทั้งหมด')
            print("\n[i] คลิก 'พิมพ์แบบฟอร์มนัดหมาย' ...")

            # เตรียม catch new page (popup) ก่อนคลิก
            popup_page = None
            try:
                with ctx.expect_page(timeout=8000) as pop_info:
                    appt_frame.evaluate(r"""() => {
                        const bts = Array.from(document.querySelectorAll('button'));
                        const b = bts.find(x => /^\s*พิมพ์แบบฟอร์มนัดหมาย\s*$/.test((x.textContent||'').trim()));
                        if (b) b.click();
                    }""")
                popup_page = pop_info.value
                print(f"[POPUP page detected] {popup_page.url[:150]}")
            except Exception as e:
                print(f"[i] no popup: {str(e).splitlines()[0][:80]}")

            page.wait_for_timeout(6000)  # รอ network activities

            print("\n=== new pages opened ===")
            for i, pg in enumerate(new_pages, 1):
                print(f"  [{i}] {pg.url[:200]}")
                try:
                    pg.wait_for_load_state("domcontentloaded", timeout=8000)
                except Exception:
                    pass
                try:
                    ttl = pg.title()
                    print(f"      title={ttl}")
                except Exception:
                    pass
                try:
                    pg.screenshot(path=str(SHOTS / f"popup_{i}.png"), full_page=True)
                except Exception as e:
                    print(f"      screenshot err: {e}")

            print(f"\n=== captured requests (booking/print/pdf) ===")
            for r in captured_requests:
                print(f"  {r['method']} {r['url']}")

            print(f"\n=== captured pdf responses ===")
            for r in captured_responses:
                print(f"  {r['status']} {r['url']}  ct={r['content_type']} len={r['content_length']}")

            print(f"\n=== downloads ({len(downloads)}) ===")
            for i, d in enumerate(downloads, 1):
                fn = SHOTS / f"downloaded_{i}_{d.suggested_filename}"
                try:
                    d.save_as(str(fn))
                    print(f"  [{i}] saved → {fn} ({fn.stat().st_size:,} bytes)")
                except Exception as e:
                    print(f"  [{i}] save err: {e}")

            input("\n Enter to close ...")
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
