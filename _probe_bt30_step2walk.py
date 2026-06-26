"""ทดสอบ step2 walk เต็ม (รวม 4.2 ที่แก้แล้ว) โดยไม่ติดปัญหา step1 false-REVIEW.
ค้นหา (ignore สถานะ) → เรียก _bt30_do_step2 ตรงๆ (เดินครบ 2.1–2.9 รวมแนบเอกสาร).
รัน: .venv\\Scripts\\python.exe _probe_bt30_step2walk.py   (ปิดเองอัตโนมัติ)
"""
from __future__ import annotations

from playwright.sync_api import sync_playwright

from scrape_wa import (
    login, _read_login_accounts, _read_bt30_excel, _open_bt30_form,
    _bt30_fill_search_one, _close_register_alert, _bt30_dismiss_news,
    _bt30_do_step2, ROOT,
)

SHOTS = ROOT / "reports" / "bt30_screenshots"
SHOTS.mkdir(parents=True, exist_ok=True)
log = print


def main():
    accounts = _read_login_accounts(ROOT / "UsernameLogin.xlsx")
    acct = next(iter(accounts.values()))
    cfg = {"username": acct["username"], "password": acct["password"],
           "user_type": acct["type"], "method": acct.get("method") or "E-Workpermit"}
    rec = _read_bt30_excel(ROOT / "from_bt30.xlsx")[0]
    log(f"[i] doc_others = {rec.get('doc_others')}")

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False, args=["--ignore-certificate-errors", "--start-maximized"])
        ctx = browser.new_context(locale="th-TH", ignore_https_errors=True, viewport={"width": 1600, "height": 950})
        page = ctx.new_page()
        try:
            login(page, cfg); page.wait_for_timeout(1500)
            if not _open_bt30_form(page, log=log):
                log("[!] เปิดฟอร์มไม่สำเร็จ"); return
            r1 = _bt30_fill_search_one(page, rec, SHOTS, log=log)
            log(f"[i] ขั้นตอน1: {r1.get('status')} | {str(r1.get('note'))[:80]}")
            # ไม่ bail แม้ REVIEW — เดินหน้าต่อ (หน้ารายละเอียดโหลดอยู่)
            _close_register_alert(page)
            page.wait_for_timeout(2500)
            _bt30_dismiss_news(page, log=log)

            r2 = _bt30_do_step2(page, rec, SHOTS, log=log)
            log(f"\n[i] ขั้นตอน2: {r2.get('step2_status')}")
            log(f"[i] หมายเหตุ: {r2.get('step2_note')}")
        except Exception as e:
            log("[ERR]", e)
            try:
                page.screenshot(path=str(SHOTS / "_walk_error.png"), full_page=True)
            except Exception:
                pass
        finally:
            input("กด Enter เพื่อปิดเบราว์เซอร์...")
            ctx.close(); browser.close()


if __name__ == "__main__":
    main()
