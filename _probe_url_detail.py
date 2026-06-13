"""ทดสอบ: เปิดหน้า detail ด้วยการ 'เปลี่ยน URL ตรง' (build_detail_url) แทนการคลิกตาราง
- login บัญชีแรกใน UsernameLogin.xlsx
- เก็บรายการคำขอ (filter AP + SS)
- ลอง navigate ตรงผ่าน URL กับ 3 เลขแรก แล้วเช็คว่า detail โหลดได้ไหม
รันแบบเห็นเบราว์เซอร์ (headless=False) เผื่อมี reCAPTCHA
"""
from pathlib import Path

from playwright.sync_api import sync_playwright

import scrape_wa as S

LOGIN_XLSX = Path("UsernameLogin.xlsx")
STATUS_IDS = ["AP", "SS"]
TEST_N = 3  # ทดสอบกี่เลขแรก


def main() -> None:
    accounts = S._read_login_accounts(LOGIN_XLSX)
    if not accounts:
        print("[!] ไม่พบบัญชีใน UsernameLogin.xlsx")
        return
    ukey, acct = next(iter(accounts.items()))
    login_cfg = {
        "username": acct["username"],
        "password": acct["password"],
        "user_type": acct["type"],
        "method": acct.get("method") or "E-Workpermit",
    }
    print(f"[i] ใช้บัญชี: {acct['username']} ({acct['type']})")

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=False,
            args=["--ignore-certificate-errors", "--start-maximized"],
        )
        ctx = browser.new_context(
            locale="th-TH", ignore_https_errors=True,
            viewport={"width": 1920, "height": 1080}, accept_downloads=True,
        )
        page = ctx.new_page()
        try:
            S.login(page, login_cfg)
            S.goto_tracking(page)
            S.apply_wa_filter(page, "", status_ids=STATUS_IDS)
            rows = S.collect_all_wa_rows(page, log=print)
            print(f"\n[i] เก็บได้ {len(rows)} คำขอ — จะทดสอบ {min(TEST_N, len(rows))} เลขแรก\n")

            ok = 0
            for i, row in enumerate(rows[:TEST_N], start=1):
                req_no = row.get("reqNo", "")
                url = S.build_detail_url(row)
                print(f"--- [{i}] คำขอ {req_no}")
                print(f"    form_type = {row.get('form_type','')}")
                print(f"    URL       = {url}")
                try:
                    page.goto(url, wait_until="domcontentloaded", timeout=25_000)
                    page.wait_for_timeout(1500)
                except Exception as e:
                    print(f"    ✗ goto ล้มเหลว: {e}")
                    continue
                print(f"    landed    = {page.url}")
                loaded = S._is_detail_loaded(page)
                print(f"    detail โหลดสำเร็จ? {'✓ ใช่' if loaded else '✗ ไม่'}")
                if loaded:
                    name = S._extract_alien_eng_name(page)
                    print(f"    ชื่อ(Eng) = {name}")
                    ok += 1

            print(f"\n[สรุป] navigate ตรงสำเร็จ {ok}/{min(TEST_N, len(rows))} เลข")
            print("[i] หยุดค้างไว้ 8 วิให้ดูหน้าจอ...")
            page.wait_for_timeout(8000)
        finally:
            ctx.close(); browser.close()


if __name__ == "__main__":
    main()
