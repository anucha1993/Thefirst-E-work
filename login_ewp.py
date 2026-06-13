"""
สคริปต์ Login เข้าระบบ e-WorkPermit (https://eworkpermit.doe.go.th)
อ่าน credentials จากไฟล์ .env แล้วทำการ login อัตโนมัติด้วย Playwright

ติดตั้ง:
    pip install -r requirements.txt
    playwright install chromium

ใช้งาน:
    1) คัดลอก .env.example -> .env แล้วใส่ username/password จริง
    2) python login_ewp.py

หมายเหตุ: หน้า Login มี Google reCAPTCHA ป้องกันบอท
- ถ้าเป็น reCAPTCHA v3 (invisible) มักผ่านได้ในโหมด headful
- ถ้าโดน challenge (เลือกภาพ) ต้องคลิกเองในหน้าต่างเบราว์เซอร์
- แนะนำตั้ง EWP_HEADLESS=false เพื่อให้แก้ captcha ได้สะดวก
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import (
    Page,
    Playwright,
    TimeoutError as PWTimeoutError,
    sync_playwright,
)

LOGIN_URL = "https://eworkpermit.doe.go.th/Login"
HOME_URL = "https://eworkpermit.doe.go.th/"

SCREENSHOTS_DIR = Path(__file__).parent / "screenshots"
STORAGE_STATE = Path(__file__).parent / "storage_state.json"


def env_bool(name: str, default: bool = False) -> bool:
    val = os.getenv(name)
    if val is None:
        return default
    return val.strip().lower() in {"1", "true", "yes", "y", "on"}


def load_config() -> dict:
    load_dotenv(Path(__file__).parent / ".env")
    cfg = {
        "username": os.getenv("EWP_USERNAME", "").strip(),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน").strip(),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit").strip(),
        "headless": env_bool("EWP_HEADLESS", False),
    }
    missing = [k for k in ("username", "password") if not cfg[k]]
    if missing:
        raise SystemExit(
            f"[ERROR] ไม่พบค่าใน .env: {', '.join('EWP_' + m.upper() for m in missing)}\n"
            f"        คัดลอก .env.example -> .env แล้วใส่ค่าให้ครบ"
        )
    return cfg


def login(page: Page, cfg: dict) -> bool:
    print(f"[+] เปิดหน้า Login: {LOGIN_URL}")
    page.goto(LOGIN_URL, wait_until="domcontentloaded")

    # 1) เลือกประเภทผู้ใช้งาน (combobox ตัวที่ 2 — ตัวแรกคือ language)
    print(f"[+] เลือกประเภทผู้ใช้งาน: {cfg['user_type']}")
    user_type_select = page.locator("select").nth(1)
    user_type_select.wait_for(state="visible", timeout=10_000)
    user_type_select.select_option(label=cfg["user_type"])

    # 2) เลือกวิธี Login: คลิก label ที่ครอบ h5 "DOE e-Service"/"E-Workpermit"
    #    (input[type=radio] name=radio: value=1=DOE e-Service, 2=E-Workpermit)
    method_value = "2" if cfg["method"].lower().replace("-", "") == "eworkpermit" else "1"
    print(f"[+] เลือกวิธี Login: {cfg['method']} (value={method_value})")
    page.evaluate(
        """(v) => {
            const r = document.querySelector(`input[type=radio][name=radio][value="${v}"]`);
            if (r) { r.checked = true; r.dispatchEvent(new Event('change', {bubbles:true})); r.click(); }
        }""",
        method_value,
    )
    page.wait_for_timeout(500)
    # คลิก heading เพื่อ trigger UI ด้วย (กันกรณีฟอร์มไม่อัปเดต)
    page.get_by_role("heading", name=cfg["method"], exact=True).first.click()
    page.wait_for_timeout(800)

    # 3) กรอก username / password — ช่อง input id ขึ้นอยู่กับประเภทผู้ใช้งาน
    print("[+] กรอก username / password")
    # เลือก input ช่องแรกที่ visible จริง ของแต่ละประเภทผู้ใช้:
    #   - ผู้กระทำการแทน/นายจ้าง: #employer_login
    #   - คนต่างด้าว: #username_login (อยู่ใน div#alien)
    user_box = page.locator(
        "#employer_login:visible, #username_login:visible, #agency_login:visible"
    ).first
    user_box.wait_for(state="visible", timeout=15_000)
    user_box.fill(cfg["username"])

    pwd_box = page.locator("#password_login")
    pwd_box.wait_for(state="visible", timeout=10_000)
    pwd_box.fill(cfg["password"])

    # 4) กด "เข้าสู่ระบบ"
    print("[+] กดปุ่มเข้าสู่ระบบ")
    page.locator("#validate_login").click()

    # 5) รอผลลัพธ์: เปลี่ยน URL ออกจาก /Login หรือพบ error
    print("[+] รอผลลัพธ์การ Login ...")
    try:
        page.wait_for_url(lambda url: "/Login" not in url, timeout=30_000)
        print(f"[OK] Login สำเร็จ! URL ปัจจุบัน: {page.url}")
        return True
    except PWTimeoutError:
        # ตรวจหาข้อความ error / captcha
        body_text = page.locator("body").inner_text()
        for kw in ("captcha", "CAPTCHA", "ยืนยันว่าไม่ใช่บอท", "ไม่ถูกต้อง", "ไม่สำเร็จ"):
            if kw.lower() in body_text.lower():
                print(f"[WARN] อาจติด: {kw}")
                break
        print("[ERROR] Login ไม่สำเร็จภายในเวลา 30 วินาที")
        return False


def run(pw: Playwright, cfg: dict) -> int:
    SCREENSHOTS_DIR.mkdir(exist_ok=True)

    browser = pw.chromium.launch(
        headless=cfg["headless"],
        args=["--disable-blink-features=AutomationControlled"],
    )
    context = browser.new_context(
        locale="th-TH",
        timezone_id="Asia/Bangkok",
        viewport={"width": 1366, "height": 800},
    )
    page = context.new_page()

    exit_code = 0
    try:
        ok = login(page, cfg)
        page.screenshot(path=str(SCREENSHOTS_DIR / ("after_login.png")), full_page=True)
        if ok:
            context.storage_state(path=str(STORAGE_STATE))
            print(f"[+] บันทึก session ไว้ที่: {STORAGE_STATE}")
        else:
            exit_code = 1
            if not cfg["headless"]:
                print("[i] รอ 60 วิ ให้คุณแก้ captcha / ตรวจสอบหน้าจอ ...")
                page.wait_for_timeout(60_000)
    finally:
        context.close()
        browser.close()
    return exit_code


def main() -> int:
    cfg = load_config()
    with sync_playwright() as pw:
        return run(pw, cfg)


if __name__ == "__main__":
    sys.exit(main())
