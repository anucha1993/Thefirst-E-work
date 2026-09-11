"""ดึงตัวเลือกทั้งหมดใน dropdown #Filter_request_list (รายการคำขอ) จากเว็บจริง
เพื่อหา code ของตัวเลือกใหม่ (เช่น 'ตามมติ ครม. 14 กรกฎาคม 2569').
เปิดเบราว์เซอร์จริง — ต้องแก้ reCAPTCHA เอง (~120 วิ) อย่าปิดเบราว์เซอร์กลางคัน.
รัน: python _probe_request_list.py
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
from scrape_wa import login, goto_tracking, ensure_session, _is_logged_out, ROOT  # noqa: E402

DUMP_JS = r"""() => {
    const sel = document.getElementById('Filter_request_list');
    if (!sel) return null;
    return Array.from(sel.options).map(o => ({ val: o.value, text: (o.textContent || '').replace(/\s+/g,' ').trim() }));
}"""


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
        "headless": False,
        "login_timeout_ms": 120_000,
    }
    if not cfg["username"] or not cfg["password"]:
        print("ERR: EWP_USERNAME/EWP_PASSWORD ไม่ถูกตั้งใน .env")
        return 1

    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, args=["--start-maximized"])
        ctx = b.new_context(locale="th-TH", viewport={"width": 1500, "height": 900})
        page = ctx.new_page()
        try:
            print("[1] login...")
            login(page, cfg)
            print(f"[+] login OK ({page.url})")
            goto_tracking(page)
            if _is_logged_out(page):
                ensure_session(page, cfg)
                goto_tracking(page)
            page.wait_for_timeout(2500)

            opts = page.evaluate(DUMP_JS)
            if not opts:
                print("[!] ไม่พบ #Filter_request_list บนหน้า")
                return 2

            print(f"\n[i] พบ {len(opts)} ตัวเลือกใน dropdown 'รายการคำขอ':\n")
            for o in opts:
                print(f"    {o['val']!r:38} = {o['text']}")

            outp = ROOT / "_probe_request_list_out.json"
            outp.write_text(json.dumps(opts, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"\n[i] เขียนไฟล์: {outp.name}")

            # ไฮไลต์ตัวเลือกใหม่ที่น่าจะเป็น 14 ก.ค. 2569
            print("\n[NEW?] ตัวเลือกที่มี '2569' / '14 กรกฎาคม':")
            hit = False
            for o in opts:
                if "2569" in o["text"] or "14 กรกฎาคม" in o["text"] or "14 ก.ค" in o["text"]:
                    print(f"    >>> code = {o['val']!r}\n        text = {o['text']}")
                    hit = True
            if not hit:
                print("    (ไม่พบข้อความที่มี 2569/14 กรกฎาคม — ดูรายการเต็มด้านบน)")
        finally:
            page.wait_for_timeout(500)
            b.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
