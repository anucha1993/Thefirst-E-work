"""รันดาวน์โหลดเอกสารผลอนุญาต (ใบแจ้งผล/ใบรับคำขอ/บต.50) สำหรับ 1 เลขคำขอ โดยตรง
(กดสร้างเอกสารอัตโนมัติถ้ายังไม่เคยสร้าง แล้ว poll รอจนพร้อมดาวน์โหลด)
ใช้: python _run_result_docs_one.py 69125300074621
ไฟล์ที่ได้จะอยู่ใน reports/result_docs_manual/
"""
from __future__ import annotations
import os
import sys
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import (  # noqa: E402
    login, goto_tracking, _is_logged_out,
    _search_request, _open_first_detail, _extract_alien_eng_name,
    _download_response_doc_named, RESULT_DOC_TYPES, _receipt_safe_name,
)

OUT_DIR = Path(__file__).parent / "reports" / "result_docs_manual"
DOC_KEYS = ["result_notice", "request_receipt", "bt50"]


def main() -> int:
    req_no = sys.argv[1] if len(sys.argv) > 1 else "69125300074621"
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }
    if not cfg["username"] or not cfg["password"]:
        print("ERR: EWP_USERNAME/EWP_PASSWORD ไม่ถูกตั้งใน .env")
        return 1
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"[i] บัญชี: {cfg['username']} | เลขคำขอ: {req_no}")
    print(f"[i] เอกสาร: {', '.join(RESULT_DOC_TYPES[d]['label'] for d in DOC_KEYS)}")

    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, args=["--start-maximized"])
        ctx = b.new_context(locale="th-TH", viewport={"width": 1500, "height": 900},
                            ignore_https_errors=True, accept_downloads=True)
        page = ctx.new_page()
        try:
            print("[1] login... (แก้ reCAPTCHA)")
            login(page, cfg)
            print(f"[+] login OK ({page.url[:70]})")
            goto_tracking(page)
            page.wait_for_timeout(1500)
            if _is_logged_out(page):
                print("[!] session หลุดหลัง login — หยุด")
                return 1

            _search_request(page, req_no)
            if not _open_first_detail(page):
                print("[!] เปิด detail ไม่สำเร็จ (ไม่พบผลค้นหา)")
                return 1
            print(f"[+] เปิด detail สำเร็จ: {page.url}")

            name_eng = _extract_alien_eng_name(page) or req_no
            name_safe = _receipt_safe_name(name_eng)
            print(f"[+] ชื่อคนต่างด้าว(Eng): {name_eng!r}")

            for dk in DOC_KEYS:
                print(f"\n--- {RESULT_DOC_TYPES[dk]['label']} ---")
                res = _download_response_doc_named(
                    page, OUT_DIR, name_safe, RESULT_DOC_TYPES[dk], log=print, req_no=req_no,
                )
                print(f"    => {res}")

            print(f"\n[เสร็จ] ไฟล์ที่ได้อยู่ใน: {OUT_DIR}")
            page.wait_for_timeout(3000)
        finally:
            ctx.close()
            b.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
