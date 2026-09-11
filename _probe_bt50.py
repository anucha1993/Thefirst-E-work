"""ทดสอบดาวน์โหลดเอกสาร 'แบบ บต.50 อ.6' (BT50) จากคำขอจริง โดยเรียกใช้ฟังก์ชันจริงใน scrape_wa
เปิดเบราว์เซอร์จริง — ต้องแก้ reCAPTCHA เอง (~120 วิ) อย่าปิดเบราว์เซอร์กลางคัน.
รัน: python _probe_bt50.py [เลขคำขอ]
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
    login, goto_tracking, ensure_session, _is_logged_out,
    _search_request, _open_first_detail, _extract_alien_eng_name,
    _download_doc_pdf, RESULT_DOC_TYPES, ROOT,
)

REQ_NO = sys.argv[1] if len(sys.argv) > 1 else "69125200578664"


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
    print(f"[i] บัญชี: {cfg['username']} ({cfg['user_type']}) | คำขอ: {REQ_NO}")

    out_pdf = ROOT / "_probe_bt50_out.pdf"
    if out_pdf.exists():
        out_pdf.unlink()

    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, args=["--start-maximized"])
        ctx = b.new_context(locale="th-TH", viewport={"width": 1500, "height": 900},
                            ignore_https_errors=True, accept_downloads=True)
        page = ctx.new_page()
        try:
            print("[1] login... (แก้ reCAPTCHA ในเบราว์เซอร์)")
            login(page, cfg)
            print(f"[+] login OK ({page.url[:70]})")

            goto_tracking(page)
            if _is_logged_out(page):
                ensure_session(page, cfg)
                goto_tracking(page)
            page.wait_for_timeout(1800)

            print(f"[2] ค้นหาคำขอ {REQ_NO} ...")
            _search_request(page, REQ_NO)
            if not _open_first_detail(page):
                print("[!] ไม่พบผลค้นหา / เปิดรายละเอียดไม่ได้")
                return 2
            name_eng = _extract_alien_eng_name(page) or "-"
            print(f"[+] เปิด detail แล้ว | ชื่อคนต่างด้าว: {name_eng}")

            _bt50 = RESULT_DOC_TYPES["bt50"]
            _dl_cfg = {"label": _bt50["label"], "tab_pattern": r"เอกสารตอบรับ",
                       "find": "row_link", "label_pattern": _bt50["label_pattern"]}
            print("[3] ดาวน์โหลด บต.50 (row_link, label_pattern=%r) ..."
                  % _bt50["label_pattern"])
            err = _download_doc_pdf(page, _dl_cfg, out_pdf, log=print)
            if err:
                print(f"[✗] FAIL: {err}")
                return 3
            size = out_pdf.stat().st_size if out_pdf.exists() else 0
            head = out_pdf.read_bytes()[:5] if out_pdf.exists() else b""
            print(f"[✓] SUCCESS: {out_pdf.name} ({size // 1024} KB, header={head!r})")
            return 0
        finally:
            page.wait_for_timeout(1200)
            ctx.close(); b.close()


if __name__ == "__main__":
    raise SystemExit(main())
