"""Debug: สำหรับแต่ละลิงก์ GetDocumentConfirm บนหน้า detail ของเลขคำขอที่กำหนด
dump innerText ของ parentElement ที่ระดับ 0-10 ชั้น เพื่อดูว่า climb ไปเจอ label ตรงไหน
ใช้: python _probe_climb_debug.py 69125300074621
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
    _search_request, _open_first_detail,
)


def main() -> int:
    req_no = sys.argv[1] if len(sys.argv) > 1 else "69125300074621"
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }
    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, args=["--start-maximized"])
        ctx = b.new_context(locale="th-TH", viewport={"width": 1500, "height": 900},
                            ignore_https_errors=True, accept_downloads=True)
        page = ctx.new_page()
        try:
            login(page, cfg)
            goto_tracking(page)
            page.wait_for_timeout(1200)
            if _is_logged_out(page):
                print("logged out"); return 1
            _search_request(page, req_no)
            if not _open_first_detail(page):
                print("open detail failed"); return 1
            page.evaluate(
                r"""() => { const f=[...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a')]
                    .find(a=>/เอกสารตอบรับ/.test(a.innerText||'')); if(f) f.click(); }"""
            )
            page.wait_for_timeout(1500)
            dump = page.evaluate(
                r"""() => {
                    const links = [...document.querySelectorAll('[onclick*="GetDocumentConfirm"]')];
                    return links.map((a, idx) => {
                        const levels = [];
                        let node = a.parentElement;
                        for (let i = 0; i < 10 && node; i++) {
                            levels.push({
                                i,
                                tag: node.tagName.toLowerCase(),
                                cls: node.className || '',
                                textLen: (node.innerText||'').trim().length,
                                textHead: (node.innerText||'').trim().replace(/\s+/g,' ').slice(0,80),
                            });
                            node = node.parentElement;
                        }
                        return {idx, onclick: a.getAttribute('onclick'), levels};
                    });
                }"""
            )
            import json
            print(json.dumps(dump, ensure_ascii=False, indent=2))
            page.wait_for_timeout(2000)
        finally:
            ctx.close()
            b.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
