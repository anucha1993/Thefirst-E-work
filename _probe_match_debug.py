"""Debug: รันโค้ด matching JS ตัวเดียวกับใน scrape_wa.py ตรงๆ เพื่อดูว่าทำไม NOT_FOUND
ใช้: python _probe_match_debug.py 69125300074621
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

PATTERNS = {
    "result_notice": r"ใบแจ้งผลใบอนุญาตทำงาน",
    "request_receipt": r"ใบรับคำขอใบอนุญาตทำงาน",
    "bt50": r"บต\.?\s*50",
}


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

            for name, pattern in PATTERNS.items():
                result = page.evaluate(
                    r"""(pat) => {
                        const re = new RegExp(pat);
                        const pane = document.querySelector('#tab-response') || document;
                        const links = [...pane.querySelectorAll('[onclick*="GetDocumentConfirm"]')];
                        const debugInfo = [];
                        const matches = [];
                        for (const a of links) {
                            let node = a.parentElement, text = '';
                            let steps = 0;
                            for (let i = 0; i < 8 && node; i++) {
                                text = (node.innerText || '').trim();
                                steps = i;
                                if (text && re.test(text)) break;
                                if ((node.className || '').includes('col-12')) break;
                                node = node.parentElement;
                            }
                            const isMatch = !!(text && re.test(text));
                            debugInfo.push({onclick: a.getAttribute('onclick'), steps, isMatch, textHead: text.slice(0,60)});
                            if (isMatch) matches.push(text);
                        }
                        return {matchesCount: matches.length, debugInfo, linksTotal: links.length};
                    }""",
                    pattern,
                )
                print(f"\n=== {name} (pattern={pattern!r}) ===")
                print(f"  linksTotal={result['linksTotal']}  matchesCount={result['matchesCount']}")
                for d in result["debugInfo"]:
                    print(f"    onclick={d['onclick']}  steps={d['steps']}  isMatch={d['isMatch']}  text={d['textHead']!r}")

            page.wait_for_timeout(2000)
        finally:
            ctx.close()
            b.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
