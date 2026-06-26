"""Probe: navigate to receipt list page to inspect selectors."""
import sys
sys.stdout.reconfigure(encoding="utf-8")
from pathlib import Path
from playwright.sync_api import sync_playwright
from scrape_wa import (
    login, _read_login_accounts, _search_request,
    _is_logged_out, TRACKING_URL,
)

ROOT = Path(__file__).parent
accounts = _read_login_accounts(ROOT / "UsernameLogin.xlsx")
key = "raweewan68297@mailforeign.com"
acct = accounts[key]
print(f"Login: {acct['username']} type={acct['type']}")

REQ_NO = "69125200694755"

with sync_playwright() as pw:
    browser = pw.chromium.launch(headless=False,
        args=["--ignore-certificate-errors", "--start-maximized"])
    ctx = browser.new_context(
        locale="th-TH", ignore_https_errors=True,
        viewport={"width": 1920, "height": 1080},
    )
    page = ctx.new_page()
    login(page, {"username": acct["username"], "password": acct["password"],
                 "user_type": acct["type"], "method": acct.get("method") or "E-Workpermit"})
    page.wait_for_timeout(2000)
    page.goto(TRACKING_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(3000)

    print("=== Search request ===")
    _search_request(page, REQ_NO)
    page.wait_for_timeout(2500)

    # Click row to enter detail
    clicked = page.evaluate(r"""() => {
        const a = document.querySelector('a[onclick*="openDetail"]');
        if (a) { a.click(); return true; }
        return false;
    }""")
    print(f"Clicked openDetail: {clicked}")
    page.wait_for_timeout(5000)
    print(f"After detail nav: {page.url}")

    # Click การชำระเงิน tab
    page.evaluate(r"""() => {
        const a = Array.from(document.querySelectorAll('a[href^="#"], .nav-link, .nav a'))
            .find(e => /^การชำระเงิน$/.test((e.textContent||'').trim()));
        if (a) a.click();
    }""")
    page.wait_for_timeout(2500)

    # Inspect ดูใบเสร็จรับเงิน button
    btn_info = page.evaluate(r"""() => {
        const vis = el => { const r = el.getBoundingClientRect(); return r.width>0 && r.height>0; };
        const b = Array.from(document.querySelectorAll('button,a'))
            .find(e => vis(e) && /ดูใบเสร็จรับเงิน/.test((e.textContent||'').trim()));
        if (!b) return null;
        return {
            tag: b.tagName, id: b.id||'', cls: b.className||'',
            onclick: b.getAttribute('onclick')||'',
            href: b.getAttribute('href')||'',
            text: (b.textContent||'').trim().slice(0,80),
        };
    }""")
    print(f"ดูใบเสร็จรับเงิน button: {btn_info}")

    # Click it
    try:
        with page.expect_navigation(timeout=15000, wait_until="domcontentloaded"):
            page.evaluate(r"""() => {
                const b = Array.from(document.querySelectorAll('button,a'))
                    .find(e => /ดูใบเสร็จรับเงิน/.test((e.textContent||'').trim()));
                if (b) b.click();
            }""")
    except Exception as e:
        print(f"Nav exception: {e}")
    page.wait_for_timeout(4000)
    print(f"After ดูใบเสร็จรับเงิน click URL: {page.url}")

    # Dump list page
    info = page.evaluate(r"""() => {
        const rows = Array.from(document.querySelectorAll('table tbody tr'));
        const rowsData = rows.slice(0, 5).map(tr => {
            return Array.from(tr.querySelectorAll('td')).map(td => (td.innerText||'').trim());
        });
        // ปุ่ม pagination
        const pagBtns = Array.from(document.querySelectorAll('.pagination a, .pagination button, [aria-label*="page"]'))
            .map(b => ((b.textContent||'').trim() || b.getAttribute('aria-label') || '').slice(0,30));
        // ปุ่ม "ใบเสร็จรับเงิน" download
        const dlBtns = Array.from(document.querySelectorAll('button,a'))
            .filter(b => /ใบเสร็จรับเงิน/.test((b.textContent||'').trim()))
            .slice(0, 3).map(b => ({
                tag: b.tagName, cls: b.className||'',
                onclick: (b.getAttribute('onclick')||'').slice(0,200),
                href: (b.getAttribute('href')||'').slice(0,200),
                text: (b.textContent||'').trim().slice(0,40),
            }));
        // หาคำว่า "ตรวจสอบข้อมูลรายการชำระเงิน" ปุ่ม
        const refreshBtn = Array.from(document.querySelectorAll('button,a'))
            .find(b => /ตรวจสอบข้อมูลรายการ/.test((b.textContent||'').trim()));
        return {
            url: location.href,
            title: document.title,
            firstRows: rowsData,
            paginationBtns: pagBtns,
            downloadBtns: dlBtns,
            refreshBtn: refreshBtn ? refreshBtn.outerHTML.slice(0,300) : null,
            totalText: (document.body.innerText.match(/รายการคำขอ[^\n]+/)||[''])[0],
        };
    }""")
    import json
    out_data = {"receipt_btn_info": btn_info, "list_info": info}
    Path("_t_probe_receipt_page.json").write_text(
        json.dumps(out_data, ensure_ascii=False, indent=2), encoding="utf-8")
    print("WROTE _t_probe_receipt_page.json")
    page.wait_for_timeout(5000)
    ctx.close(); browser.close()
