"""Probe step-2 (สาขา + รายชื่อแรงงาน) ของฟอร์ม INFORM"""
import json, sys
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
import scrape_wa as s

LOG = Path("_step2.log")
LOG.write_text("", encoding="utf-8")

def w(m):
    print(m)
    with LOG.open("a", encoding="utf-8") as f:
        f.write(str(m) + "\n")

cfg = s.load_config()
accts = s._read_login_accounts(Path("UsernameLogin.xlsx"))
acct = list(accts.values())[0]
cfg["username"] = acct["username"]; cfg["password"] = acct["password"]
cfg["user_type"] = acct["type"]; cfg["method"] = acct["method"] or "E-Service"

recs = s._read_inform_data(Path("FormRequestEmployment.xlsx"))
rec = recs[0]

with sync_playwright() as pw:
    browser = pw.chromium.launch(headless=False, args=["--ignore-certificate-errors","--start-maximized"])
    ctx = browser.new_context(locale="th-TH", ignore_https_errors=True,
                              viewport={"width":1920,"height":1080})
    page = ctx.new_page()
    s.login(page, cfg)
    page.wait_for_timeout(2000)

    s._open_inform_form(page, log=w)
    s._inform_search_employer(page, rec["emp_type"], rec["emp_id"], log=w)
    s._inform_consent_first(page, log=w)
    s._inform_click_next_step1(page, log=w)
    page.wait_for_timeout(3000)
    w(f"Step 2 URL: {page.url}")

    # dump all selects on step 2
    info = page.evaluate(r"""() => {
        const sels = Array.from(document.querySelectorAll('select')).map(s => ({
            id: s.id, name: s.name, visible: !!(s.offsetParent),
            options: Array.from(s.options).slice(0,15).map(o => ({v:o.value, t:(o.textContent||'').trim().slice(0,80)})),
            count: s.options.length,
        }));
        // find buttons with 'เลือกสาขา' or 'สาขา'
        const btns = Array.from(document.querySelectorAll('button,a.btn,.btn'))
            .filter(b => /เลือกสาขา|สาขา|รีเฟรช/.test((b.textContent||'').trim()))
            .map(b => ({text:(b.textContent||'').trim(), id:b.id, cls:b.className}))
            .slice(0,20);
        // headings around step 2
        const headings = Array.from(document.querySelectorAll('h1,h2,h3,h4,h5,h6,strong,.fw-bold'))
            .map(h => (h.textContent||'').trim().slice(0,100))
            .filter(t => t && t.length > 5)
            .slice(0,30);
        return {selects: sels, branchBtns: btns, headings};
    }""")
    w("=== STEP 2 DOM ===")
    w(json.dumps(info, ensure_ascii=False, indent=2))

    # Also check tables
    tables = page.evaluate(r"""() => {
        return Array.from(document.querySelectorAll('table')).map(t => {
            const headers = Array.from(t.querySelectorAll('thead th, thead td')).map(h => (h.textContent||'').trim().slice(0,60));
            const firstRow = (t.querySelectorAll('tbody tr')[0])
                ? Array.from(t.querySelectorAll('tbody tr')[0].querySelectorAll('td')).map(d => (d.textContent||'').trim().slice(0,80))
                : [];
            return {headers, firstRow, rowCount: t.querySelectorAll('tbody tr').length};
        });
    }""")
    w("=== TABLES on step 2 ===")
    w(json.dumps(tables, ensure_ascii=False, indent=2))

    page.wait_for_timeout(30_000)
    ctx.close(); browser.close()
