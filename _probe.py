import os, json
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

load_dotenv()
U = os.environ["EWP_USERNAME"]; P = os.environ["EWP_PASSWORD"]

with sync_playwright() as pw:
    b = pw.chromium.launch(headless=False)
    ctx = b.new_context(locale="th-TH")
    p = ctx.new_page()
    p.goto("https://eworkpermit.doe.go.th/Login", wait_until="domcontentloaded")
    p.locator("select").nth(1).select_option(label="ผู้กระทำการแทน")
    p.wait_for_timeout(800)
    p.evaluate("""() => {
        const r = document.querySelector('input[type=radio][name=radio][value="2"]');
        if (r) { r.checked = true; r.dispatchEvent(new Event('change', {bubbles:true})); r.click(); }
    }""")
    p.get_by_role("heading", name="E-Workpermit", exact=True).first.click()
    p.wait_for_timeout(800)
    p.locator("#employer_login").fill(U)
    p.locator("#password_login").fill(P)
    p.locator("#validate_login").click()
    p.wait_for_url(lambda u: "/Login" not in u, timeout=30000)
    print("LOGGED IN URL:", p.url)
    ctx.storage_state(path="storage_state.json")
    # ตอนนี้คลิก e-Tracking
    p.evaluate("openpageTracking && openpageTracking()")
    p.wait_for_timeout(4000)
    print("AFTER e-Tracking URL:", p.url)
    p.screenshot(path="screenshots/etracking.png", full_page=True)
    # dump tabs/filters info
    info = p.evaluate("""() => {
        const collect = (sel) => Array.from(document.querySelectorAll(sel)).map(e => e.innerText && e.innerText.trim().slice(0,120)).filter(Boolean);
        return {
            title: document.title,
            tabs: collect('[role=tab], .nav-link, .nav-tabs a'),
            statusChips: collect('.badge, .pill, .chip, .status-tag').slice(0,40),
            buttons: collect('button').slice(0,30),
            headings: collect('h1,h2,h3,h4,h5').slice(0,20),
        };
    }""")
    print(json.dumps(info, indent=2, ensure_ascii=False))
    input("Press Enter to close...")
    ctx.close(); b.close()
