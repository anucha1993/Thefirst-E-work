import os, json
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

load_dotenv()
U = os.environ["EWP_USERNAME"]; P = os.environ["EWP_PASSWORD"]

def login(p):
    p.goto("https://eworkpermit.doe.go.th/Login", wait_until="domcontentloaded")
    p.locator("select").nth(1).select_option(label="ผู้กระทำการแทน")
    p.wait_for_timeout(600)
    p.evaluate("""() => {
        const r = document.querySelector('input[type=radio][name=radio][value="2"]');
        if (r) { r.checked = true; r.dispatchEvent(new Event('change',{bubbles:true})); r.click(); }
    }""")
    p.get_by_role("heading", name="E-Workpermit", exact=True).first.click()
    p.wait_for_timeout(600)
    p.locator("#employer_login").fill(U)
    p.locator("#password_login").fill(P)
    p.locator("#validate_login").click()
    p.wait_for_url(lambda u: "/Login" not in u, timeout=30000)

with sync_playwright() as pw:
    b = pw.chromium.launch(headless=False)
    ctx = b.new_context(locale="th-TH", viewport={"width":1500,"height":900})
    p = ctx.new_page()
    login(p)
    p.evaluate("openpageTracking && openpageTracking()")
    p.wait_for_url("**/Permit/Tracking", timeout=20000)
    p.wait_for_timeout(3000)

    info = p.evaluate(r"""() => {
      // Find status pill containing "รอยื่นเอกสารเพิ่มเติม"
      const all = Array.from(document.querySelectorAll('*'));
      const matches = all.filter(e => {
          const t = (e.innerText || '').trim();
          return t && t.length < 80 && t.includes('รอยื่นเอกสารเพิ่มเติม') && e.children.length <= 3;
      }).slice(0, 10);
      return matches.map(e => ({tag: e.tagName, cls: e.className, id: e.id, html: e.outerHTML.slice(0,400)}));
    }""")
    print("รอยื่นเอกสารเพิ่มเติม MATCHES:")
    print(json.dumps(info, indent=2, ensure_ascii=False))

    # Try clicking it
    try:
        p.get_by_text("รอยื่นเอกสารเพิ่มเติม", exact=False).first.click()
        p.wait_for_timeout(2500)
        print("After click - URL:", p.url)
    except Exception as e:
        print("CLICK ERR:", e)

    # Dump first row data structure
    row_info = p.evaluate(r"""() => {
      const rows = Array.from(document.querySelectorAll('table tbody tr'));
      return {
        rowCount: rows.length,
        sample: rows.slice(0,3).map(r => ({html: r.outerHTML.slice(0,1500), text: r.innerText.replace(/\s+/g,' ').slice(0,400)}))
      };
    }""")
    print("ROWS:")
    print(json.dumps(row_info, indent=2, ensure_ascii=False))

    p.screenshot(path="screenshots/tracking_filtered.png", full_page=True)
    input("Press Enter to close...")
    ctx.close(); b.close()
