import json
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

wb = load_workbook("Exm.xlsx", data_only=True); ws = wb.active
hdr = [c.value for c in ws[1]]; r = {hdr[i]: c.value for i, c in enumerate(ws[2])}
TAX = str(r.get("TaxID") or "")

with sync_playwright() as pw:
    b = pw.chromium.launch(headless=False, args=["--ignore-certificate-errors"])
    ctx = b.new_context(locale="th-TH", ignore_https_errors=True)
    page = ctx.new_page()
    page.goto("https://eworkpermit.doe.go.th/Login/Register", wait_until="domcontentloaded"); page.wait_for_timeout(2500)
    page.locator("#User_typeRegister").select_option(value="1"); page.wait_for_timeout(500)
    page.locator("input#alienOption1").check(force=True); page.wait_for_timeout(300)
    page.locator("#typeAlienUser").select_option(value="1"); page.wait_for_timeout(300)
    page.locator("#inputAlien1").fill(TAX); page.wait_for_timeout(200)
    page.locator("#validate_register").click(); page.wait_for_timeout(3500)

    # ดู ALL checkboxes (รวม hidden) และ scroll containers
    info = page.evaluate(r"""() => {
      const all_cb = Array.from(document.querySelectorAll('input[type=checkbox]')).map(e => ({id:e.id, name:e.name, cls:e.className, vis: e.offsetParent !== null, visParent: e.parentElement ? getComputedStyle(e.parentElement).visibility : ''}));
      // any element with onscroll or classes hint
      const candidates = Array.from(document.querySelectorAll('*')).filter(e => {
        const cs = getComputedStyle(e);
        return (cs.overflowY === 'scroll' || cs.overflowY === 'auto') && e.scrollHeight > e.clientHeight + 30 && e.offsetParent !== null;
      }).map(e => ({tag:e.tagName, id:e.id, cls:e.className.slice(0,80), sh:e.scrollHeight, ch:e.clientHeight}));
      // search source for navigateConsent / alien-term hints
      const scripts = Array.from(document.querySelectorAll('script:not([src])')).map(s => s.textContent).join('\n');
      const m = scripts.match(/alien-term-accept-btn|navigateConsent|disabled\s*=\s*false/g) || [];
      // also dump first 5 inline scripts that mention term
      const inline = scripts.split('\n').filter(l => /term|consent|disabled/i.test(l)).slice(0,40);
      return {all_cb, scrollables: candidates, hits: m, inline};
    }""")
    print(json.dumps(info, ensure_ascii=False, indent=2))

    # Try scrolling each scrollable to bottom
    page.evaluate(r"""() => {
      Array.from(document.querySelectorAll('*')).forEach(e => {
        const cs = getComputedStyle(e);
        if ((cs.overflowY === 'scroll' || cs.overflowY === 'auto') && e.scrollHeight > e.clientHeight + 30) {
          e.scrollTop = e.scrollHeight;
        }
      });
      window.scrollTo(0, document.body.scrollHeight);
    }""")
    page.wait_for_timeout(1500)
    enabled = page.evaluate("() => { const b=document.getElementById('gonext_0'); return b ? b.disabled : null; }")
    print("gonext_0 disabled after scroll:", enabled)
    page.screenshot(path="screenshots/reg_term_scrolled.png", full_page=True)

    if enabled:
        # try checking any hidden checkbox
        print("Trying to click each checkbox...")
        cbs = page.locator("input[type=checkbox]").all()
        for cb in cbs:
            try:
                cb.evaluate("e => e.click()")
            except Exception as ex:
                print("cb click err:", ex)
        page.wait_for_timeout(800)
        enabled = page.evaluate("() => document.getElementById('gonext_0').disabled")
        print("gonext_0 disabled after cb click:", enabled)

    ctx.close(); b.close()
