import sys, json
from playwright.sync_api import sync_playwright

def dump(page, label):
    info = page.evaluate(r"""() => {
      const visible = e => e.offsetParent !== null && getComputedStyle(e).visibility !== 'hidden';
      const radios = Array.from(document.querySelectorAll('input[type=radio]')).filter(visible).map(e => {
        let lbl = '';
        if (e.id) { const l = document.querySelector(`label[for="${e.id}"]`); if (l) lbl = l.textContent.trim().slice(0,80); }
        if (!lbl && e.parentElement) lbl = e.parentElement.textContent.trim().slice(0,80);
        return {id: e.id, name: e.name, value: e.value, checked: e.checked, lbl};
      });
      const checkboxes = Array.from(document.querySelectorAll('input[type=checkbox]')).filter(visible).map(e => {
        let lbl = ''; if (e.id) { const l = document.querySelector(`label[for="${e.id}"]`); if (l) lbl = l.textContent.trim().slice(0,80); }
        if (!lbl && e.parentElement) lbl = e.parentElement.textContent.trim().slice(0,80);
        return {id: e.id, name: e.name, lbl};
      });
      const inputs = Array.from(document.querySelectorAll('input:not([type=hidden]):not([type=radio]):not([type=checkbox]):not([type=image])')).filter(visible).map(e => ({id: e.id, name: e.name, type: e.type, placeholder: e.placeholder}));
      const selects = Array.from(document.querySelectorAll('select')).filter(visible).map(e => ({id: e.id, name: e.name, opts: Array.from(e.options).map(o => `${o.value}::${o.text.trim()}`)}));
      const buttons = Array.from(document.querySelectorAll('button, a.btn')).filter(visible).map(e => ({tag: e.tagName, id: e.id, txt: (e.textContent||'').trim().slice(0,60), onclick: e.getAttribute('onclick')})).slice(0,20);
      return {radios, checkboxes, inputs, selects, buttons};
    }""")
    print(f"\n=== {label} ===\n" + json.dumps(info, indent=2, ensure_ascii=False))

with sync_playwright() as pw:
    b = pw.chromium.launch(headless=False, args=["--ignore-certificate-errors"])
    ctx = b.new_context(locale="th-TH", ignore_https_errors=True)
    page = ctx.new_page()
    page.goto("https://eworkpermit.doe.go.th/Login/Register", wait_until="domcontentloaded")
    page.wait_for_timeout(3000)
    dump(page, "STEP A: initial")

    # 1.1 เลือกประเภทผู้ใช้งาน "คนต่างด้าว" (value=1)
    page.locator("#User_typeRegister").select_option(value="1")
    page.wait_for_timeout(1500)
    page.screenshot(path="screenshots/reg_step2.png", full_page=True)
    dump(page, "STEP B: after คนต่างด้าว")

    # 1.2 เลือก "มีใบอนุญาตทำงานแล้ว"
    page.locator("#alienOption1").check(force=True)
    page.wait_for_timeout(1200)
    page.screenshot(path="screenshots/reg_step3.png", full_page=True)
    dump(page, "STEP C: after มีใบอนุญาต")

    ctx.close(); b.close()
