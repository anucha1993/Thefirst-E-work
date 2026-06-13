import sys, json
from playwright.sync_api import sync_playwright

with sync_playwright() as pw:
    b = pw.chromium.launch(headless=False, args=["--ignore-certificate-errors"])
    ctx = b.new_context(locale="th-TH", ignore_https_errors=True)
    page = ctx.new_page()
    page.goto("https://eworkpermit.doe.go.th/Login/Register", wait_until="domcontentloaded")
    page.wait_for_timeout(3500)
    page.screenshot(path="screenshots/reg_step1.png", full_page=True)
    info = page.evaluate(r"""() => {
      const radios = Array.from(document.querySelectorAll('input[type=radio]')).map(e => {
        let lbl = '';
        if (e.id) { const l = document.querySelector(`label[for="${e.id}"]`); if (l) lbl = l.textContent.trim().slice(0,80); }
        if (!lbl && e.parentElement) lbl = e.parentElement.textContent.trim().slice(0,80);
        return {id: e.id, name: e.name, value: e.value, checked: e.checked, visible: e.offsetParent!==null, lbl};
      });
      const checkboxes = Array.from(document.querySelectorAll('input[type=checkbox]')).map(e => {
        let lbl = '';
        if (e.id) { const l = document.querySelector(`label[for="${e.id}"]`); if (l) lbl = l.textContent.trim().slice(0,80); }
        if (!lbl && e.parentElement) lbl = e.parentElement.textContent.trim().slice(0,80);
        return {id: e.id, name: e.name, value: e.value, lbl, visible: e.offsetParent!==null};
      });
      const inputs = Array.from(document.querySelectorAll('input:not([type=hidden]):not([type=radio]):not([type=checkbox]):not([type=image])')).filter(e => e.offsetParent!==null).map(e => ({id: e.id, name: e.name, type: e.type, placeholder: e.placeholder}));
      const buttons = Array.from(document.querySelectorAll('button, a.btn')).filter(e => e.offsetParent !== null).map(e => ({tag: e.tagName, id: e.id, cls: e.className.slice(0,80), txt: (e.textContent||'').trim().slice(0,60), onclick: e.getAttribute('onclick')}));
      const sects = Array.from(document.querySelectorAll('h1,h2,h3,h4,h5,legend')).filter(e => e.offsetParent !== null).map(e => ({tag: e.tagName, txt: (e.textContent||'').trim().slice(0,80)}));
      return {url: location.href, radios, checkboxes, inputs, buttons: buttons.slice(0,30), sections: sects};
    }""")
    print(json.dumps(info, indent=2, ensure_ascii=False))
    ctx.close(); b.close()
