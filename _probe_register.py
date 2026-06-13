import sys, json
from playwright.sync_api import sync_playwright

with sync_playwright() as pw:
    b = pw.chromium.launch(headless=False, args=["--ignore-certificate-errors"])
    ctx = b.new_context(locale="th-TH", ignore_https_errors=True)
    page = ctx.new_page()
    page.goto("https://eworkpermit.doe.go.th/", wait_until="domcontentloaded")
    page.wait_for_timeout(3000)
    # หาปุ่มลงทะเบียน
    info = page.evaluate(r"""() => {
      const bts = Array.from(document.querySelectorAll('a, button')).filter(e => /ลงทะเบียน/.test((e.textContent||'').trim())).map(e => ({tag: e.tagName, id: e.id, cls: e.className, txt: (e.textContent||'').trim().slice(0,50), href: e.href, onclick: e.getAttribute('onclick')}));
      return {url: location.href, registerBtns: bts.slice(0, 10)};
    }""")
    print("=== HOME ==="); print(json.dumps(info, indent=2, ensure_ascii=False))
    # คลิกปุ่มลงทะเบียนตัวแรก
    if info["registerBtns"]:
        try:
            page.locator("a, button").filter(has_text="ลงทะเบียน").first.click(timeout=5000)
            page.wait_for_timeout(3000)
        except Exception as e:
            print("click err:", e)
    print("=== AFTER CLICK ==="); print("URL:", page.url)
    page.screenshot(path="screenshots/register_step1.png", full_page=True)
    info2 = page.evaluate(r"""() => {
      const collect = (sel) => Array.from(document.querySelectorAll(sel)).slice(0, 30).map(e => ({tag: e.tagName, id: e.id, name: e.name, type: e.type, value: e.value, txt: (e.labels && e.labels[0] ? e.labels[0].textContent.trim() : (e.parentElement ? e.parentElement.textContent.trim().slice(0,80) : '')).slice(0,80)}));
      return {
        radios: collect('input[type=radio]'),
        checkboxes: collect('input[type=checkbox]'),
        inputs: Array.from(document.querySelectorAll('input:not([type=hidden]):not([type=radio]):not([type=checkbox])')).slice(0, 30).map(e => ({id: e.id, name: e.name, type: e.type, placeholder: e.placeholder, label: e.previousElementSibling ? e.previousElementSibling.textContent.trim().slice(0,80) : ''})),
        selects: Array.from(document.querySelectorAll('select')).slice(0, 20).map(e => ({id: e.id, name: e.name, opts: Array.from(e.options).map(o => o.text.trim().slice(0,40))})),
        buttons: Array.from(document.querySelectorAll('button, a.btn')).slice(0, 30).map(e => ({tag: e.tagName, id: e.id, cls: e.className, txt: (e.textContent||'').trim().slice(0,60)})),
        headings: Array.from(document.querySelectorAll('h1,h2,h3,h4,h5,legend,label')).slice(0, 30).map(e => ({tag: e.tagName, txt: (e.textContent||'').trim().slice(0,80)}))
      };
    }""")
    print(json.dumps(info2, indent=2, ensure_ascii=False))
    input("Press Enter to close...")
    ctx.close(); b.close()
