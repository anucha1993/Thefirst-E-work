"""Probe: List actual options in BT.44 search modal selects.
รัน: .\.venv\Scripts\python.exe _probe_bt44_options.py
"""
import json
from playwright.sync_api import sync_playwright
from scrape_wa import load_config, login

cfg = load_config(require_login=True)

with sync_playwright() as pw:
    b = pw.chromium.launch(headless=False, args=["--ignore-certificate-errors", "--start-maximized"])
    ctx = b.new_context(locale="th-TH", ignore_https_errors=True,
                        viewport={"width": 1920, "height": 1080})
    page = ctx.new_page()
    login(page, cfg)
    page.wait_for_timeout(1500)

    page.goto("https://eworkpermit.doe.go.th/", wait_until="domcontentloaded", timeout=30_000)
    page.wait_for_timeout(2000)

    page.evaluate(r"""() => { document.querySelector('a.lang_menu_service')?.click(); }""")
    page.wait_for_timeout(1200)

    page.evaluate(r"""() => { try { if (typeof openCity === 'function') openCity('tab_CHANGE_REQ', new Event('click')); } catch (e) {} }""")
    page.wait_for_timeout(900)

    page.evaluate(r"""() => { document.querySelector('#CHANGE_EMPLOYER')?.click(); }""")
    try:
        page.wait_for_url("**/WorkPermit**", timeout=30_000)
    except:
        pass
    page.wait_for_timeout(2000)

    page.evaluate(r"""() => {
        const b = Array.from(document.querySelectorAll('button,a'))
          .find(e => /search_alien_modal\.show/.test(e.getAttribute('onclick') || ''));
        if (b) b.click();
    }""")
    page.wait_for_timeout(1500)

    options = page.evaluate(r"""() => {
        const result = {
            alien_prefix: [],
            nationality_al: [],
            sexCheck: []
        };
        for (const [id, opts] of Object.entries(result)) {
            const sel = document.getElementById(id);
            if (sel) {
                opts.push(...Array.from(sel.options).map(o => ({value: o.value, text: o.text.trim()})));
            }
        }
        return result;
    }""")
    print("\n=== ACTUAL SELECT OPTIONS ===")
    print(json.dumps(options, ensure_ascii=False, indent=2))

    print("\n--- เปิดค้างไว้ให้ตรวจสอบ ---")
    input("กด Enter เพื่อปิด...")
    ctx.close(); b.close()
