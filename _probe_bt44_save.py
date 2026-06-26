"""Probe: BT.44 บันทึก → analyze modal detection.
รัน: .\.venv\Scripts\python.exe _probe_bt44_save.py
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

    # 1) Navigate to BT.44 form
    page.evaluate(r"""() => {
        const a = document.querySelector('a.lang_menu_service');
        if (a) a.click();
    }""")
    page.wait_for_timeout(1200)

    page.evaluate(r"""() => {
        try { if (typeof openCity === 'function') openCity('tab_CHANGE_REQ', new Event('click')); } catch (e) {}
    }""")
    page.wait_for_timeout(900)

    page.evaluate(r"""() => {
        const t = document.querySelector('#CHANGE_EMPLOYER');
        if (t) t.click();
    }""")
    try:
        page.wait_for_url("**/WorkPermit**", timeout=30_000)
    except:
        pass
    page.wait_for_timeout(2000)

    # 2) Open search modal
    page.evaluate(r"""() => {
        const b = Array.from(document.querySelectorAll('button,a'))
          .find(e => /search_alien_modal\.show/.test(e.getAttribute('onclick') || ''));
        if (b) b.click();
    }""")
    page.wait_for_timeout(1500)

    # 3) Fill form with test data (EI EI THWIN)
    page.evaluate(r"""() => {
        document.getElementById('alien_id').value = '';
        document.getElementById('alien_prefix').value = 'นางสาว';
        document.getElementById('other_name').value = 'EI EI THWIN';
        document.getElementById('nationality_al').value = 'เมียนมา';
        document.getElementById('sexCheck').value = 'หญิง';
        document.getElementById('birthDateCheck').value = '25/11/2005';
    }""")
    page.wait_for_timeout(500)
    page.screenshot(path="screenshots/bt44_probe_filled.png", full_page=True)

    # 4) Click บันทึก
    print("\n[*] กดบันทึก...")
    page.evaluate(r"""() => {
        const b = document.getElementById('btn_search_alien_submit');
        if (b) b.click();
    }""")

    # 5) Wait in steps and capture state
    for step in [800, 1500, 2800, 3500, 4200, 5000]:
        page.wait_for_timeout(step - (step if step == 800 else (step - [800, 1500, 2800, 3500, 4200][['800', '1500', '2800', '3500', '4200', '5000'].index(str(step))-1] if step > 800 else 0)))
        state = page.evaluate(r"""() => {
            // Check if spinner is showing
            const spinner = Array.from(document.querySelectorAll('div,span'))
              .find(e => /spinner|loading|โหลด/.test((e.className||'')+(e.style.display||'')) && e.offsetParent !== null);
            
            // Check modal visibility
            const modal = document.querySelector('div[id="search_alien_modal"],div[class*="modal"][class*="show"]');
            const modal_open = modal && modal.offsetParent !== null;
            
            // Check submit button visibility
            const submit_btn = document.getElementById('btn_search_alien_submit');
            const submit_visible = submit_btn && submit_btn.offsetParent !== null;
            
            // Try to detect alert/success message (SweetAlert or Bootstrap modal)
            let alert_text = '';
            for (const sel of ['.swal2-popup', '.swal2-container', '.modal.show', '.modal[style*="display: block"]']) {
              const el = document.querySelector(sel);
              if (el && el.offsetParent !== null) {
                const title = el.querySelector('.swal2-title, .modal-title');
                const body = el.querySelector('.swal2-html-container, .swal2-content, .modal-body');
                alert_text = ((title?.textContent||'')+(body?.textContent||'')).trim().slice(0,200);
                break;
              }
            }
            
            // Check if table has new rows (aliens table)
            const alienTable = document.querySelector('table');  
            const rowCount = alienTable ? alienTable.querySelectorAll('tbody tr').length : 0;
            
            return {
              spinner_active: !!spinner,
              modal_open,
              submit_visible,
              alert_text,
              row_count: rowCount
            };
        }""")
        print(f"\n[+] At {step}ms:")
        print(json.dumps(state, ensure_ascii=False, indent=2))
        page.screenshot(path=f"screenshots/bt44_probe_{step}ms.png", full_page=True)

    print("\n--- เปิดค้างไว้ให้ตรวจสอบ ---")
    input("กด Enter เพื่อปิด...")
    ctx.close(); b.close()
