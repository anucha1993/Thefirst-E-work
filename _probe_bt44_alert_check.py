#!/usr/bin/env python3
"""Probe: ตรวจสอบว่า _capture_register_alert จับอะไร หลังกด บันทึก """
import sys
from playwright.sync_api import sync_playwright

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            args=["--ignore-certificate-errors", "--start-maximized"]
        )
        context = browser.new_context(locale="th-TH", viewport={"width": 1920, "height": 1080}, ignore_https_errors=True)
        page = context.new_page()

        try:
            # Login + nav to form
            from scrape_wa import login
            import dotenv
            cfg = dotenv.dotenv_values(".env")
            login(page, cfg)
            page.goto("https://eworkpermit.doe.go.th/", wait_until="domcontentloaded")
            page.wait_for_timeout(2000)
            
            # Open menu → CHANGE_REQ
            page.locator("a.lang_menu_service").click()
            page.wait_for_timeout(1500)
            page.evaluate(r"""() => window.openCity('tab_CHANGE_REQ')""")
            page.wait_for_timeout(2000)
            
            # Click CHANGE_EMPLOYER
            page.locator("#CHANGE_EMPLOYER").click()
            page.wait_for_url("**/WorkPermit**", timeout=5000)
            page.wait_for_timeout(1000)
            
            # Fill & submit search
            page.locator("#alien_prefix").select_option("น")  # นาง
            page.locator("#alien_id").fill("")
            page.locator("#other_name").fill("TEST NAME")
            page.locator("#birthDateCheck").fill("25/11/2005")
            page.locator("#nationality_al").select_option("BU")  # Myanmar
            page.locator("#sexCheck").select_option("F")  # Female
            
            page.screenshot(path="probe_before_submit.png", full_page=True)
            
            # Click บันทึก
            page.evaluate(r"""() => {
                const b = document.getElementById('btn_search_alien_submit');
                if (b) b.click();
            }""")
            
            # Wait at various checkpoints & check what alert captures
            prev_t = 0
            for t in [500, 1000, 1500, 2000, 3000, 4000, 5000]:
                page.wait_for_timeout(t - prev_t)
                prev_t = t
                
                # Capture using exact code from scrape_wa
                alert_text = ""
                try:
                    alert_text = page.evaluate(r"""() => {
                      for (const el of document.querySelectorAll('.swal2-popup, .swal2-container')) {
                        if (el.offsetParent === null) continue;
                        const t = el.querySelector('.swal2-title');
                        const c = el.querySelector('.swal2-html-container, .swal2-content');
                        return ((t?t.textContent.trim():'') + ' | ' + (c?c.textContent.trim():'')).slice(0, 600);
                      }
                      for (const el of document.querySelectorAll('.modal.show, .modal[style*="display: block"]')) {
                        const t = el.querySelector('.modal-title');
                        const b = el.querySelector('.modal-body');
                        return ((t?t.textContent.trim():'') + ' | ' + (b?b.textContent.trim().slice(0,500):''));
                      }
                      return '';
                    }""")
                except:
                    pass
                
                modal_open = page.evaluate(r"""() => {
                    const b = document.getElementById('btn_search_alien_submit');
                    return !!(b && b.offsetParent !== null);
                }""")
                
                # Check if search_alien_modal is visible
                search_modal_vis = page.evaluate(r"""() => {
                    const m = document.getElementById('search_alien_modal');
                    return m ? (m.offsetParent !== null ? 'visible' : 'hidden_dom') : 'not_found';
                }""")
                
                # Check what modal/alert elements exist
                modals = page.evaluate(r"""() => {
                    const result = {};
                    document.querySelectorAll('.swal2-popup, .swal2-container').forEach(el => {
                        result['swal2'] = el.offsetParent !== null ? 'visible' : 'hidden';
                    });
                    document.querySelectorAll('.modal.show').forEach(el => {
                        result['modal.show'] = el.offsetParent !== null ? 'visible' : 'hidden';
                    });
                    document.querySelectorAll('.modal').forEach((el, i) => {
                        if (i < 3) result[`modal[${i}]`] = {
                            id: el.id,
                            visible: el.offsetParent !== null,
                            class: el.className,
                            display: el.style.display
                        };
                    });
                    return result;
                }""")
                
                print(f"[{t}ms] alert='{alert_text[:100]}...' | modal_open={modal_open} | search_modal={search_modal_vis} | modals={modals}")
                
                if t in [2000, 3000, 4000]:
                    page.screenshot(path=f"probe_at_{t}ms.png", full_page=True)
            
        finally:
            browser.close()

if __name__ == "__main__":
    main()
