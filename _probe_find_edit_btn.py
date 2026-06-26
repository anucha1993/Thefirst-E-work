#!/usr/bin/env python3
"""
Probe: หา Edit button ที่ Step 2.2 โดยใช้วิธีทั้งหมด
บันทึก ALL clickable elements ลงไฟล์
"""

from pathlib import Path
from playwright.sync_api import sync_playwright
import json
import time
import sys

sys.path.insert(0, str(Path(__file__).parent))
from scrape_wa import (login, _bt44_fill_search_one, _open_bt44_form, 
                       _read_bt44_excel, _wait_loading_disappeared)

LOGIN = "Vararat@mailforeign.com"
PASSWORD = "Q1w2e3r4@!"

def main():
    excel_recs = _read_bt44_excel("from_bt44.xlxs.xlsx")
    if not excel_recs:
        print("No Excel data")
        return
    
    rec = excel_recs[0]
    screenshot_dir = Path("reports/bt44_screenshots")
    screenshot_dir.mkdir(parents=True, exist_ok=True)
    
    with sync_playwright() as pw:
        browser = pw.chromium.launch(args=["--ignore-certificate-errors", "--start-maximized"])
        ctx = browser.new_context(
            locale="th-TH",
            ignore_https_errors=True,
            viewport={"width": 1920, "height": 1080}
        )
        page = ctx.new_page()
        
        try:
            # === LOGIN ===
            print("[*] Navigating to login page...")
            page.goto("https://eworkpermit.doe.go.th/", wait_until="load", timeout=30000)
            page.wait_for_load_state("networkidle", timeout=15000)
            page.wait_for_timeout(2000)
            
            print("[*] Logging in...")
            login(page, {"username": LOGIN, "password": PASSWORD, "user_type": "ผู้กระทำการแทน", "method": "e-workpermit"})
            page.wait_for_timeout(1500)
            
            # === OPEN BT44 ===
            print("[*] Opening BT44 form...")
            _open_bt44_form(page, log=print)
            page.wait_for_timeout(1000)
            
            # === FILL STEP 1 ===
            print("[*] Filling Step 1...")
            res1 = _bt44_fill_search_one(page, rec, screenshot_dir, log=print)
            print(f"    Step 1: {res1.get('status')}")
            page.wait_for_timeout(2000)
            
            # === CLOSE MODAL ===
            page.evaluate(r"""() => {
                const modals = document.querySelectorAll('.modal, .swal2-popup, [role="dialog"]');
                for (const m of modals) {
                    if (m.offsetParent !== null) {
                        const closeBtn = m.querySelector('.close, .btn-close, .swal2-close');
                        if (closeBtn) closeBtn.click();
                    }
                }
            }""")
            page.wait_for_timeout(500)
            
            # === STEP 2.1 ===
            print("[*] Step 2.1: Checking consent...")
            page.evaluate(r"""() => {
                const chk = document.getElementById('check_truth');
                if (chk) {
                    chk.checked = true;
                    chk.dispatchEvent(new Event('change', { bubbles: true }));
                }
            }""")
            page.wait_for_timeout(500)
            
            # Click ถัดไป
            page.evaluate(r"""() => {
                const btn = Array.from(document.querySelectorAll('button, a'))
                    .find(b => /ถัดไป|NEXT/i.test((b.textContent || '').trim()));
                if (btn) btn.click();
            }""")
            
            page.wait_for_timeout(1500)
            _wait_loading_disappeared(page, timeout_ms=10000, log=print)
            page.wait_for_timeout(1000)
            
            # === NOW PROBE ALL CLICKABLES ===
            print("\n[*] Extracting ALL clickable elements...")
            
            all_elements = page.evaluate(r"""() => {
                const result = [];
                
                // Get ALL potential clickable elements
                const selectors = [
                    'button', 'a', '[role="button"]', 'input[type="button"]',
                    '[onclick]', '[ng-click]', '[v-on:click]', '[data-click]',
                    '.btn', '.btn-link', '.link', '[class*="button"]', '[class*="edit"]'
                ];
                
                const seen = new Set();
                
                for (const selector of selectors) {
                    try {
                        const elements = document.querySelectorAll(selector);
                        for (const elem of elements) {
                            const key = elem.outerHTML.substring(0, 100);
                            if (seen.has(key)) continue;
                            seen.add(key);
                            
                            const visible = elem.offsetParent !== null;
                            const style = window.getComputedStyle(elem);
                            const text = (elem.textContent || elem.value || '').trim().substring(0, 100);
                            const html = elem.outerHTML.substring(0, 200);
                            
                            if (text || elem.tagName === 'BUTTON' || elem.tagName === 'A') {
                                result.push({
                                    tag: elem.tagName,
                                    text: text,
                                    id: elem.id || '',
                                    class: elem.className || '',
                                    visible: visible,
                                    display: style.display,
                                    opacity: style.opacity,
                                    pointer: style.cursor,
                                    onclick: !!elem.onclick,
                                    ng_click: elem.getAttribute('ng-click') || '',
                                    data_action: elem.getAttribute('data-action') || '',
                                    html: html
                                });
                            }
                        }
                    } catch (e) {}
                }
                
                return result;
            }""")
            
            print(f"\nFound {len(all_elements)} clickable elements:")
            print("\n" + "="*100)
            
            edit_candidates = []
            
            for i, elem in enumerate(all_elements, 1):
                if elem["text"] or elem["tag"] in ["BUTTON", "A"]:
                    visible_marker = "✓" if elem["visible"] else "✗"
                    print(f"\n[{i}] {visible_marker} {elem['tag']}")
                    print(f"    Text: {elem['text']}")
                    if elem['id']:
                        print(f"    ID: {elem['id']}")
                    if elem['class']:
                        print(f"    Class: {elem['class'][:80]}")
                    if elem['onclick']:
                        print(f"    Has onclick: YES")
                    if elem['ng_click']:
                        print(f"    ng-click: {elem['ng_click'][:80]}")
                    print(f"    Display: {elem['display']}, Opacity: {elem['opacity']}, Cursor: {elem['pointer']}")
                    
                    # Check if this is edit button
                    if any(kw in elem['text'].lower() for kw in ['แก้ไข', 'edit', 'modify', 'pencil', 'ปากกา']):
                        print(f"    >>> EDIT BUTTON CANDIDATE! <<<")
                        edit_candidates.append(i)
            
            print("\n" + "="*100)
            print(f"\nEDIT BUTTON CANDIDATES: {edit_candidates}")
            print("\n✓ Full output saved above")
            
            # Save to file
            with open("debug_clickables.json", "w", encoding="utf-8") as f:
                json.dump(all_elements, f, ensure_ascii=False, indent=2)
            print("✓ Detailed JSON saved to: debug_clickables.json")
            
            # Take screenshots
            page.screenshot(path="step2_debug_full.png", full_page=True)
            print("✓ Screenshot: step2_debug_full.png")
            
        except Exception as e:
            print(f"Error: {e}")
            import traceback
            traceback.print_exc()
        finally:
            ctx.close()
            browser.close()

if __name__ == "__main__":
    main()
