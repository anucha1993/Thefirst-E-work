from pathlib import Path
import json
from playwright.sync_api import sync_playwright
from scrape_wa import (
    load_config, _read_bt44_excel, _open_bt44_form, _bt44_fill_search_one,
    _bt44_step2_consent, _bt44_step2_verify_permit, _bt44_step2_edit_address,
    _wait_loading_disappeared, login, _select_option_by_text,
)

cfg = load_config(require_login=True)
rec = _read_bt44_excel(Path('from_bt44.xlxs.xlsx'))[0]
ss = Path('reports/bt44_screenshots')
ss.mkdir(parents=True, exist_ok=True)

with sync_playwright() as pw:
    browser = pw.chromium.launch(headless=False, args=['--ignore-certificate-errors', '--start-maximized'])
    ctx = browser.new_context(locale='th-TH', ignore_https_errors=True, viewport={'width': 1920, 'height': 1080})
    page = ctx.new_page()
    try:
        login(page, cfg)
        page.wait_for_timeout(1500)
        assert _open_bt44_form(page, log=print)
        _bt44_fill_search_one(page, rec, ss, log=print)
        page.wait_for_timeout(800)
        page.evaluate(r"""() => {
            const modals = document.querySelectorAll('.modal, .swal2-popup, [role="dialog"]');
            for (const m of modals) {
                if (m.offsetParent !== null) {
                    const btn = Array.from(m.querySelectorAll('button,a')).find(b => /ปิด|ยืนยัน|ตกลง|ok/i.test((b.textContent||'').trim()));
                    if (btn) btn.click();
                }
            }
        }""")
        page.wait_for_timeout(500)
        assert _bt44_step2_consent(page, log=print)
        _bt44_step2_verify_permit(page, rec, log=print)
        _bt44_step2_edit_address(page, rec, ss, log=print)
        page.evaluate(r"""() => {
            const btn = Array.from(document.querySelectorAll('button,a,[role="button"]')).find(b => /ถัดไป|NEXT|Next/i.test((b.textContent || '').trim()) && b.offsetParent !== null);
            if (btn) btn.click();
        }""")
        page.wait_for_timeout(1200)
        _wait_loading_disappeared(page, timeout_ms=12000, log=print)
        page.wait_for_timeout(1000)

        # open employer modal
        page.locator('#changeEmployer').click(timeout=5000, force=True)
        page.wait_for_timeout(1200)
        _wait_loading_disappeared(page, timeout_ms=8000, log=print)

        # fill search type
        ok = _select_option_by_text(page, 'search-type', rec.get('change_emp_search_type', '') or rec.get('ประเภทการค้นหา-เปลี่ยนนายจ้าง', ''))
        print('search-type selected', ok)
        page.evaluate(r"""(v) => {
            const t = document.getElementById('search-keyword');
            if (t) {
                t.value = v;
                t.dispatchEvent(new Event('input', {bubbles:true}));
                t.dispatchEvent(new Event('change', {bubbles:true}));
            }
        }""", rec.get('change_emp_keyword', '') or rec.get('ระบุ-เปลี่ยนนายจ้าง', ''))
        page.locator('#search-action').click(timeout=5000, force=True)
        try:
            page.wait_for_function(
                r"""() => {
                    const txt = document.body.innerText || '';
                    return /ไม่สามารถดำเนินการต่อได้ เนื่องจากไม่พบข้อมูลนายจ้างในระบบ/.test(txt)
                        || /ข้อมูลนายจ้างเดิม/.test(txt)
                        || /คืนค่าเดิม/.test(txt)
                        || /นายจ้างหลัก/.test(txt);
                }""",
                timeout=8000,
            )
        except Exception:
            pass
        _wait_loading_disappeared(page, timeout_ms=8000, log=print)
        page.screenshot(path=str(ss / 'step3_after_search_employer.png'), full_page=True)

        # capture current visible text/buttons before save
        data1 = page.evaluate(r"""() => {
            const isVisible = el => {
                if (!el) return false;
                const st = window.getComputedStyle(el);
                return el.offsetParent !== null && st.display !== 'none' && st.visibility !== 'hidden' && st.opacity !== '0';
            };
            const text = el => (el.textContent || '').replace(/\s+/g, ' ').trim();
            return {
                headings: Array.from(document.querySelectorAll('h1,h2,h3,h4,h5,label,legend,.modal-title,.swal2-title')).filter(isVisible).map(e => text(e)).filter(Boolean).slice(0,120),
                buttons: Array.from(document.querySelectorAll('button,a,[role="button"]')).filter(isVisible).map(e => ({text:text(e),id:e.id||'',cls:(e.className||'').toString().slice(0,120)})).filter(x => x.text).slice(0,120),
                selects: Array.from(document.querySelectorAll('select')).filter(isVisible).map(s => ({id:s.id||'',name:s.name||'',title:text(s.closest('div,section,fieldset,tr,.modal,.form-group')||s).slice(0,300),options:Array.from(s.options).slice(0,20).map(o => (o.textContent||'').trim())})),
                inputs: Array.from(document.querySelectorAll('input,textarea')).filter(isVisible).map(i => ({id:i.id||'',name:i.name||'',type:i.type||'',placeholder:i.placeholder||'',value:i.value||'',title:text(i.closest('div,section,fieldset,tr,.modal,.form-group')||i).slice(0,300)})).slice(0,200),
                bodySample: (document.body.innerText || '').slice(0,5000),
            };
        }""")

        # try save employer modal
        page.evaluate(r"""() => {
            const btn = Array.from(document.querySelectorAll('button,a,[role="button"]')).find(b => /บันทึก/.test((b.textContent||'').trim()) && b.offsetParent !== null);
            if (btn) btn.click();
        }""")
        page.wait_for_timeout(2000)
        _wait_loading_disappeared(page, timeout_ms=8000, log=print)
        page.screenshot(path=str(ss / 'step3_after_save_employer.png'), full_page=True)
        data2 = page.evaluate(r"""() => {
            const isVisible = el => {
                if (!el) return false;
                const st = window.getComputedStyle(el);
                return el.offsetParent !== null && st.display !== 'none' && st.visibility !== 'hidden' && st.opacity !== '0';
            };
            const text = el => (el.textContent || '').replace(/\s+/g, ' ').trim();
            return {
                headings: Array.from(document.querySelectorAll('h1,h2,h3,h4,h5,label,legend,.modal-title,.swal2-title')).filter(isVisible).map(e => text(e)).filter(Boolean).slice(0,120),
                buttons: Array.from(document.querySelectorAll('button,a,[role="button"]')).filter(isVisible).map(e => ({text:text(e),id:e.id||'',cls:(e.className||'').toString().slice(0,120)})).filter(x => x.text).slice(0,120),
                selects: Array.from(document.querySelectorAll('select')).filter(isVisible).map(s => ({id:s.id||'',name:s.name||'',title:text(s.closest('div,section,fieldset,tr,.modal,.form-group')||s).slice(0,300),options:Array.from(s.options).slice(0,20).map(o => (o.textContent||'').trim())})),
                inputs: Array.from(document.querySelectorAll('input,textarea')).filter(isVisible).map(i => ({id:i.id||'',name:i.name||'',type:i.type||'',placeholder:i.placeholder||'',value:i.value||'',title:text(i.closest('div,section,fieldset,tr,.modal,.form-group')||i).slice(0,300)})).slice(0,200),
                bodySample: (document.body.innerText || '').slice(0,7000),
            };
        }""")

        Path('step3_after_employer_probe.json').write_text(json.dumps({'before_save': data1, 'after_save': data2}, ensure_ascii=False, indent=2), encoding='utf-8')
        print(json.dumps({'before_save': data1, 'after_save': data2}, ensure_ascii=False, indent=2))
    finally:
        ctx.close(); browser.close()
