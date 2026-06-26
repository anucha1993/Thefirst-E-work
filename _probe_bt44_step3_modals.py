from pathlib import Path
import json
from playwright.sync_api import sync_playwright
from scrape_wa import (
    load_config, _read_bt44_excel, _open_bt44_form, _bt44_fill_search_one,
    _bt44_step2_consent, _bt44_step2_verify_permit, _bt44_step2_edit_address,
    _wait_loading_disappeared, login,
)

cfg = load_config(require_login=True)
recs = _read_bt44_excel(Path('from_bt44.xlxs.xlsx'))
rec = recs[0]
ss = Path('reports/bt44_screenshots')
ss.mkdir(parents=True, exist_ok=True)

def dump_visible(page, label):
    data = page.evaluate(r"""(label) => {
        const isVisible = el => {
            if (!el) return false;
            const st = window.getComputedStyle(el);
            return el.offsetParent !== null && st.display !== 'none' && st.visibility !== 'hidden' && st.opacity !== '0';
        };
        const text = el => (el.textContent || '').replace(/\s+/g, ' ').trim();
        return {
            label,
            headings: Array.from(document.querySelectorAll('h1,h2,h3,h4,h5,label,legend,.modal-title,.swal2-title')).filter(isVisible).map(e => text(e)).filter(Boolean).slice(0,120),
            buttons: Array.from(document.querySelectorAll('button,a,[role="button"]')).filter(isVisible).map(e => ({text:text(e),id:e.id||'',cls:(e.className||'').toString().slice(0,120)})).filter(x => x.text).slice(0,120),
            selects: Array.from(document.querySelectorAll('select')).filter(isVisible).map(s => ({id:s.id||'',name:s.name||'',title:text(s.closest('div,section,fieldset,tr,.modal,.form-group')||s).slice(0,300),options:Array.from(s.options).slice(0,15).map(o => (o.textContent||'').trim())})),
            inputs: Array.from(document.querySelectorAll('input,textarea')).filter(isVisible).map(i => ({id:i.id||'',name:i.name||'',type:i.type||'',placeholder:i.placeholder||'',value:i.value||'',title:text(i.closest('div,section,fieldset,tr,.modal,.form-group')||i).slice(0,300)})).slice(0,200),
            bodySample: (document.body.innerText || '').slice(0,5000),
        };
    }""", label)
    return data

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
        next_clicked = False
        try:
            page.locator('#gonextSubmit').first.click(timeout=5000, force=True)
            next_clicked = True
        except Exception:
            pass
        if not next_clicked:
            page.evaluate(r"""() => {
                const btn = Array.from(document.querySelectorAll('button,a,[role="button"]')).find(b => /ถัดไป|NEXT|Next/i.test((b.textContent || '').trim()) && b.offsetParent !== null);
                if (btn) btn.click();
            }""")
        page.wait_for_timeout(1200)
        _wait_loading_disappeared(page, timeout_ms=12000, log=print)
        page.wait_for_timeout(800)

        results = []
        results.append(dump_visible(page, 'step3_base'))
        page.screenshot(path=str(ss / 'step3_base.png'), full_page=True)

        # change employer modal
        page.evaluate(r"""() => { document.querySelector('#changeEmployer')?.click(); }""")
        page.wait_for_timeout(1500)
        _wait_loading_disappeared(page, timeout_ms=8000, log=print)
        results.append(dump_visible(page, 'changeEmployer_open'))
        page.screenshot(path=str(ss / 'step3_change_employer.png'), full_page=True)

        # close modal best-effort
        page.evaluate(r"""() => {
            const btn = Array.from(document.querySelectorAll('button,a,[role="button"]')).find(b => /ยกเลิก|ปิด|close/i.test((b.textContent || '').trim()) && b.offsetParent !== null);
            if (btn) btn.click();
        }""")
        page.wait_for_timeout(1000)

        # edit workplace modal
        try:
            page.locator("button:has-text('แก้ไขสถานประกอบการ'), a:has-text('แก้ไขสถานประกอบการ')").first.click(timeout=5000, force=True)
        except Exception:
            page.evaluate(r"""() => {
                const btn = Array.from(document.querySelectorAll('button,a,[role="button"]')).find(b => /แก้ไขสถานประกอบการ/.test((b.textContent || '').trim()) && b.offsetParent !== null);
                if (btn) btn.click();
            }""")
        page.wait_for_timeout(1500)
        _wait_loading_disappeared(page, timeout_ms=8000, log=print)
        results.append(dump_visible(page, 'edit_workplace_open'))
        page.screenshot(path=str(ss / 'step3_edit_workplace.png'), full_page=True)

        Path('step3_modals_probe.json').write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding='utf-8')
        print(json.dumps(results, ensure_ascii=False, indent=2))
    finally:
        ctx.close(); browser.close()
