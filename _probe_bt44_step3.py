from pathlib import Path
import json
from playwright.sync_api import sync_playwright
from scrape_wa import (
    load_config,
    _read_bt44_excel,
    _open_bt44_form,
    _bt44_fill_search_one,
    _bt44_step2_consent,
    _bt44_step2_verify_permit,
    _bt44_step2_edit_address,
    _wait_loading_disappeared,
    login,
)

cfg = load_config(require_login=True)
recs = _read_bt44_excel(Path('from_bt44.xlxs.xlsx'))
rec = recs[0]
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
        r1 = _bt44_fill_search_one(page, rec, ss, log=print)
        print('STEP1', r1)
        page.wait_for_timeout(800)
        # close leftover modal
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
        ok, permit = _bt44_step2_verify_permit(page, rec, log=print)
        print('PERMIT', ok, permit)
        r2 = _bt44_step2_edit_address(page, rec, ss, log=print)
        print('STEP2.2', r2)
        next_clicked = False
        try:
            next_btn = page.locator('#gonextSubmit').first
            next_btn.click(timeout=5000, force=True)
            next_clicked = True
        except Exception:
            pass
        if not next_clicked:
            try:
                alt_next = page.locator("button:has-text('ถัดไป'), a:has-text('ถัดไป')").first
                if alt_next.count() > 0:
                    alt_next.click(timeout=5000, force=True)
                    next_clicked = True
            except Exception:
                pass
        if not next_clicked:
            page.evaluate(r"""() => {
                const btn = Array.from(document.querySelectorAll('button,a,[role="button"]'))
                    .find(b => /ถัดไป|NEXT|Next/i.test((b.textContent || '').trim()) && b.offsetParent !== null);
                if (btn) btn.click();
            }""")
        page.wait_for_timeout(1200)
        _wait_loading_disappeared(page, timeout_ms=12000, log=print)
        page.wait_for_timeout(1000)
        page.screenshot(path=str(ss / 'step3_probe.png'), full_page=True)
        data = page.evaluate(r"""() => {
            const isVisible = el => {
                if (!el) return false;
                const st = window.getComputedStyle(el);
                return el.offsetParent !== null && st.display !== 'none' && st.visibility !== 'hidden' && st.opacity !== '0';
            };
            const text = el => (el.textContent || '').replace(/\s+/g, ' ').trim();
            return {
                url: location.href,
                headings: Array.from(document.querySelectorAll('h1,h2,h3,h4,h5,legend,.card-title,.page-title,label')).filter(isVisible).map(e => text(e)).filter(Boolean).slice(0,120),
                buttons: Array.from(document.querySelectorAll('button,a,[role="button"]')).filter(isVisible).map(e => ({text:text(e),id:e.id||'',cls:(e.className||'').toString().slice(0,120)})).filter(x => x.text).slice(0,120),
                selects: Array.from(document.querySelectorAll('select')).filter(isVisible).map(s => ({id:s.id||'',name:s.name||'',title:text(s.closest('div,section,fieldset,tr')||s).slice(0,250),options:Array.from(s.options).slice(0,10).map(o => (o.textContent||'').trim())})),
                inputs: Array.from(document.querySelectorAll('input,textarea')).filter(isVisible).map(i => ({id:i.id||'',name:i.name||'',type:i.type||'',placeholder:i.placeholder||'',value:i.value||'',title:text(i.closest('div,section,fieldset,tr')||i).slice(0,250)})).slice(0,160),
                bodySample: (document.body.innerText || '').slice(0,4000),
            };
        }""")
        Path('step3_probe.json').write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
        print(json.dumps(data, ensure_ascii=False, indent=2))
    finally:
        ctx.close(); browser.close()
