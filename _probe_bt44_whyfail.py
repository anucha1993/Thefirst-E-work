"""Probe: ทำไมค้นหานายจ้าง (นายจ้างต่างชาติ + 0215566011533) ผ่านโค้ดไม่เจอ
แต่ค้นมือเจอ — ตรวจว่า select 'ประเภทค้นหา' ติดจริงไหม + placeholder เปลี่ยนไหม
"""
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

search_type = (rec.get('change_emp_search_type') or '').strip()
keyword = (rec.get('change_emp_keyword') or '').strip()
print(f"[probe] search_type={search_type!r} keyword={keyword!r}")

with sync_playwright() as pw:
    browser = pw.chromium.launch(headless=False, args=['--ignore-certificate-errors', '--start-maximized'])
    ctx = browser.new_context(locale='th-TH', ignore_https_errors=True, viewport={'width': 1920, 'height': 1080})
    page = ctx.new_page()
    out = {}
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

        # เปิด modal เปลี่ยนนายจ้าง
        page.locator('#changeEmployer').click(timeout=5000, force=True)
        page.wait_for_timeout(1200)
        _wait_loading_disappeared(page, timeout_ms=8000, log=print)

        # ---- ก่อนเลือก: อ่านสถานะ select + ช่องกรอก ----
        out['before_select'] = page.evaluate(r"""() => {
            const s = document.getElementById('search-type');
            const k = document.getElementById('search-keyword');
            return {
                select_value: s ? s.value : null,
                select_selected_text: s && s.selectedIndex >= 0 ? (s.options[s.selectedIndex].textContent||'').trim() : null,
                has_select2: !!document.querySelector('#search-type + .select2, .select2-container'),
                select2_text: (document.querySelector('.select2-selection__rendered')||{}).textContent || null,
                keyword_placeholder: k ? k.placeholder : null,
                onchange_attr: s ? (s.getAttribute('onchange')||'') : null,
            };
        }""")

        # ---- เลือกประเภทค้นหาด้วย helper เดิม ----
        ok = _select_option_by_text(page, 'search-type', search_type)
        print('[probe] _select_option_by_text ->', ok)
        page.wait_for_timeout(800)

        # ---- หลังเลือก: อ่านสถานะอีกครั้ง (placeholder ควรเปลี่ยนถ้า onchange ทำงาน) ----
        out['after_select'] = page.evaluate(r"""() => {
            const s = document.getElementById('search-type');
            const k = document.getElementById('search-keyword');
            return {
                select_value: s ? s.value : null,
                select_selected_text: s && s.selectedIndex >= 0 ? (s.options[s.selectedIndex].textContent||'').trim() : null,
                select2_text: (document.querySelector('.select2-selection__rendered')||{}).textContent || null,
                keyword_placeholder: k ? k.placeholder : null,
            };
        }""")

        # ---- กรอก keyword + ค้นหา ----
        page.evaluate(r"""(v) => {
            const t = document.getElementById('search-keyword');
            if (t) { t.value = v; t.dispatchEvent(new Event('input', {bubbles:true})); t.dispatchEvent(new Event('change', {bubbles:true})); }
        }""", keyword)
        page.wait_for_timeout(300)
        page.locator('#search-action').click(timeout=5000, force=True)
        page.wait_for_timeout(2500)
        _wait_loading_disappeared(page, timeout_ms=8000, log=print)
        page.screenshot(path=str(ss / 'step3_whyfail.png'), full_page=True)

        body = page.evaluate(r"""() => (document.body.innerText || '')""")
        out['found'] = 'ข้อมูลนายจ้างใหม่' in body
        out['not_found_alert'] = 'ไม่พบข้อมูลนายจ้างในระบบ' in body
        # ดึงชื่อบริษัทใหม่ถ้าเจอ
        import re
        m = re.search(r'ข้อมูลนายจ้างใหม่[\s\S]{0,400}', body)
        out['new_employer_block'] = (m.group(0)[:400] if m else '')
        out['body_tail'] = body[-1500:]

        print(json.dumps(out, ensure_ascii=False, indent=2))
        Path('step3_whyfail.json').write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding='utf-8')
        page.wait_for_timeout(2000)
    finally:
        ctx.close()
        browser.close()
