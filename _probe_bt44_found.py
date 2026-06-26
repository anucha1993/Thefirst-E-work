"""Probe: ค้นหานายจ้าง (เจอ) ด้วย select2 helper → กดบันทึก → จับ DOM ของ Step 3.2/3.3
ใช้ _select2_pick เพื่อให้ select2 UI อัปเดตจริง แล้วเดินต่อไปยังเหตุผล/สถานที่/ประเภทงาน
"""
from pathlib import Path
import json
import re
import time
from playwright.sync_api import sync_playwright
from scrape_wa import (
    load_config, _read_bt44_excel, _open_bt44_form, _bt44_fill_search_one,
    _bt44_step2_consent, _bt44_step2_verify_permit, _bt44_step2_edit_address,
    _wait_loading_disappeared, login, _select2_pick,
)

cfg = load_config(require_login=True)
rec = _read_bt44_excel(Path('from_bt44.xlxs.xlsx'))[0]
ss = Path('reports/bt44_screenshots')
ss.mkdir(parents=True, exist_ok=True)

search_type = (rec.get('change_emp_search_type') or '').strip()
keyword = (rec.get('change_emp_keyword') or '').strip()
print(f"[probe] search_type={search_type!r} keyword={keyword!r}")

_T_START = time.time()
def tlog(msg=""):
    print(f"[t+{time.time() - _T_START:6.1f}s] {msg}")

DUMP_JS = r"""() => {
    const isVisible = el => {
        if (!el) return false;
        const st = window.getComputedStyle(el);
        return el.offsetParent !== null && st.display !== 'none' && st.visibility !== 'hidden' && st.opacity !== '0';
    };
    const text = el => (el.textContent || '').replace(/\s+/g, ' ').trim();
    return {
        headings: Array.from(document.querySelectorAll('h1,h2,h3,h4,h5,label,legend,.modal-title,.swal2-title')).filter(isVisible).map(e => text(e)).filter(Boolean).slice(0,150),
        buttons: Array.from(document.querySelectorAll('button,a,[role="button"]')).filter(isVisible).map(e => ({text:text(e),id:e.id||'',cls:(e.className||'').toString().slice(0,120)})).filter(x => x.text).slice(0,150),
        selects: Array.from(document.querySelectorAll('select')).filter(isVisible).map(s => ({id:s.id||'',name:s.name||'',title:text(s.closest('div,section,fieldset,tr,.modal,.form-group')||s).slice(0,200),options:Array.from(s.options).slice(0,30).map(o => (o.textContent||'').trim())})),
        inputs: Array.from(document.querySelectorAll('input,textarea')).filter(isVisible).map(i => ({id:i.id||'',name:i.name||'',type:i.type||'',placeholder:i.placeholder||'',value:i.value||'',title:text(i.closest('div,section,fieldset,tr,.modal,.form-group')||i).slice(0,200)})).slice(0,200),
        bodySample: (document.body.innerText || '').slice(0,7000),
    };
}"""

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
        _t0 = time.time()
        _bt44_step2_edit_address(page, rec, ss, log=tlog)
        print(f"[probe] Step 2.2 address modal took {time.time() - _t0:.1f}s")
        page.evaluate(r"""() => {
            const btn = Array.from(document.querySelectorAll('button,a,[role="button"]')).find(b => /ถัดไป|NEXT|Next/i.test((b.textContent || '').trim()) && b.offsetParent !== null);
            if (btn) btn.click();
        }""")
        page.wait_for_timeout(1200)
        _wait_loading_disappeared(page, timeout_ms=12000, log=print)
        page.wait_for_timeout(1000)

        # เปิด modal เปลี่ยนนายจ้าง
        page.locator('#changeEmployer').click(timeout=5000, force=True)
        page.wait_for_timeout(1000)
        _wait_loading_disappeared(page, timeout_ms=6000, log=print, allow_modal=True)

        # เลือกประเภทค้นหา ผ่าน select2 helper
        ok = _select2_pick(page, 'search-type', search_type, log=print)
        print('[probe] _select2_pick ->', ok)
        page.wait_for_timeout(500)

        # อ่านสถานะ select2 + placeholder หลังเลือก
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
        print('[probe] after_select:', json.dumps(out['after_select'], ensure_ascii=False))

        # กรอก keyword + ค้นหา — ใช้ Playwright interaction จริง (fill + click)
        kw = page.locator('#search-keyword').first
        kw.scroll_into_view_if_needed(timeout=4000)
        kw.click(timeout=4000)
        kw.fill('')
        kw.fill(keyword)
        page.wait_for_timeout(300)
        sa = page.locator('#search-action').first
        sa.scroll_into_view_if_needed(timeout=4000)
        sa.click(timeout=6000)
        # รอผลค้นหา: เจอ 'ข้อมูลนายจ้างใหม่' หรือ alert ไม่พบ
        try:
            page.wait_for_function(
                r"""() => {
                    const t = document.body.innerText || '';
                    return /ข้อมูลนายจ้างใหม่/.test(t) || /ไม่พบข้อมูลนายจ้างในระบบ/.test(t);
                }""",
                timeout=12000,
            )
        except Exception:
            pass
        page.wait_for_timeout(1500)
        _wait_loading_disappeared(page, timeout_ms=8000, log=print, allow_modal=True)
        page.screenshot(path=str(ss / 'step3_found_search.png'), full_page=True)

        body = page.evaluate(r"""() => (document.body.innerText || '')""")
        out['found'] = 'ข้อมูลนายจ้างใหม่' in body
        out['not_found_alert'] = 'ไม่พบข้อมูลนายจ้างในระบบ' in body
        m = re.search(r'ข้อมูลนายจ้างใหม่[\s\S]{0,500}', body)
        out['new_employer_block'] = (m.group(0)[:500] if m else '')
        print('[probe] found=', out['found'], 'not_found_alert=', out['not_found_alert'])

        out['modal_before_save'] = page.evaluate(DUMP_JS)

        # เลือกเหตุผลการเปลี่ยนนายจ้าง (reason-type) = ค่าจาก Excel แล้วจับ textarea 'อื่นๆ' ที่โผล่
        reason = (rec.get('change_emp_reason') or '').strip()
        reason_other = (rec.get('change_emp_reason_other') or '').strip()
        print(f"[probe] reason={reason!r} reason_other={reason_other!r}")
        if reason:
            ok_reason = _select2_pick(page, 'reason-type', reason, log=print)
            print('[probe] reason _select2_pick ->', ok_reason)
            page.wait_for_timeout(800)
        # จับ textarea/input ทั้งหมดในโมดัลหลังเลือกเหตุผล (หา selector ของช่อง 'อื่นๆ')
        out['after_reason'] = page.evaluate(r"""() => {
            const vis = el => {
                if (!el) return false;
                const st = getComputedStyle(el);
                return el.offsetParent !== null && st.display !== 'none' && st.visibility !== 'hidden';
            };
            const dump = [];
            for (const el of document.querySelectorAll('textarea, input[type="text"]')) {
                dump.push({
                    tag: el.tagName.toLowerCase(),
                    id: el.id || '',
                    name: el.name || '',
                    placeholder: el.placeholder || '',
                    visible: vis(el),
                    value: el.value || '',
                });
            }
            return dump;
        }""")
        print('[probe] after_reason inputs:', json.dumps(out['after_reason'], ensure_ascii=False))
        # ถ้ามีช่องเหตุผลอื่นๆ ที่มองเห็น ลองพิมพ์ค่าจริง
        if reason_other:
            filled_other = page.evaluate(r"""(v) => {
                const vis = el => {
                    if (!el) return false;
                    const st = getComputedStyle(el);
                    return el.offsetParent !== null && st.display !== 'none' && st.visibility !== 'hidden';
                };
                const ta = Array.from(document.querySelectorAll('textarea, input[type="text"]'))
                    .find(e => vis(e) && !/search-keyword/.test(e.id) && (e.tagName === 'TEXTAREA' || /reason|other|อื่น/i.test((e.id||'') + (e.name||'') + (e.placeholder||''))));
                if (!ta) return '';
                ta.value = v;
                ta.dispatchEvent(new Event('input', {bubbles:true}));
                ta.dispatchEvent(new Event('change', {bubbles:true}));
                return ta.id || ta.name || ta.tagName.toLowerCase();
            }""", reason_other)
            print('[probe] filled reason_other into:', filled_other)
        page.wait_for_timeout(400)

        # กดบันทึก (ในmodal)
        saved = page.evaluate(r"""() => {
            const btn = Array.from(document.querySelectorAll('.modal button, .modal a, [role="dialog"] button, [role="dialog"] a, button, a'))
              .find(b => /บันทึก/.test((b.textContent||'').trim()) && b.offsetParent !== null);
            if (btn) { btn.click(); return true; }
            return false;
        }""")
        print('[probe] save clicked=', saved)
        page.wait_for_timeout(2000)
        _wait_loading_disappeared(page, timeout_ms=8000, log=print, allow_modal=True)
        # ปิด swal สำเร็จถ้ามี
        page.evaluate(r"""() => {
            const btn = Array.from(document.querySelectorAll('.swal2-confirm, button, a'))
              .find(b => /ตกลง|ยืนยัน|ปิด|ok/i.test((b.textContent||'').trim()) && b.offsetParent !== null);
            if (btn) btn.click();
        }""")
        page.wait_for_timeout(1500)
        _wait_loading_disappeared(page, timeout_ms=8000, log=print, allow_modal=True)
        page.screenshot(path=str(ss / 'step3_after_save_found.png'), full_page=True)

        # จับ DOM หลังบันทึก — ควรเห็นเหตุผลเปลี่ยนนายจ้าง / สถานที่ทำงาน / ประเภทงาน (Step 3.2/3.3)
        out['after_save'] = page.evaluate(DUMP_JS)

        # ---- Step 3.2: เปิดโมดัล 'เลือกสถานที่ทำงาน' (วินิจฉัยว่าทำไมไม่เปิด) ----
        # 1) จับ HTML + onclick ของปุ่ม และหา modal/div ที่ซ่อนอยู่
        out['workbtn_diag'] = page.evaluate(r"""() => {
            const norm = s => (s||'').replace(/\s+/g,' ').trim();
            const vis = el => {
                if (!el) return false;
                const st = getComputedStyle(el);
                const r = el.getBoundingClientRect();
                return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 0 && r.height > 0;
            };
            const btns = Array.from(document.querySelectorAll('a,button'))
                .filter(b => /เลือกสถานที่ทำงาน/.test(norm(b.textContent)));
            const btnInfo = btns.map(b => ({
                tag: b.tagName.toLowerCase(),
                visible: vis(b),
                text: norm(b.textContent).slice(0,40),
                onclick: b.getAttribute('onclick') || '',
                href: b.getAttribute('href') || '',
                cls: (b.className||'').toString().slice(0,160),
                outer: b.outerHTML.slice(0, 500),
            }));
            const modals = Array.from(document.querySelectorAll('div,section,form'))
                .filter(d => /work|address|branch|สถานที่|สาขา|business|job|cate|place/i.test((d.id||'') + ' ' + (d.className||'')))
                .map(d => ({ id: d.id||'', cls: (d.className||'').toString().slice(0,120), visible: vis(d),
                    display: getComputedStyle(d).display, tag: d.tagName.toLowerCase() }))
                .slice(0, 50);
            return { btnInfo, modals };
        }""")
        print('[probe] workbtn_diag btns:', json.dumps(out['workbtn_diag']['btnInfo'], ensure_ascii=False)[:1500])
        print('[probe] workbtn_diag modals:', json.dumps(out['workbtn_diag']['modals'], ensure_ascii=False)[:2000])

        # 2) ลองเปิดหลายวิธี แล้วดูว่าวิธีไหนทำให้โมดัลโผล่
        def modal_open_now():
            return page.evaluate(r"""() => {
                const vis = el => { if(!el) return false; const st=getComputedStyle(el); const r=el.getBoundingClientRect(); return st.display!=='none'&&st.visibility!=='hidden'&&r.width>40&&r.height>40; };
                const t = document.body.innerText || '';
                const hasSelects = Array.from(document.querySelectorAll('select')).some(s => vis(s) && Array.from(s.options).some(o => /BT\d|สำนักงาน/.test(o.textContent||'')));
                return { hasSelects, hasHeading: /สถานที่ทำงาน\/สาขา|ประเภทกิจการ/.test(t) };
            }""")

        opened = False
        # วิธี A: คลิกจริง (ไม่ force)
        try:
            wbtn = page.locator("a:visible:has-text('เลือกสถานที่ทำงาน'), button:visible:has-text('เลือกสถานที่ทำงาน')").last
            wbtn.scroll_into_view_if_needed(timeout=2000)
            wbtn.click(timeout=4000)
            page.wait_for_timeout(1500)
            st = modal_open_now(); print('[probe] after real click:', json.dumps(st, ensure_ascii=False))
            opened = st['hasSelects']
        except Exception as e:
            print('[probe] real click error:', str(e)[:120])
        # วิธี B: JS .click() ตรงปุ่ม
        if not opened:
            page.evaluate(r"""() => {
                const norm = s => (s||'').replace(/\s+/g,' ').trim();
                const b = Array.from(document.querySelectorAll('a,button')).find(x => /เลือกสถานที่ทำงาน/.test(norm(x.textContent)) && x.offsetParent !== null);
                if (b) b.click();
            }""")
            page.wait_for_timeout(1500)
            st = modal_open_now(); print('[probe] after JS click:', json.dumps(st, ensure_ascii=False))
            opened = st['hasSelects']
        # วิธี C: ถ้ามี onclick ให้ eval โดยตรง
        if not opened:
            ran = page.evaluate(r"""() => {
                const norm = s => (s||'').replace(/\s+/g,' ').trim();
                const b = Array.from(document.querySelectorAll('a,button')).find(x => /เลือกสถานที่ทำงาน/.test(norm(x.textContent)));
                if (!b) return 'no btn';
                const oc = b.getAttribute('onclick');
                if (oc) { try { (new Function('event', oc)).call(b, new Event('click')); return 'ran onclick: '+oc.slice(0,120); } catch(e){ return 'onclick err: '+e.message; } }
                return 'no onclick attr';
            }""")
            print('[probe] onclick eval:', ran)
            page.wait_for_timeout(1500)
            st = modal_open_now(); print('[probe] after onclick eval:', json.dumps(st, ensure_ascii=False))
            opened = st['hasSelects']
        print('[probe] workplace modal opened =', opened)
        _wait_loading_disappeared(page, timeout_ms=8000, log=print, allow_modal=True)
        page.screenshot(path=str(ss / 'step3_workplace_modal.png'), full_page=True)
        out['workplace_modal'] = page.evaluate(DUMP_JS)
        # dump รายละเอียด select/input ที่มองเห็นในโมดัล workplace
        out['workplace_fields'] = page.evaluate(r"""() => {
            const vis = el => {
                if (!el) return false;
                const st = getComputedStyle(el);
                return el.offsetParent !== null && st.display !== 'none' && st.visibility !== 'hidden';
            };
            const selects = [];
            for (const s of document.querySelectorAll('select')) {
                selects.push({ id: s.id||'', name: s.name||'', visible: vis(s),
                    options: Array.from(s.options).map(o => o.textContent.trim()).slice(0, 30) });
            }
            const inputs = [];
            for (const i of document.querySelectorAll('input, textarea')) {
                inputs.push({ tag: i.tagName.toLowerCase(), id: i.id||'', name: i.name||'',
                    type: i.type||'', placeholder: i.placeholder||'', visible: vis(i), value: i.value||'' });
            }
            const radios = [];
            for (const r of document.querySelectorAll('input[type="radio"], input[type="checkbox"]')) {
                const lbl = r.closest('label') || (r.id ? document.querySelector(`label[for="${r.id}"]`) : null);
                radios.push({ id: r.id||'', name: r.name||'', visible: vis(r),
                    label: (lbl ? lbl.textContent.trim() : '').slice(0, 80), checked: r.checked });
            }
            return { selects, inputs, radios };
        }""")
        print('[probe] workplace_fields selects:', json.dumps(out['workplace_fields']['selects'], ensure_ascii=False)[:1500])

        # ---- เลือกสาขา: select ที่ visible และมี option ขึ้นต้น 'สำนักงาน' ----
        branch_picked = page.evaluate(r"""() => {
            const vis = el => {
                if (!el) return false;
                const st = getComputedStyle(el);
                const r = el.getBoundingClientRect();
                return st.display !== 'none' && st.visibility !== 'hidden' && (el.offsetParent !== null || r.width > 0);
            };
            // หา select ที่ option ใดขึ้นต้น 'สำนักงาน' (สาขาเป้าหมาย) — เลือกตัวที่ visible ก่อน
            const selects = Array.from(document.querySelectorAll('select')).filter(s =>
                Array.from(s.options).some(o => /^\s*สำนักงาน/.test(o.textContent || '')));
            const sel = selects.find(vis) || selects[0];
            if (!sel) return { ok:false, reason:'no branch select' };
            const opt = Array.from(sel.options).find(o => /^\s*สำนักงาน/.test(o.textContent || ''));
            if (!opt) return { ok:false, reason:'no สำนักงาน option' };
            sel.value = opt.value;
            sel.dispatchEvent(new Event('input', { bubbles: true }));
            sel.dispatchEvent(new Event('change', { bubbles: true }));
            if (window.jQuery) { try { jQuery(sel).val(opt.value).trigger('change'); jQuery(sel).trigger({type:'select2:select', params:{data:{id:opt.value, text:opt.textContent}}}); } catch(e){} }
            return { ok:true, value: opt.value, text: (opt.textContent||'').replace(/\s+/g,' ').trim().slice(0,80), selId: sel.id||'', selName: sel.name||'' };
        }""")
        print('[probe] branch_picked:', json.dumps(branch_picked, ensure_ascii=False))
        page.wait_for_timeout(1500)
        _wait_loading_disappeared(page, timeout_ms=8000, log=print, allow_modal=True)
        page.screenshot(path=str(ss / 'step3_branch_selected.png'), full_page=True)

        # ---- จับ field งานที่โผล่หลังเลือกสาขา (businessType, permitCateWork, jobPosition, jobDescription) ----
        out['after_branch'] = page.evaluate(r"""() => {
            const vis = el => {
                if (!el) return false;
                const st = getComputedStyle(el);
                const r = el.getBoundingClientRect();
                return st.display !== 'none' && st.visibility !== 'hidden' && (el.offsetParent !== null || r.width > 0);
            };
            const selects = [];
            for (const s of document.querySelectorAll('select')) {
                if (!vis(s)) continue;
                selects.push({ id: s.id||'', name: s.name||'',
                    options: Array.from(s.options).map(o => (o.textContent||'').replace(/\s+/g,' ').trim()).slice(0, 40) });
            }
            const inputs = [];
            for (const i of document.querySelectorAll('input, textarea')) {
                if (!vis(i)) continue;
                inputs.push({ tag: i.tagName.toLowerCase(), id: i.id||'', name: i.name||'',
                    type: i.type||'', placeholder: i.placeholder||'', value: i.value||'' });
            }
            // ปุ่มที่มองเห็นในโมดัล
            const buttons = [];
            for (const b of document.querySelectorAll('button, a')) {
                if (!vis(b)) continue;
                const t = (b.textContent||'').replace(/\s+/g,' ').trim();
                if (t) buttons.push({ text: t.slice(0,40), id: b.id||'', cls: (b.className||'').slice(0,60) });
            }
            return { selects, inputs, buttons };
        }""")
        print('[probe] after_branch selects:', json.dumps(out['after_branch']['selects'], ensure_ascii=False)[:2000])
        print('[probe] after_branch inputs:', json.dumps(out['after_branch']['inputs'], ensure_ascii=False)[:1500])

        # ---- Step 3.3: เลือกประเภทกิจการ + ประเภทงาน + กรอกลักษณะงาน ----
        work_biz = (rec.get('work_biz') or '').strip()
        work_job = (rec.get('work_permit_job') or '').strip()
        work_detail = (rec.get('work_detail') or '').strip()
        print(f"[probe] work_biz={work_biz!r} work_job={work_job!r} work_detail={work_detail!r}")

        def pick_visible_select(opt_substr, want):
            return page.evaluate(r"""(args) => {
                const [optSubstr, want] = args;
                const norm = s => (s||'').replace(/\s+/g,' ').trim();
                const vis = el => {
                    if (!el) return false;
                    const st = getComputedStyle(el);
                    const r = el.getBoundingClientRect();
                    return st.display !== 'none' && st.visibility !== 'hidden' && (el.offsetParent !== null || r.width > 0);
                };
                const selects = Array.from(document.querySelectorAll('select')).filter(s =>
                    vis(s) && Array.from(s.options).some(o => norm(o.textContent).includes(optSubstr)));
                const sel = selects[0];
                if (!sel) return { ok:false, reason:'no select with '+optSubstr };
                const w = norm(want);
                let opt = Array.from(sel.options).find(o => norm(o.textContent) === w)
                       || Array.from(sel.options).find(o => norm(o.textContent).includes(w) && w)
                       || Array.from(sel.options).find(o => norm(o.textContent) && norm(o.textContent) !== '-- กรุณาเลือก --' && norm(o.textContent) !== 'กรุณาเลือก');
                if (!opt) return { ok:false, reason:'no option', selId: sel.id||'' };
                sel.value = opt.value;
                sel.dispatchEvent(new Event('input', { bubbles: true }));
                sel.dispatchEvent(new Event('change', { bubbles: true }));
                if (window.jQuery) { try { jQuery(sel).val(opt.value).trigger('change'); jQuery(sel).trigger({type:'select2:select', params:{data:{id:opt.value, text:opt.textContent}}}); } catch(e){} }
                return { ok:true, selId: sel.id||'', text: norm(opt.textContent).slice(0,60) };
            }""", [opt_substr, want])

        biz_res = pick_visible_select('BT', work_biz)
        print('[probe] businessType pick:', json.dumps(biz_res, ensure_ascii=False))
        page.wait_for_timeout(800)
        _wait_loading_disappeared(page, timeout_ms=6000, log=print, allow_modal=True)

        job_res = pick_visible_select('กรรมกร', work_job)
        print('[probe] permitCateWork pick:', json.dumps(job_res, ensure_ascii=False))
        page.wait_for_timeout(800)
        _wait_loading_disappeared(page, timeout_ms=6000, log=print, allow_modal=True)

        # กรอกลักษณะงาน (jobDescription) — มี element id ซ้ำหลายตัว ต้องเล็งเฉพาะตัวที่ "มองเห็น"
        # (.first จะไปโดน template ที่ซ่อนอยู่ ทำให้ wait_for(visible) timeout)
        jd_filled = False
        if work_detail:
            for attempt in range(3):
                try:
                    jd = page.locator('#jobDescription:visible, input[name="jobDescription"]:visible').last
                    jd.wait_for(state='visible', timeout=5000)
                    jd.scroll_into_view_if_needed(timeout=2000)
                    jd.click(timeout=3000)
                    jd.fill('')
                    jd.fill(work_detail)
                    jd.dispatch_event('input')
                    jd.dispatch_event('change')
                    val = jd.input_value(timeout=2000)
                    if (val or '').strip() == work_detail:
                        jd_filled = True
                        print(f'[probe] filled jobDescription (attempt {attempt+1}): {val!r}')
                        break
                    print(f'[probe] jobDescription value mismatch (attempt {attempt+1}): {val!r}')
                except Exception as e:
                    print(f'[probe] jobDescription fill error (attempt {attempt+1}):', str(e)[:120])
                page.wait_for_timeout(700)
        # JS fallback: เซ็ตค่าบน element ที่มองเห็นโดยตรง + dispatch events ให้ validation ผ่าน
        if not jd_filled and work_detail:
            r = page.evaluate(r"""(val) => {
                const vis = el => { if(!el) return false; const st=getComputedStyle(el); const rc=el.getBoundingClientRect(); return st.display!=='none'&&st.visibility!=='hidden'&&rc.width>0&&rc.height>0; };
                const els = Array.from(document.querySelectorAll('#jobDescription, input[name="jobDescription"]')).filter(vis);
                const el = els[els.length-1];
                if (!el) return {ok:false, reason:'no visible jobDescription'};
                el.focus();
                el.value = val;
                el.dispatchEvent(new Event('input', {bubbles:true}));
                el.dispatchEvent(new Event('change', {bubbles:true}));
                el.dispatchEvent(new Event('blur', {bubbles:true}));
                return {ok:true, value: el.value};
            }""", work_detail)
            print('[probe] jobDescription JS fallback:', json.dumps(r, ensure_ascii=False))
            jd_filled = bool(r.get('ok'))
        print('[probe] jd_filled=', jd_filled)
        page.wait_for_timeout(500)
        page.screenshot(path=str(ss / 'step3_job_filled.png'), full_page=True)
        out['after_job'] = page.evaluate(DUMP_JS)

        # บันทึกในโมดัล workplace
        msave = False
        try:
            sb = page.locator("button:visible:has-text('บันทึก'), a:visible:has-text('บันทึก')").last
            if sb.count() > 0:
                sb.scroll_into_view_if_needed(timeout=2000)
                sb.click(timeout=4000, force=True)
                msave = True
        except Exception as e:
            print('[probe] modal save error:', str(e)[:120])
        print('[probe] workplace modal save clicked=', msave)
        page.wait_for_timeout(1500)
        _wait_loading_disappeared(page, timeout_ms=8000, log=print, allow_modal=True)
        # ปิด swal สำเร็จถ้ามี
        page.evaluate(r"""() => {
            const btn = Array.from(document.querySelectorAll('.swal2-confirm, button, a'))
              .find(b => /ตกลง|ยืนยัน|ปิด|ok/i.test((b.textContent||'').trim()) && b.offsetParent !== null);
            if (btn) btn.click();
        }""")
        page.wait_for_timeout(1500)
        _wait_loading_disappeared(page, timeout_ms=6000, log=print, allow_modal=True)
        page.screenshot(path=str(ss / 'step3_workplace_saved.png'), full_page=True)
        out['after_workplace_save'] = page.evaluate(DUMP_JS)

        Path('step3_found_probe.json').write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding='utf-8')
        print('[probe] saved step3_found_probe.json')
        page.wait_for_timeout(3000)
    finally:
        ctx.close()
        browser.close()
