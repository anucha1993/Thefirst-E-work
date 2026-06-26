"""Probe Step 4 (แนบเอกสาร): เดินครบ Step 1-3.4 ด้วยฟังก์ชันใน scrape_wa.py ที่ทำงานได้แล้ว
จากนั้น dump โครงสร้าง DOM ของหน้าแนบเอกสาร (file inputs + label + required + ปุ่ม)
เพื่อดูว่าแต่ละช่องเอกสารผูกกับ input[type=file] ตัวไหน, อันไหน required (*), ปุ่มเพิ่มเอกสาร/ถัดไป
"""
from pathlib import Path
import json
import time
from playwright.sync_api import sync_playwright
from scrape_wa import (
    load_config, _read_bt44_excel, _open_bt44_form, _bt44_fill_search_one,
    _bt44_step2_consent, _bt44_step2_verify_permit, _bt44_step2_edit_address,
    _bt44_step3_change_employer, _bt44_step3_workplace,
    _wait_loading_disappeared, login,
)

cfg = load_config(require_login=True)
rec = _read_bt44_excel(Path('from_bt44.xlxs.xlsx'))[0]
ss = Path('reports/bt44_screenshots')
ss.mkdir(parents=True, exist_ok=True)

_T0 = time.time()
def tlog(msg=""):
    print(f"[t+{time.time() - _T0:6.1f}s] {msg}")

# ---- dump โครงสร้างหน้าแนบเอกสาร ----
DUMP_UPLOAD_JS = r"""() => {
    const norm = s => (s || '').replace(/\s+/g, ' ').trim();
    const vis = el => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && (el.offsetParent !== null || r.width > 0 || r.height > 0);
    };
    // หา "การ์ด/แถว" ที่ห่อ input[type=file] เพื่อดึงหัวข้อเอกสาร + ตัวบ่งชี้ required
    const fileInputs = Array.from(document.querySelectorAll('input[type=file]')).map((inp, idx) => {
        // เดินขึ้นหา container ที่มีข้อความหัวข้อ
        let card = inp;
        let title = '';
        for (let up = 0; up < 8 && card; up++) {
            card = card.parentElement;
            if (!card) break;
            const t = norm(card.innerText || '');
            if (t && t.length > 3) { title = t; if (t.length > 12) break; }
        }
        // หา label ใกล้เคียง
        const lbl = inp.closest('label') || (inp.id ? document.querySelector(`label[for="${inp.id}"]`) : null);
        const required = /\*/.test(title) || inp.required || /required/i.test(inp.className || '');
        return {
            idx, id: inp.id || '', name: inp.name || '', accept: inp.accept || '',
            visible: vis(inp), dataAttrs: Array.from(inp.attributes).filter(a => a.name.startsWith('data-')).map(a => `${a.name}=${a.value}`),
            labelText: norm(lbl ? lbl.textContent : '').slice(0, 120),
            cardTitle: title.slice(0, 160),
            required,
        };
    });
    // หัวข้อเอกสารทั้งหมด (label/หัวข้อที่มี * ) เพื่อเทียบ required
    const docLabels = Array.from(document.querySelectorAll('label,h4,h5,h6,p,span,div'))
        .filter(el => vis(el))
        .map(el => norm(el.textContent))
        .filter(t => /\*/.test(t) && t.length < 200 && t.length > 8);
    // ปุ่มสำคัญ
    const buttons = Array.from(document.querySelectorAll('button,a,[role=button]'))
        .filter(el => vis(el))
        .map(el => ({ text: norm(el.textContent).slice(0, 50), id: el.id || '',
            da: el.getAttribute('data-action') || '', cls: (el.className || '').toString().slice(0, 100) }))
        .filter(b => b.text);
    return {
        url: location.href,
        heading: norm((document.querySelector('h1,h2,h3,.step-title,.card-title') || {}).textContent || '').slice(0, 200),
        fileInputs,
        docLabels: Array.from(new Set(docLabels)).slice(0, 60),
        buttons: buttons.slice(0, 80),
        bodySample: norm(document.body.innerText || '').slice(0, 4000),
    };
}"""

with sync_playwright() as pw:
    browser = pw.chromium.launch(headless=False, args=['--ignore-certificate-errors', '--start-maximized'])
    ctx = browser.new_context(locale='th-TH', ignore_https_errors=True, viewport={'width': 1920, 'height': 1080})
    page = ctx.new_page()
    out = {}
    try:
        tlog('login...')
        login(page, cfg)
        page.wait_for_timeout(1500)
        assert _open_bt44_form(page, log=print), 'open form failed'
        tlog('Step 1 search...')
        _bt44_fill_search_one(page, rec, ss, log=print)
        page.wait_for_timeout(800)
        page.evaluate(r"""() => {
            for (const m of document.querySelectorAll('.modal, .swal2-popup, [role="dialog"]')) {
                if (m.offsetParent !== null) {
                    const b = Array.from(m.querySelectorAll('button,a')).find(x => /ปิด|ยืนยัน|ตกลง|ok/i.test((x.textContent||'').trim()));
                    if (b) b.click();
                }
            }
        }""")
        page.wait_for_timeout(500)
        tlog('Step 2.1 consent...')
        assert _bt44_step2_consent(page, log=print), 'consent failed'
        tlog('Step 2.3 verify permit...')
        _bt44_step2_verify_permit(page, rec, log=print)
        tlog('Step 2.2 edit address...')
        _bt44_step2_edit_address(page, rec, ss, log=print)
        # กดถัดไปหลัง Step 2.2
        try:
            page.locator('#gonextSubmit').first.click(timeout=5000, force=True)
        except Exception:
            page.evaluate(r"""() => {
                const b = Array.from(document.querySelectorAll('button,a')).find(x => /ถัดไป/.test((x.textContent||'').trim()) && x.offsetParent !== null);
                if (b) b.click();
            }""")
        page.wait_for_timeout(1200)
        _wait_loading_disappeared(page, timeout_ms=10000, log=print)
        tlog('Step 3.1 change employer...')
        r31 = _bt44_step3_change_employer(page, rec, log=print)
        print('[probe] step3.1 ->', json.dumps(r31, ensure_ascii=False))
        assert r31.get('status') == 'SUCCESS', f"step3.1 not success: {r31}"
        tlog('Step 3.3-3.4 workplace...')
        r33 = _bt44_step3_workplace(page, rec, log=print)
        print('[probe] step3.3 ->', json.dumps(r33, ensure_ascii=False))
        assert r33.get('status') == 'SUCCESS', f"step3.3 not success: {r33}"

        # ---- ตอนนี้อยู่ Step 4 (แนบเอกสาร) ----
        tlog('Step 4 reached — dumping DOM...')
        page.wait_for_timeout(1500)
        _wait_loading_disappeared(page, timeout_ms=8000, log=print, allow_modal=True)
        page.screenshot(path=str(ss / 'step4_attach_docs.png'), full_page=True)
        out['step4'] = page.evaluate(DUMP_UPLOAD_JS)
        Path('step4_probe.json').write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding='utf-8')
        print('[probe] saved step4_probe.json')
        print('[probe] heading:', out['step4'].get('heading'))
        print('[probe] fileInputs count:', len(out['step4'].get('fileInputs', [])))

        # ---- ทดลอง upload รูปถ่าย (idx 6) เพื่อจับ crop modal ----
        def find_input_by_th(substr):
            return page.evaluate(r"""(substr) => {
                const inps = Array.from(document.querySelectorAll('input[type=file]'));
                for (const inp of inps) {
                    const raw = inp.getAttribute('data-document-th') || '';
                    let th = raw;
                    try { th = (JSON.parse(raw).TH) || raw; } catch(e) {}
                    if (th && th.includes(substr)) return inp.id || inp.name || '';
                }
                return '';
            }""", substr)

        photo_id = find_input_by_th('รูปถ่าย')
        print('[probe] photo input id =', photo_id)
        photo_path = str((Path('Files') / 'EI EI THWIN.jpg').resolve())
        if photo_id:
            try:
                page.set_input_files(f'#{photo_id}', photo_path)
                print('[probe] set photo file ok')
            except Exception as e:
                print('[probe] set photo via id failed:', str(e)[:120])
                # fallback by name
                page.set_input_files(f'input[name="{photo_id}"]', photo_path)
            page.wait_for_timeout(2500)
            _wait_loading_disappeared(page, timeout_ms=6000, log=print, allow_modal=True)
            page.screenshot(path=str(ss / 'step4_crop_modal.png'), full_page=True)
            out['crop_modal'] = page.evaluate(DUMP_UPLOAD_JS)
            # จับ element เฉพาะใน crop modal (ปุ่ม/heading/canvas/img)
            out['crop_detail'] = page.evaluate(r"""() => {
                const norm = s => (s||'').replace(/\s+/g,' ').trim();
                const vis = el => { if(!el) return false; const st=getComputedStyle(el); const r=el.getBoundingClientRect(); return st.display!=='none'&&st.visibility!=='hidden'&&r.width>20&&r.height>20; };
                const modals = Array.from(document.querySelectorAll('.modal, [role=dialog], .swal2-popup')).filter(vis);
                const btns = Array.from(document.querySelectorAll('button,a')).filter(vis)
                    .map(b => ({text:norm(b.textContent).slice(0,40), id:b.id||'', da:b.getAttribute('data-action')||'', cls:(b.className||'').toString().slice(0,80)}))
                    .filter(b => b.text);
                const imgs = Array.from(document.querySelectorAll('img,canvas')).filter(vis)
                    .map(i => ({tag:i.tagName, id:i.id||'', cls:(i.className||'').toString().slice(0,80)}));
                return { modalCount: modals.length, buttons: btns.slice(0,40), imgs: imgs.slice(0,15) };
            }""")
            print('[probe] crop_detail buttons:', json.dumps(out['crop_detail']['buttons'], ensure_ascii=False)[:1200])
            print('[probe] crop_detail imgs:', json.dumps(out['crop_detail']['imgs'], ensure_ascii=False)[:600])

            # ลองกดบันทึกใน crop modal
            try:
                cb = page.locator("button:visible:has-text('บันทึก'), a:visible:has-text('บันทึก')").last
                if cb.count() > 0:
                    cb.click(timeout=4000, force=True)
                    print('[probe] clicked crop บันทึก')
            except Exception as e:
                print('[probe] crop บันทึก click err:', str(e)[:120])
            page.wait_for_timeout(2000)
            _wait_loading_disappeared(page, timeout_ms=6000, log=print, allow_modal=True)
            page.screenshot(path=str(ss / 'step4_after_crop.png'), full_page=True)
            out['after_crop'] = page.evaluate(DUMP_UPLOAD_JS)

        # ---- ทดลองเปิด modal 'เพิ่มเอกสาร' (popup-upload) เพื่อจับโครงสร้าง ----
        try:
            ab = page.locator("[data-action='popup-upload']:visible, button:visible:has-text('เพิ่มเอกสาร')").last
            if ab.count() > 0:
                ab.scroll_into_view_if_needed(timeout=2000)
                ab.click(timeout=4000, force=True)
                print('[probe] clicked เพิ่มเอกสาร')
        except Exception as e:
            print('[probe] เพิ่มเอกสาร click err:', str(e)[:120])
        page.wait_for_timeout(1500)
        page.screenshot(path=str(ss / 'step4_adddoc_modal.png'), full_page=True)
        out['adddoc_modal'] = page.evaluate(r"""() => {
            const norm = s => (s||'').replace(/\s+/g,' ').trim();
            const vis = el => { if(!el) return false; const st=getComputedStyle(el); const r=el.getBoundingClientRect(); return st.display!=='none'&&st.visibility!=='hidden'&&(el.offsetParent!==null||r.width>0); };
            const fileInputs = Array.from(document.querySelectorAll('input[type=file]')).filter(vis)
                .map(i => ({id:i.id||'', name:i.name||'', accept:i.accept||''}));
            const textInputs = Array.from(document.querySelectorAll('input[type=text],input:not([type]),textarea')).filter(vis)
                .map(i => ({id:i.id||'', name:i.name||'', placeholder:i.placeholder||''}));
            const btns = Array.from(document.querySelectorAll('button,a')).filter(vis)
                .map(b => ({text:norm(b.textContent).slice(0,40), id:b.id||'', da:b.getAttribute('data-action')||'', cls:(b.className||'').toString().slice(0,80)}))
                .filter(b => b.text);
            return { fileInputs, textInputs, buttons: btns.slice(0,40) };
        }""")
        print('[probe] adddoc_modal fileInputs:', json.dumps(out['adddoc_modal']['fileInputs'], ensure_ascii=False))
        print('[probe] adddoc_modal textInputs:', json.dumps(out['adddoc_modal']['textInputs'], ensure_ascii=False))
        print('[probe] adddoc_modal buttons:', json.dumps(out['adddoc_modal']['buttons'], ensure_ascii=False)[:1000])

        Path('step4_probe.json').write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding='utf-8')
        print('[probe] saved step4_probe.json (full)')
        page.wait_for_timeout(4000)
    finally:
        ctx.close()
        browser.close()
