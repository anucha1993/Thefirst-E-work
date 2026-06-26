"""Probe Step 5 (สรุปคำขอ) + Step 6 (ยืนยันตัวตน)
ใช้ฟังก์ชันใน scrape_wa.py ที่ทำเสร็จแล้วเดินถึง Step 4 → ดัมพ์ DOM ของ Step 5 + 6
"""
from pathlib import Path
import json
import time
from playwright.sync_api import sync_playwright
from scrape_wa import (
    load_config, _read_bt44_excel, _open_bt44_form, _bt44_fill_search_one,
    _bt44_step2_consent, _bt44_step2_verify_permit, _bt44_step2_edit_address,
    _bt44_step3_change_employer, _bt44_step3_workplace, _bt44_step4_attach_docs,
    _wait_loading_disappeared, login,
)

cfg = load_config(require_login=True)
rec = _read_bt44_excel(Path('from_bt44.xlxs.xlsx'))[0]
ss = Path('reports/bt44_screenshots')
ss.mkdir(parents=True, exist_ok=True)

_T0 = time.time()
def tlog(msg=""):
    print(f"[t+{time.time() - _T0:6.1f}s] {msg}")

DUMP_JS = r"""() => {
    const norm = s => (s || '').replace(/\s+/g, ' ').trim();
    const vis = el => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && (el.offsetParent !== null || r.width > 0);
    };
    const buttons = Array.from(document.querySelectorAll('button,a,[role=button],input[type=button],input[type=submit]'))
        .filter(vis)
        .map(el => ({
            text: norm(el.textContent || el.value || '').slice(0, 60),
            id: el.id || '',
            da: el.getAttribute('data-action') || '',
            type: el.tagName,
            cls: (el.className || '').toString().slice(0, 100),
            onclick: (el.getAttribute('onclick') || '').slice(0, 80),
        }))
        .filter(b => b.text);
    const checkboxes = Array.from(document.querySelectorAll('input[type=checkbox]'))
        .map(c => {
            const lbl = c.closest('label') || (c.id ? document.querySelector(`label[for="${c.id}"]`) : null);
            const lblTxt = norm(lbl ? lbl.textContent : '').slice(0, 200);
            // หาข้อความใกล้เคียง
            let nearTxt = '';
            let p = c.parentElement;
            for (let up = 0; up < 5 && p; up++) {
                const t = norm(p.innerText || '');
                if (t.length > 20) { nearTxt = t.slice(0, 200); break; }
                p = p.parentElement;
            }
            return { id: c.id || '', name: c.name || '', checked: c.checked,
                     visible: vis(c), label: lblTxt, near: nearTxt };
        });
    const fileInputs = Array.from(document.querySelectorAll('input[type=file]'))
        .map(i => ({ id: i.id || '', name: i.name || '', accept: i.accept || '',
                     visible: vis(i),
                     dataAttrs: Array.from(i.attributes).filter(a => a.name.startsWith('data-')).map(a => `${a.name}=${a.value.slice(0,80)}`) }));
    const imgs = Array.from(document.querySelectorAll('img'))
        .filter(vis)
        .filter(i => (i.src || '').length > 10)
        .map(i => ({ src: (i.src || '').slice(-60), id: i.id || '', cls: (i.className || '').toString().slice(0, 60) }));
    return {
        url: location.href,
        heading: norm((document.querySelector('h1,h2,h3,.step-title,.card-title') || {}).textContent || '').slice(0, 200),
        buttons: buttons.slice(0, 80),
        checkboxes,
        fileInputs,
        imgs: imgs.slice(0, 20),
        bodySample: norm(document.body.innerText || '').slice(0, 6000),
    };
}"""


def click_next(page, label='ถัดไป'):
    """กดปุ่มถัดไป (#button_next ถ้ามี ไม่งั้นปุ่ม visible ที่ข้อความ ตรง)"""
    try:
        nb = page.locator("#button_next:visible").first
        if nb.count() > 0:
            nb.scroll_into_view_if_needed(timeout=2000)
            nb.click(timeout=4000)
            return True
    except Exception:
        pass
    try:
        nb = page.locator(f"button:visible:has-text('{label}'), a:visible:has-text('{label}')").last
        if nb.count() > 0:
            nb.scroll_into_view_if_needed(timeout=2000)
            nb.click(timeout=4000)
            return True
    except Exception:
        pass
    return False


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
        s1 = _bt44_fill_search_one(page, rec, ss, log=print)
        assert s1.get('status') == 'SUCCESS', f"Step1 fail: {s1}"
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
        assert _bt44_step2_consent(page, log=print)
        tlog('Step 2.3 verify permit...')
        ok, cur = _bt44_step2_verify_permit(page, rec, log=print)
        if not ok:
            print(f'[probe] permit mismatch: cur={cur} expected={rec.get("workpermit_no")}')
            raise SystemExit('permit mismatch — ลองคนใหม่หรือทำคำขอใหม่')
        tlog('Step 2.2 edit address...')
        _bt44_step2_edit_address(page, rec, ss, log=print)
        try:
            page.locator('#gonextSubmit').first.click(timeout=5000, force=True)
        except Exception:
            page.evaluate(r"""() => { const b = Array.from(document.querySelectorAll('button,a')).find(x => /ถัดไป/.test((x.textContent||'').trim()) && x.offsetParent !== null); if (b) b.click(); }""")
        page.wait_for_timeout(1200)
        _wait_loading_disappeared(page, timeout_ms=10000, log=print)
        tlog('Step 3.1 change employer...')
        assert _bt44_step3_change_employer(page, rec, log=print).get('status') == 'SUCCESS'
        tlog('Step 3.3-3.4 workplace...')
        assert _bt44_step3_workplace(page, rec, log=print).get('status') == 'SUCCESS'
        tlog('Step 4 attach docs...')
        s4 = _bt44_step4_attach_docs(page, rec, ss, log=print)
        assert s4.get('status') == 'SUCCESS', f"Step4 fail: {s4}"
        tlog('Step 4 done — at Step 5 (สรุปคำขอ) ?')
        page.wait_for_timeout(1500)
        _wait_loading_disappeared(page, timeout_ms=8000, log=print, allow_modal=True)
        page.screenshot(path=str(ss / 'step5_initial.png'), full_page=True)
        out['step5_initial'] = page.evaluate(DUMP_JS)
        tlog(f"step5_initial heading: {out['step5_initial'].get('heading','')[:80]}")

        # ---- 5.1 click ถัดไป ----
        tlog('Step 5.1 click ถัดไป...')
        click_next(page)
        page.wait_for_timeout(1500)
        _wait_loading_disappeared(page, timeout_ms=8000, log=print, allow_modal=True)
        page.screenshot(path=str(ss / 'step5_after_5_1.png'), full_page=True)
        out['step5_after_5_1'] = page.evaluate(DUMP_JS)
        tlog(f"after 5.1 heading: {out['step5_after_5_1'].get('heading','')[:80]}")

        # ---- 5.2 click ถัดไป ----
        tlog('Step 5.2 click ถัดไป...')
        click_next(page)
        page.wait_for_timeout(1500)
        _wait_loading_disappeared(page, timeout_ms=8000, log=print, allow_modal=True)
        page.screenshot(path=str(ss / 'step5_after_5_2.png'), full_page=True)
        out['step5_after_5_2'] = page.evaluate(DUMP_JS)
        tlog(f"after 5.2 heading: {out['step5_after_5_2'].get('heading','')[:80]}")
        print('[probe] checkboxes on 5.3 page:')
        for c in out['step5_after_5_2'].get('checkboxes', []):
            if c.get('visible'):
                print(f"   id={c['id']!r} name={c['name']!r} checked={c['checked']} | {c['label'][:80]} | {c['near'][:80]}")

        # บันทึกก่อน เผื่อ probe ค้าง
        Path('step5_probe.json').write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding='utf-8')
        tlog('saved step5_probe.json (initial)')

        # ---- 5.3 check consent checkbox + click consentButton ----
        tlog('Step 5.3 check consent checkbox + click...')
        checked = page.evaluate(r"""() => {
            const cb = document.querySelector('input[name="consent"]');
            if (!cb) return 'NO_CB';
            if (cb.checked) return 'ALREADY';
            // ลอง native click ก่อน (จะ trigger UI handlers ทั้งหมด)
            try { cb.click(); } catch(e){}
            if (cb.checked) return 'OK_CLICK';
            // ถ้ายังไม่ติ๊ก → set property + fire change
            cb.checked = true;
            cb.dispatchEvent(new Event('change', { bubbles: true }));
            cb.dispatchEvent(new Event('input', { bubbles: true }));
            return cb.checked ? 'OK_FORCE' : 'FAIL';
        }""")
        print('[probe] consent check result:', checked)
        page.wait_for_timeout(800)
        page.screenshot(path=str(ss / 'step5_3_consented.png'), full_page=True)

        # กด consentButton
        clicked = False
        try:
            cb_btn = page.locator('#consentButton').first
            if cb_btn.count() > 0:
                cb_btn.scroll_into_view_if_needed(timeout=2000)
                cb_btn.click(timeout=4000)
                clicked = True
        except Exception as e:
            print('[probe] consentButton click err:', str(e)[:100])
        if not clicked:
            page.evaluate(r"""() => { const b = document.getElementById('consentButton'); if (b) b.click(); }""")
        page.wait_for_timeout(3500)
        _wait_loading_disappeared(page, timeout_ms=15000, log=print, allow_modal=True)
        page.wait_for_timeout(1500)
        page.screenshot(path=str(ss / 'step6_initial.png'), full_page=True)
        out['step6_initial'] = page.evaluate(DUMP_JS)
        tlog(f"step6 heading: {out['step6_initial'].get('heading','')[:80]}")
        Path('step5_probe.json').write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding='utf-8')

        # ดูปุ่มอัปโหลดภาพใน Step 6
        print('[probe] step6 buttons:')
        for b in out['step6_initial'].get('buttons', []):
            t = b.get('text', '')
            if any(kw in t for kw in ['ภาพ', 'อัปโหลด', 'อัพโหลด', 'ยืนยัน', 'ถัดไป', 'ยกเลิก', 'ก่อนหน้า', 'ถ่าย', 'เลือก']):
                print(f"   {t!r:50} id={b.get('id','')!r:24} da={b.get('da','')!r:18} cls={b.get('cls','')[:60]}")

        # ---- ทดลองกด 'อัปโหลดภาพ' ใน Step 6 ----
        tlog('Step 6 — click อัปโหลดภาพ...')
        opened_up = False
        try:
            up = page.locator(
                "button:visible:has-text('อัปโหลดภาพ'), button:visible:has-text('อัพโหลดภาพ'), "
                "a:visible:has-text('อัปโหลดภาพ'), button:visible:has-text('อัปโหลด'), "
                "a:visible:has-text('อัปโหลด')"
            ).first
            if up.count() > 0:
                up.scroll_into_view_if_needed(timeout=2000)
                up.click(timeout=4000)
                opened_up = True
                print('[probe] clicked อัปโหลดภาพ')
        except Exception as e:
            print('[probe] อัปโหลดภาพ click err:', str(e)[:100])
        page.wait_for_timeout(1500)
        page.screenshot(path=str(ss / 'step6_upload_clicked.png'), full_page=True)
        out['step6_upload_clicked'] = page.evaluate(DUMP_JS)

        # หา file input ที่เพิ่ง visible (สำหรับอัปโหลดภาพหน้า/selfie)
        print('[probe] step6 visible fileInputs after click:')
        for f in out['step6_upload_clicked'].get('fileInputs', []):
            if f.get('visible') or 'image' in (f.get('accept', '') or '').lower():
                print(f"   id={f.get('id','')!r} name={f.get('name','')!r} accept={f.get('accept','')!r} visible={f.get('visible')}")
                for da in f.get('dataAttrs', []):
                    print(f"       {da[:120]}")
        # buttons in upload dialog
        print('[probe] step6 buttons after click:')
        for b in out['step6_upload_clicked'].get('buttons', []):
            t = b.get('text', '')
            if any(kw in t for kw in ['ยืนยัน', 'ยกเลิก', 'ถ่าย', 'เลือกไฟล์', 'อัปโหลด', 'อัพโหลด', 'บันทึก', 'ตกลง']):
                print(f"   {t!r:50} id={b.get('id','')!r:24} da={b.get('da','')!r:18} cls={b.get('cls','')[:60]}")

        Path('step5_probe.json').write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding='utf-8')
        tlog('saved step5_probe.json (with step6)')

        # ---- หาไฟล์ภาพหน้า จาก rec['docs'] ที่ th_name มี "รูปถ่าย" ----
        photo_path = None
        for d in rec.get('docs', []):
            if 'รูปถ่าย' in (d.get('th_name','') or ''):
                photo_path = d.get('path')
                break
        print(f'[probe] photo path: {photo_path}')

        if photo_path and Path(photo_path).exists():
            tlog('Step 6 — set image into #upload-popup...')
            try:
                # set_input_files ใช้ได้แม้ input ซ่อน
                page.set_input_files('#upload-popup', str(photo_path))
                print('[probe] ✓ set #upload-popup')
            except Exception as e:
                print('[probe] set_input_files err:', str(e)[:120])
                # fallback: ใช้ file chooser ผ่านปุ่ม "เลือกไฟล์"
                try:
                    with page.expect_file_chooser() as fc:
                        page.locator("button:visible:has-text('เลือกไฟล์'), a:visible:has-text('เลือกไฟล์'), label:visible:has-text('เลือกไฟล์')").first.click()
                    fc.value.set_files(str(photo_path))
                    print('[probe] ✓ via file chooser')
                except Exception as e2:
                    print('[probe] file chooser err:', str(e2)[:120])
            page.wait_for_timeout(1500)
            page.screenshot(path=str(ss / 'step6_after_select_file.png'), full_page=True)
            out['step6_after_select_file'] = page.evaluate(DUMP_JS)

            # คลิกยืนยัน (ใน modal: btn btn-primary float-right ที่ text มี "ยืนยัน")
            tlog('Step 6 — click ยืนยัน in modal...')
            try:
                ok_btn = page.locator(
                    ".modal:visible button:has-text('ยืนยัน'), "
                    ".swal2-popup:visible button:has-text('ยืนยัน'), "
                    "[role='dialog']:visible button:has-text('ยืนยัน'), "
                    "button.btn-primary.float-right:visible:has-text('ยืนยัน')"
                ).first
                if ok_btn.count() > 0:
                    ok_btn.click(timeout=4000)
                    print('[probe] ✓ clicked ยืนยัน')
                else:
                    page.evaluate(r"""() => {
                        const b = Array.from(document.querySelectorAll('button')).find(x => /ยืนยัน/.test((x.textContent||'').trim()) && x.offsetParent !== null);
                        if (b) b.click();
                    }""")
                    print('[probe] ✓ ยืนยัน via JS fallback')
            except Exception as e:
                print('[probe] ยืนยัน err:', str(e)[:120])

            page.wait_for_timeout(4000)
            _wait_loading_disappeared(page, timeout_ms=20000, log=print, allow_modal=True)
            page.wait_for_timeout(3000)  # รอให้ระบบประมวลผลผลลัพธ์
            page.screenshot(path=str(ss / 'step6_after_confirm.png'), full_page=True)
            out['step6_after_confirm'] = page.evaluate(DUMP_JS)

            # อ่านสถานะ
            body_after = out['step6_after_confirm'].get('bodySample', '') or ''
            status = None
            if 'ยืนยันตัวตนสำเร็จ' in body_after:
                status = 'ยืนยันตัวตนสำเร็จ'
            elif 'ยืนยันตัวตนไม่สำเร็จ' in body_after:
                status = 'ยืนยันตัวตนไม่สำเร็จ'
            elif 'ไม่สำเร็จ' in body_after:
                status = 'ไม่สำเร็จ (อื่น)'
            elif 'สำเร็จ' in body_after:
                status = 'สำเร็จ (อื่น)'
            print(f'[probe] STATUS DETECTED: {status!r}')
            # ค้นหา substring รอบๆ คำว่า "ยืนยันตัวตน"
            import re
            for m in re.finditer(r'.{0,40}ยืนยันตัวตน.{0,80}', body_after):
                print(f'   ctx: {m.group(0)!r}')

            print('[probe] step6_after_confirm visible buttons:')
            for b in out['step6_after_confirm'].get('buttons', []):
                t = b.get('text','')
                if any(kw in t for kw in ['ยืนยัน','ถัดไป','ก่อนหน้า','ยกเลิก','ใหม่','อัปโหลด','ถ่าย']):
                    print(f"   {t!r:40} id={b.get('id','')!r:24} da={b.get('da','')!r:18} cls={b.get('cls','')[:60]}")
        else:
            print(f'[probe] ⚠ ไม่พบไฟล์ภาพ: {photo_path}')

        Path('step5_probe.json').write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding='utf-8')
        tlog('saved step5_probe.json FINAL')

        tlog('PAUSE 30s for inspection...')
        page.wait_for_timeout(30000)
    finally:
        ctx.close()
        browser.close()
