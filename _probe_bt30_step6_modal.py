"""สำรวจโมดอล 'อัปโหลดภาพ' (OncOpenModalAddFileIden) ของ Step 6 'วิธีการยืนยันตัวตน'.
เดินตรงไปยัง Step 6 ด้วย selector ที่รู้แล้ว:
  Step4  ถัดไป  #NextStepThreePageOneRenew
  Step5.1 ถัดไป (ปุ่มไม่มี id, class .nextStep) → คลิกปุ่ม 'ถัดไป' ที่มองเห็น/ใช้งานได้
  Step5.2 ติ๊ก #check_truth → ปุ่ม #nextstepfour59 จะ enable แล้วกด
  Step6  คลิกปุ่ม onclick=OncOpenModalAddFileIden() → dump โมดอล (ปุ่มบันทึก/ช่องไฟล์) + ซอร์สฟังก์ชัน JS
หยุดแค่นี้ — ไม่กดบันทึก/ไม่กดถัดไป/ไม่ส่งคำขอ/ไม่ชำระเงิน. ปิดเบราว์เซอร์เอง.
บันทึก: _probe_bt30_step6_modal_out.json
รัน: .venv\\Scripts\\python.exe _probe_bt30_step6_modal.py
"""
from __future__ import annotations

import json

from playwright.sync_api import sync_playwright

from scrape_wa import (
    login, _read_login_accounts, _read_bt30_excel, _open_bt30_form,
    _bt30_fill_search_one, _close_register_alert, _bt30_dismiss_news,
    _bt30_do_step2, ROOT,
)

SHOTS = ROOT / "reports" / "bt30_screenshots"
SHOTS.mkdir(parents=True, exist_ok=True)
log = print

CLICK_NEXT_JS = r"""() => {
    const vis = el => { const r = el.getBoundingClientRect(); const s = getComputedStyle(el);
        return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden'; };
    const cand = Array.from(document.querySelectorAll('button,a'))
        .filter(e => vis(e) && !e.disabled && /ถัดไป/.test((e.textContent || '').trim())
            && (e.textContent || '').trim().length < 20
            && !/ย้อนกลับ|ยกเลิก/.test((e.textContent || '').trim()));
    if (cand.length) { cand[cand.length - 1].click(); return true; }
    return false;
}"""

MODAL_DUMP_JS = r"""() => {
    const vis = el => { const r = el.getBoundingClientRect(); const s = getComputedStyle(el);
        return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden'; };
    const txt = el => (el.textContent || '').replace(/\s+/g, ' ').trim();
    const modals = Array.from(document.querySelectorAll('.modal.show,.modal[style*="display: block"],[role=dialog],.swal2-popup'))
        .filter(m => { const r = m.getBoundingClientRect(); return r.width > 0 && r.height > 0; });
    const result = modals.map(m => ({
        id: m.id || '', cls: (m.className || '').toString().slice(0, 120),
        btns: Array.from(m.querySelectorAll('button,a,input[type=button],input[type=submit]')).map(b => ({
            tag: b.tagName.toLowerCase(), id: b.id || '', cls: (b.className || '').toString().slice(0, 90),
            text: txt(b).slice(0, 50), onclick: (b.getAttribute('onclick') || '').slice(0, 150), disabled: !!b.disabled
        })),
        files: Array.from(m.querySelectorAll('input[type=file]')).map(f => ({
            id: f.id || '', name: f.name || '', onchange: (f.getAttribute('onchange') || '').slice(0, 120), accept: f.accept || ''
        })),
        imgs: Array.from(m.querySelectorAll('img')).map(i => ({ id: i.id || '', src: (i.src || '').slice(0, 80) })).slice(0, 8),
        html: m.outerHTML.slice(0, 8000),
    }));
    return result;
}"""

FN_SRC_JS = r"""() => {
    const out = {};
    let names = [];
    try { names = Object.getOwnPropertyNames(window).filter(n => {
        try { return typeof window[n] === 'function' && /Iden|SaveAddFile|FileOrther|AddNewFile|checkTypeNameSize|selectFileupload/i.test(n); }
        catch (e) { return false; }
    }); } catch (e) {}
    names.forEach(n => { try { out[n] = window[n].toString().slice(0, 3000); } catch (e) { out[n] = 'ERR:' + e; } });
    return out;
}"""


def main():
    accounts = _read_login_accounts(ROOT / "UsernameLogin.xlsx")
    acct = next(iter(accounts.values()))
    cfg = {"username": acct["username"], "password": acct["password"],
           "user_type": acct["type"], "method": acct.get("method") or "E-Workpermit"}
    rec = _read_bt30_excel(ROOT / "from_bt30.xlsx")[0]

    out: dict = {}
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False, args=["--ignore-certificate-errors", "--start-maximized"])
        ctx = browser.new_context(locale="th-TH", ignore_https_errors=True, viewport={"width": 1600, "height": 950})
        page = ctx.new_page()
        try:
            login(page, cfg); page.wait_for_timeout(1500)
            if not _open_bt30_form(page, log=log):
                log("[!] เปิดฟอร์มไม่สำเร็จ"); return
            r1 = _bt30_fill_search_one(page, rec, SHOTS, log=log)
            log(f"[i] ขั้นตอน1: {r1.get('status')}")
            _close_register_alert(page); page.wait_for_timeout(2500)
            _bt30_dismiss_news(page, log=log)

            r2 = _bt30_do_step2(page, rec, SHOTS, log=log)
            log(f"[i] ขั้นตอน2: {r2.get('step2_status')}")
            if r2.get("step2_status") != "SUCCESS":
                log("[!] step2 ไม่ SUCCESS — หยุด")
                out["aborted"] = r2.get("step2_note")
                (ROOT / "_probe_bt30_step6_modal_out.json").write_text(
                    json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
                return
            page.wait_for_timeout(2500); _bt30_dismiss_news(page, log=log)

            # Step 4: แนบเอกสาร หน้า 2/2 (เอกสารนายจ้าง) — กดถัดไป
            page.evaluate("() => { const b = document.getElementById('NextStepThreePageOneRenew'); if (b) b.click(); }")
            log("[Step4] กด #NextStepThreePageOneRenew")
            page.wait_for_timeout(3000); _bt30_dismiss_news(page, log=log)

            # Step 5.1: สรุปคำขอ หน้า 1/2 — กดถัดไป
            page.evaluate(CLICK_NEXT_JS)
            log("[Step5.1] กดถัดไป (สรุปคำขอ 1/2)")
            page.wait_for_timeout(3000); _bt30_dismiss_news(page, log=log)

            # Step 5.2: สรุปคำขอ หน้า 2/2 — ติ๊ก #check_truth → กด #nextstepfour59
            ck = page.evaluate(r"""() => {
                const c = document.getElementById('check_truth');
                if (!c) return 'no-checkbox';
                if (!c.checked) { c.click(); if (!c.checked) { c.checked = true; c.dispatchEvent(new Event('change', {bubbles:true})); } }
                return c.checked ? 'checked' : 'fail';
            }""")
            log(f"[Step5.2] check_truth → {ck}")
            page.wait_for_timeout(800)
            nx = page.evaluate(r"""() => {
                const b = document.getElementById('nextstepfour59');
                if (!b) return 'no-btn';
                if (b.disabled) return 'disabled';
                b.click(); return 'clicked';
            }""")
            log(f"[Step5.2] nextstepfour59 → {nx}")
            page.wait_for_timeout(3500); _bt30_dismiss_news(page, log=log)

            # Step 6: วิธีการยืนยันตัวตน — คลิกปุ่มอัปโหลดภาพ (OncOpenModalAddFileIden)
            heads = page.evaluate("() => Array.from(document.querySelectorAll('h1,h2,h3,h4,h5')).map(e=>e.textContent.trim()).filter(Boolean).slice(0,5)")
            log(f"[Step6] heads={heads}")
            opened = page.evaluate(r"""() => {
                const vis = el => { const r = el.getBoundingClientRect(); const s = getComputedStyle(el);
                    return r.width>0 && r.height>0 && s.display!=='none' && s.visibility!=='hidden'; };
                const b = Array.from(document.querySelectorAll('button,a'))
                  .find(e => vis(e) && /OncOpenModalAddFileIden/.test(e.getAttribute('onclick')||''));
                if (b) { b.click(); return 'clicked:onclick'; }
                if (typeof window.OncOpenModalAddFileIden === 'function') { window.OncOpenModalAddFileIden(); return 'called:fn'; }
                return 'not-found';
            }""")
            log(f"[Step6] เปิดโมดอลอัปโหลดภาพ → {opened}")
            page.wait_for_timeout(2000)

            out["step6_modals"] = page.evaluate(MODAL_DUMP_JS)
            out["fn_src"] = page.evaluate(FN_SRC_JS)
            out["step6_open_result"] = opened
            log(f"[Step6] modal count={len(out['step6_modals'])}; "
                f"fns={list(out['fn_src'].keys())}")
            for m in out["step6_modals"]:
                log(f"  modal #{m['id']} btns={[(b['text'],b['id'],b['onclick'][:40]) for b in m['btns']]}")
                log(f"           files={[(f['id'],f['onchange'][:40]) for f in m['files']]}")

            (ROOT / "_probe_bt30_step6_modal_out.json").write_text(
                json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
            log("[i] บันทึก → _probe_bt30_step6_modal_out.json")
            try:
                page.screenshot(path=str(SHOTS / "_step6_modal.png"), full_page=True)
            except Exception:
                pass
        except Exception as e:
            log("[ERR]", e)
            try:
                (ROOT / "_probe_bt30_step6_modal_out.json").write_text(
                    json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
                page.screenshot(path=str(SHOTS / "_step6_modal_error.png"), full_page=True)
            except Exception:
                pass
        finally:
            page.wait_for_timeout(800)
            ctx.close(); browser.close()


if __name__ == "__main__":
    main()
