"""สำรวจ Step 4–6 ของ บต.30 (หลังขั้นแนบเอกสาร 2.9).
เดินครบ _bt30_do_step2 (ถึง 2.9 attach-next) แล้ว dump DOM ของหน้าปัจจุบัน
ไล่ไปทีละหน้า: หน้าที่กด 'ถัดไป' เฉยๆ → กดต่อ; หน้าที่มี checkbox → ติ๊กทุกช่องแล้วกดถัดไป;
หน้าที่มีปุ่ม 'อัพโหลดภาพ'/ไฟล์ → เปิดโมดอลมาดูแล้ว 'หยุด' (ไม่กดบันทึก/ไม่กดถัดไป/ไม่ส่งคำขอ/ไม่ชำระเงิน).
บันทึกผลที่ _probe_bt30_step456_out.json   (ปิดเบราว์เซอร์เองอัตโนมัติ)
รัน: .venv\\Scripts\\python.exe _probe_bt30_step456.py
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

DUMP_JS = r"""() => {
    const vis = el => { const r = el.getBoundingClientRect(); const s = getComputedStyle(el);
        return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden'; };
    const txt = el => (el.textContent || '').replace(/\s+/g, ' ').trim();
    const heads = Array.from(document.querySelectorAll('h1,h2,h3,h4,h5'))
        .filter(vis).map(e => txt(e).slice(0, 100)).filter(Boolean).slice(0, 15);
    const btns = Array.from(document.querySelectorAll('button,a,input[type=button],input[type=submit]'))
        .filter(vis).map(e => ({
            tag: e.tagName.toLowerCase(), id: e.id || '', cls: (e.className || '').toString().slice(0, 70),
            text: txt(e).slice(0, 50), onclick: (e.getAttribute('onclick') || '').slice(0, 120),
            disabled: !!e.disabled
        })).slice(0, 70);
    const checks = Array.from(document.querySelectorAll('input[type=checkbox]')).map(e => {
        let lbl = ''; if (e.id) { const l = document.querySelector('label[for="' + (window.CSS && CSS.escape ? CSS.escape(e.id) : e.id) + '"]'); if (l) lbl = txt(l); }
        if (!lbl && e.closest('label')) lbl = txt(e.closest('label'));
        if (!lbl && e.parentElement) lbl = txt(e.parentElement);
        return { id: e.id || '', name: e.name || '', checked: e.checked, vis: vis(e), label: lbl.slice(0, 200) };
    }).slice(0, 25);
    const files = Array.from(document.querySelectorAll('input[type=file]')).map(e => ({
        id: e.id || '', name: e.name || '', onchange: (e.getAttribute('onchange') || '').slice(0, 100),
        accept: e.accept || '', vis: vis(e)
    })).slice(0, 25);
    const uploads = Array.from(document.querySelectorAll('button,a,label,div,span'))
        .filter(e => /อัพโหลด|อัปโหลด|upload/i.test(e.textContent || '') && vis(e) && (e.textContent || '').trim().length < 40)
        .map(e => ({ tag: e.tagName.toLowerCase(), id: e.id || '', text: txt(e).slice(0, 50), onclick: (e.getAttribute('onclick') || '').slice(0, 120) }))
        .slice(0, 15);
    const steps = Array.from(document.querySelectorAll('[class*=step],[class*=Step],[class*=wizard],[class*=progress],[class*=stepper]'))
        .filter(vis).map(e => txt(e).slice(0, 140)).filter(t => t && t.length < 140).slice(0, 12);
    return { url: location.href, heads, btns, checks, files, uploads, steps };
}"""

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

CHECK_ALL_JS = r"""() => {
    const vis = el => { const r = el.getBoundingClientRect(); const s = getComputedStyle(el);
        return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden'; };
    let n = 0;
    Array.from(document.querySelectorAll('input[type=checkbox]')).forEach(e => {
        // ติ๊กแม้ input ถูกซ่อนแต่มี label/parent ที่มองเห็น (ธีม custom checkbox)
        const shown = vis(e) || (e.parentElement && vis(e.parentElement)) || (e.closest('label') && vis(e.closest('label')));
        if (shown && !e.checked) {
            e.click();
            if (!e.checked) { e.checked = true; e.dispatchEvent(new Event('change', { bubbles: true })); }
            n++;
        }
    });
    return n;
}"""

PAY_RE = "ชำระเงิน|ชำระค่า|ยืนยันการยื่น|ส่งคำขอ|ยื่นคำขอ|submit|payment|ชำระ"


def main():
    accounts = _read_login_accounts(ROOT / "UsernameLogin.xlsx")
    acct = next(iter(accounts.values()))
    cfg = {"username": acct["username"], "password": acct["password"],
           "user_type": acct["type"], "method": acct.get("method") or "E-Workpermit"}
    rec = _read_bt30_excel(ROOT / "from_bt30.xlsx")[0]

    dumps: list[dict] = []
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
            _close_register_alert(page)
            page.wait_for_timeout(2500)
            _bt30_dismiss_news(page, log=log)

            r2 = _bt30_do_step2(page, rec, SHOTS, log=log)
            log(f"[i] ขั้นตอน2: {r2.get('step2_status')} — แนบเอกสารถึง 2.9")
            if r2.get("step2_status") != "SUCCESS":
                log("[!] step2 ไม่ SUCCESS — หยุด (ไม่เดินต่อ step4-6)")
                (ROOT / "_probe_bt30_step456_out.json").write_text(
                    json.dumps({"aborted": r2.get("step2_note")}, ensure_ascii=False, indent=2), encoding="utf-8")
                return

            page.wait_for_timeout(2500)
            _bt30_dismiss_news(page, log=log)

            for phase in range(7):
                d = page.evaluate(DUMP_JS)
                d["phase"] = phase
                dumps.append(d)
                vis_checks = [c for c in d["checks"] if c["vis"] or c["label"]]
                has_upload = bool(d["uploads"]) or any(f["vis"] for f in d["files"])
                pay_btns = [b for b in d["btns"]
                            if __import__("re").search(PAY_RE, b["text"], __import__("re").I)]
                log(f"[dump {phase}] url=...{d['url'][-50:]} | heads={d['heads'][:2]} | "
                    f"checks={len(vis_checks)} upload={has_upload} pay={len(pay_btns)} | steps={d['steps'][:1]}")

                if pay_btns:
                    log(f"[STOP] เจอปุ่มชำระเงิน/ส่งคำขอ: {[b['text'] for b in pay_btns][:3]} → ไม่กดอะไรต่อ")
                    break

                if has_upload:
                    log(f"[Step6] หน้าอัพโหลดภาพ — เปิดโมดอลมาดู (uploads={[u['text'] for u in d['uploads']][:3]})")
                    opened = page.evaluate(r"""() => {
                        const vis = el => { const r = el.getBoundingClientRect(); const s = getComputedStyle(el);
                            return r.width>0 && r.height>0 && s.display!=='none' && s.visibility!=='hidden'; };
                        const t = Array.from(document.querySelectorAll('button,a,label,div,span'))
                          .find(e => /อัพโหลด|อัปโหลด/.test(e.textContent||'') && vis(e) && (e.textContent||'').trim().length < 40);
                        if (t) { t.click(); return (t.textContent||'').trim().slice(0,40); }
                        return '';
                    }""")
                    log(f"[Step6] คลิก '{opened}' แล้ว — รอโมดอล")
                    page.wait_for_timeout(1800)
                    dm = page.evaluate(DUMP_JS)
                    dm["phase"] = f"{phase}_modal"
                    # เก็บ HTML ของ modal ที่กำลังแสดง เพื่อดูปุ่มบันทึก/ช่องอัพโหลด
                    dm["modal_html"] = page.evaluate(r"""() => {
                        const ms = Array.from(document.querySelectorAll('.modal.show,[role=dialog]'))
                          .filter(m => { const r = m.getBoundingClientRect(); return r.width>0 && r.height>0; });
                        return ms.map(m => m.outerHTML).join('\n<!-- ---- -->\n').slice(0, 6000);
                    }""")
                    dumps.append(dm)
                    log(f"[Step6] โมดอล: btns={[b['text'] for b in dm['btns'] if b['text']][:8]}")
                    log("[STOP] หยุดที่ Step 6 (ไม่กดบันทึก/ไม่กดถัดไป)")
                    break

                if vis_checks:
                    n = page.evaluate(CHECK_ALL_JS)
                    log(f"[Step5.2] ติ๊ก checkbox {n} ช่อง (id={[c['id'] for c in vis_checks][:3]})")
                    page.wait_for_timeout(500)

                clicked = page.evaluate(CLICK_NEXT_JS)
                if not clicked:
                    log(f"[STOP] ไม่พบปุ่ม 'ถัดไป' ที่ phase {phase}")
                    break
                log(f"[phase {phase}] กด 'ถัดไป' → หน้าถัดไป")
                page.wait_for_timeout(3000)
                _bt30_dismiss_news(page, log=log)

            (ROOT / "_probe_bt30_step456_out.json").write_text(
                json.dumps(dumps, ensure_ascii=False, indent=2), encoding="utf-8")
            log(f"[i] บันทึก {len(dumps)} dumps → _probe_bt30_step456_out.json")
            try:
                page.screenshot(path=str(SHOTS / "_step456_last.png"), full_page=True)
            except Exception:
                pass
        except Exception as e:
            log("[ERR]", e)
            try:
                (ROOT / "_probe_bt30_step456_out.json").write_text(
                    json.dumps(dumps, ensure_ascii=False, indent=2), encoding="utf-8")
                page.screenshot(path=str(SHOTS / "_step456_error.png"), full_page=True)
            except Exception:
                pass
        finally:
            page.wait_for_timeout(800)
            ctx.close(); browser.close()


if __name__ == "__main__":
    main()
