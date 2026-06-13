# -*- coding: utf-8 -*-
import json
from playwright.sync_api import sync_playwright
import scrape_wa as S

cfg = S.load_config()
accts = S._read_login_accounts(S.ROOT / "UsernameLogin.xlsx")
acct = next(iter(accts.values()))
login_cfg = dict(cfg)
login_cfg["username"] = acct["username"]
login_cfg["password"] = acct.get("password") or cfg.get("password", "")
login_cfg["user_type"] = acct.get("type") or cfg.get("user_type", "")
login_cfg["method"] = acct.get("method") or cfg.get("method", "E-Service")

data = S._read_inform_data(S.ROOT / "FormRequestEmployment.xlsx")
rec = data[0]

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, args=["--ignore-certificate-errors", "--start-maximized"])
    ctx = browser.new_context(locale="th-TH", ignore_https_errors=True,
                              viewport={"width": 1920, "height": 1080}, accept_downloads=True)
    page = ctx.new_page()
    S.login(page, login_cfg)
    S._open_inform_form(page, log=print)
    print("[+] form opened")

    # ใช้ฟังก์ชันจริงในการค้นหา+บันทึก
    ok = S._inform_search_employer(page, rec["emp_type"], rec["emp_id"], log=print)
    print("search_employer ->", ok)
    page.wait_for_timeout(2000)
    page.screenshot(path=str(S.ROOT / "_diag_after_save.png"), full_page=True)

    info = page.evaluate(r"""() => {
        const re = /RE\d{10,}/;
        const txt = document.body.innerText || '';
        const m = txt.match(/นายจ้าง\s*(\d+)\s*รายการ/);
        const out = {empCountText: m ? m[0] : null, visBtns: [], reEls: [], visibleModalText: null};
        const md = document.querySelector('.modal.show, .swal2-popup');
        if (md && md.offsetParent !== null) out.visibleModalText = (md.innerText||'').replace(/\s+/g,' ').slice(0,300);
        out.visBtns = Array.from(document.querySelectorAll('button, a.btn, .btn'))
            .filter(b => b.offsetParent !== null)
            .map(b => ({id:b.id, txt:(b.textContent||'').replace(/\s+/g,' ').trim()}))
            .filter(b => b.txt && b.txt.length < 30);
        const all = Array.from(document.querySelectorAll('td, span, label'));
        for (const e of all) {
            const t = (e.childNodes.length===1) ? (e.textContent||'') : '';
            if (re.test(t)) out.reEls.push({tag:e.tag, txt:t.replace(/\s+/g,' ').trim().slice(0,80),
                row: (e.closest('tr')||{}).innerText?.replace(/\s+/g,' ').slice(0,200) || ''});
        }
        return out;
    }""")
    print("=== AFTER SAVE ===")
    print(json.dumps(info, ensure_ascii=False, indent=2))

    # ลองติ๊ก consent + กดถัดไป แล้วดูว่า URL/หน้าจอเปลี่ยนไหม
    S._inform_consent_first(page, log=print)
    consent_checked = page.evaluate(r"""() => {
        const cbs = Array.from(document.querySelectorAll('input[type=checkbox]'));
        for (const cb of cbs) {
            let p=cb.parentElement, t='', lvl=0;
            while(p&&lvl<5){ t += ' ' + (p.textContent||''); p=p.parentElement; lvl++; }
            if (/ขอรับรองว่า|มีความประสงค์ในการยื่นคำขอ/.test(t)) return cb.checked;
        }
        return null;
    }""")
    print("consent checked ->", consent_checked)
    before = page.url
    S._inform_click_next_step1(page, log=print)
    page.wait_for_timeout(2500)
    page.screenshot(path=str(S.ROOT / "_diag_after_next.png"), full_page=True)
    # dump หน้า step 2: select สาขาที่มองเห็น + ปุ่มเลือกสาขา + ตารางแรงงาน
    step2 = page.evaluate(r"""() => {
        const isVis = el => el && el.offsetParent !== null;
        const out = {visSelects: [], visBtns: [], workerTable: null};
        out.visSelects = Array.from(document.querySelectorAll('select')).filter(isVis).map(s=>({
            id:s.id, name:s.name, optCount:s.options.length,
            opts: Array.from(s.options).slice(0,8).map(o=>(o.textContent||'').replace(/\s+/g,' ').trim())
        }));
        out.visBtns = Array.from(document.querySelectorAll('button, a.btn, .btn')).filter(isVis)
            .map(b=>({id:b.id, txt:(b.textContent||'').replace(/\s+/g,' ').trim()})).filter(b=>b.txt&&b.txt.length<30);
        // ตารางแรงงานที่มองเห็น
        const tbls = Array.from(document.querySelectorAll('table')).filter(isVis);
        for (const t of tbls) {
            const hs = Array.from(t.querySelectorAll('thead th, thead td')).map(h=>(h.textContent||'').replace(/\s+/g,' ').trim());
            if (hs.some(h=>/ใบอนุญาตทำงาน|วันที่จ้าง|ชื่อ-นามสกุล/.test(h))) {
                out.workerTable = {headers: hs, rows: Array.from(t.querySelectorAll('tbody tr')).slice(0,3).map(r=>Array.from(r.querySelectorAll('td')).map(c=>(c.textContent||'').replace(/\s+/g,' ').trim()))};
                break;
            }
        }
        return out;
    }""")
    print("URL changed:", before, "->", page.url)
    print("=== STEP 2 VISIBLE ===")
    print(json.dumps(step2, ensure_ascii=False, indent=2))

    # เลือกสาขาแรก แล้วดูตารางแรงงาน
    brs = S._inform_collect_branches(page)
    print("branches:", len(brs))
    for b in brs:
        print("   value=", b['value'], "text=", b['text'][:60])
    for bi, b in enumerate(brs):
        S._inform_select_branch(page, b['value'], log=print)
        page.wait_for_timeout(2500)
        info = page.evaluate(r"""() => {
            const isVis = el => el && el.offsetParent !== null;
            const tbls = Array.from(document.querySelectorAll('table')).filter(isVis);
            for (const t of tbls) {
                const hs = Array.from(t.querySelectorAll('thead th, thead td')).map(h=>(h.textContent||'').replace(/\s+/g,' ').trim());
                if (hs.some(h=>/ใบอนุญาตทำงาน/.test(h))) {
                    const rows = Array.from(t.querySelectorAll('tbody tr'));
                    return {rowCount: rows.length, firstRow: rows.length? rows[0].innerText.replace(/\s+/g,' ').slice(0,120):'',
                        realRows: rows.filter(r=>r.querySelector('input[type=checkbox]')).map(r=>({
                            wp: Array.from(r.querySelectorAll('td')).map(c=>(c.textContent||'').replace(/\s+/g,' ').trim()),
                            cbId:(r.querySelector('input[type=checkbox]')||{}).id||''})).slice(0,5)};
                }
            }
            return null;
        }""")
        print(f"  branch[{bi}] value={b['value']} -> ", json.dumps(info, ensure_ascii=False))
        page.screenshot(path=str(S.ROOT / f"_diag_branch{bi}.png"), full_page=True)
    page.wait_for_timeout(1500)
    browser.close()



