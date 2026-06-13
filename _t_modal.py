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
    print("[+] Login OK")
    S._open_inform_form(page, log=print)

    # เปิด modal ค้นหานายจ้าง + เลือกประเภท + กรอกเลข + กดค้นหา (ไม่กดบันทึก)
    page.evaluate(r"""() => {
        const btns = Array.from(document.querySelectorAll('.bus_type_modal'));
        const t = btns.find(b => (b.textContent||'').includes('ค้นหานายจ้าง') && b.id !== 'addEmpBtn');
        if (t) t.click();
    }""")
    page.wait_for_timeout(2000)
    page.wait_for_function(r"""() => { const s=document.querySelector('#slt_employer_type'); return s && s.options.length>1; }""", timeout=15000)
    page.evaluate(r"""(want) => {
        const s = document.querySelector('#slt_employer_type');
        const norm = t => (t||'').replace(/\s+/g,' ').trim().toLowerCase();
        const w = norm(want);
        const opt = Array.from(s.options).find(o => norm(o.textContent)===w) || Array.from(s.options).find(o=>w&&norm(o.textContent).includes(w));
        if (opt){ s.value=opt.value; s.dispatchEvent(new Event('change',{bubbles:true})); if(window.jQuery) jQuery(s).trigger('change'); }
    }""", rec["emp_type"])
    page.wait_for_timeout(800)
    page.evaluate(r"""(v) => { const t=document.querySelector('#employer_id_search'); if(t){t.value=v; t.dispatchEvent(new Event('input',{bubbles:true})); t.dispatchEvent(new Event('change',{bubbles:true}));} }""", rec["emp_id"])
    page.evaluate(r"""() => { const b=document.querySelector('.searchEmployer'); if(b) b.click(); }""")
    page.wait_for_timeout(3000)
    try:
        page.wait_for_function(r"""() => /RE\d{10,}/.test(document.body.innerText||'')""", timeout=12000)
    except Exception:
        print("!! ไม่พบผล RE")

    # dump โครงสร้าง modal ที่เปิดอยู่
    info = page.evaluate(r"""() => {
        const out = {modals: [], buttons: [], checkboxes: [], radios: []};
        // หา modal ที่กำลังแสดง
        const modals = Array.from(document.querySelectorAll('.modal'));
        for (const m of modals) {
            const style = getComputedStyle(m);
            const shown = style.display !== 'none' && m.offsetParent !== null;
            if (!shown) continue;
            out.modals.push({
                id: m.id,
                cls: m.className,
                title: (m.querySelector('.modal-title, .modal-header')||{}).textContent?.replace(/\s+/g,' ').trim() || '',
                buttons: Array.from(m.querySelectorAll('button, a.btn, .btn')).filter(b=>b.offsetParent).map(b=>({id:b.id, cls:b.className, txt:(b.textContent||'').replace(/\s+/g,' ').trim()})),
                checkboxes: Array.from(m.querySelectorAll('input[type=checkbox]')).map(c=>({id:c.id, name:c.name, checked:c.checked, vis: c.offsetParent!==null})),
                radios: Array.from(m.querySelectorAll('input[type=radio]')).map(c=>({id:c.id, name:c.name, val:c.value, vis: c.offsetParent!==null})),
                tables: Array.from(m.querySelectorAll('table')).map(t=>({
                    headers: Array.from(t.querySelectorAll('thead th, thead td')).map(h=>(h.textContent||'').replace(/\s+/g,' ').trim()),
                    rows: Array.from(t.querySelectorAll('tbody tr')).slice(0,3).map(r=>({
                        html: r.innerHTML.replace(/\s+/g,' ').slice(0,400),
                        cells: Array.from(r.querySelectorAll('td')).map(c=>(c.textContent||'').replace(/\s+/g,' ').trim())
                    }))
                }))
            });
        }
        return out;
    }""")
    print("=== MODAL DUMP ===")
    print(json.dumps(info, ensure_ascii=False, indent=2))
    page.wait_for_timeout(1500)
    browser.close()
