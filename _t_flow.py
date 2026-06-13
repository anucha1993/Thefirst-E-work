# -*- coding: utf-8 -*-
import json
from pathlib import Path
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
groups = S._group_inform_records(data)
grp = groups[0]
rec = data[0]
WPS = [w["work_permit"] for w in grp["workers"]]
files_root = S.ROOT

def dump_attach(page):
    return page.evaluate(r"""() => {
        const isVis = el => el && el.offsetParent !== null;
        const inputs = Array.from(document.querySelectorAll('input[type=file]')).map(i=>({
            id:i.id, name:i.name,
            section: (()=>{ let p=i; for(let k=0;k<8&&p;k++){p=p.parentElement; if(p&&/ผู้รับมอบอำนาจ|เอกสารนายจ้าง/.test(p.textContent||'')) return (p.textContent||'').replace(/\s+/g,' ').slice(0,60);} return ''; })(),
            hasFile: i.files && i.files.length>0
        }));
        const nextBtns = Array.from(document.querySelectorAll('button, a.btn')).filter(isVis)
            .map(b=>({id:b.id, txt:(b.textContent||'').replace(/\s+/g,' ').trim()})).filter(b=>/ถัดไป|ตัดไป|บันทึก|ย้อนกลับ/.test(b.txt));
        return {inputs, nextBtns};
    }""")

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, args=["--ignore-certificate-errors", "--start-maximized"])
    ctx = browser.new_context(locale="th-TH", ignore_https_errors=True,
                              viewport={"width": 1920, "height": 1080}, accept_downloads=True)
    page = ctx.new_page()
    S.login(page, login_cfg)
    S._open_inform_form(page, log=print)
    S._inform_search_employer(page, rec["emp_type"], rec["emp_id"], log=print)
    S._inform_consent_first(page, log=print)
    S._inform_click_next_step1(page, log=print)
    # tick ALL workers across branches
    brs = S._inform_collect_branches(page)
    for b in brs:
        S._inform_select_branch(page, b["value"], log=print)
        S._inform_check_workpermits(page, WPS, log=print)
    # beginNext2 -> แนบเอกสาร
    page.evaluate(r"""() => { const b=document.querySelector('#beginNext2'); if(b) b.click(); }""")
    page.wait_for_timeout(3000)
    S._inform_dismiss_popup(page, log=print)
    page.wait_for_timeout(1000)
    print("=== ATTACH PAGE (before our attach) ===")
    print(json.dumps(dump_attach(page), ensure_ascii=False, indent=2))

    # attach authorization
    files = [files_root / f for f in rec.get("files", [])]
    S._inform_attach_authorization(page, files, log=print)
    page.wait_for_timeout(1500)
    print("=== ATTACH PAGE (after our attach) ===")
    print(json.dumps(dump_attach(page), ensure_ascii=False, indent=2))
    page.screenshot(path=str(S.ROOT / "_flow_attach.png"), full_page=True)

    # click ถัดไป on attach page
    page.evaluate(r"""() => { const b=document.querySelector('#beginNext3'); if(b) b.click(); }""")
    page.wait_for_timeout(3000)
    S._inform_dismiss_popup(page, log=print)
    page.wait_for_timeout(1500)
    page.screenshot(path=str(S.ROOT / "_flow_step3.png"), full_page=True)
    step3 = page.evaluate(r"""() => {
        const isVis = el => el && el.offsetParent !== null;
        const out = {activeStep:'', workerTable:null, dateInputs:[], checkboxes:[], buttons:[]};
        const steps = Array.from(document.querySelectorAll('.step, [class*="step"]')).filter(isVis).map(s=>(s.textContent||'').replace(/\s+/g,' ').trim()).filter(t=>/กรอกข้อมูล|แนบเอกสาร|สรุป|เสร็จ/.test(t));
        out.activeStep = steps.join(' | ').slice(0,120);
        const tbls = Array.from(document.querySelectorAll('table')).filter(isVis);
        for (const t of tbls) {
            const hs = Array.from(t.querySelectorAll('thead th, thead td')).map(h=>(h.textContent||'').replace(/\s+/g,' ').trim());
            if (hs.some(h=>/ใบอนุญาตทำงาน|วันที่จ้าง/.test(h))) {
                out.workerTable = {headers:hs, rows: Array.from(t.querySelectorAll('tbody tr')).slice(0,3).map(r=>({
                    cells: Array.from(r.querySelectorAll('td')).map(c=>(c.textContent||'').replace(/\s+/g,' ').trim()),
                    inputs: Array.from(r.querySelectorAll('input')).map(i=>({id:i.id, ph:i.placeholder, type:i.type, cls:i.className})),
                    icons: Array.from(r.querySelectorAll('button,a,i,svg')).map(e=>({tag:e.tagName, cls:e.className, txt:(e.textContent||'').trim().slice(0,15)})).slice(0,8)
                }))};
                break;
            }
        }
        out.dateInputs = Array.from(document.querySelectorAll('input')).filter(isVis).filter(i=>/dd\/mm|วันที่/.test((i.placeholder||'')+(i.getAttribute('aria-label')||''))).map(i=>({id:i.id, ph:i.placeholder, cls:i.className}));
        out.checkboxes = Array.from(document.querySelectorAll('input[type=checkbox]')).filter(isVis).map(c=>({id:c.id, label:(c.closest('label')||c.parentElement||{}).textContent?.replace(/\s+/g,' ').trim().slice(0,80)}));
        out.buttons = Array.from(document.querySelectorAll('button,a.btn')).filter(isVis).map(b=>({id:b.id, txt:(b.textContent||'').replace(/\s+/g,' ').trim()})).filter(b=>b.txt&&b.txt.length<25);
        return out;
    }""")
    print("=== STEP 3 (สรุปคำขอ?) ===")
    print(json.dumps(step3, ensure_ascii=False, indent=2))

    # === Step 3 summary: dump ALL tables w/ วันที่จ้าง header ===
    def dump_summary():
        return page.evaluate(r"""() => {
            const isVis = el => el && el.offsetParent !== null;
            const tbls = Array.from(document.querySelectorAll('table')).filter(isVis);
            const out = [];
            tbls.forEach((t, ti) => {
                const hs = Array.from(t.querySelectorAll('thead th')).map(h=>(h.textContent||'').replace(/\s+/g,' ').trim());
                if (!hs.some(h=>/วันที่จ้าง|ใบอนุญาตทำงาน/.test(h))) return;
                out.push({
                    tableIndex: ti,
                    id: t.id, cls: t.className,
                    headers: hs,
                    rowCount: t.querySelectorAll('tbody tr').length,
                    rows: Array.from(t.querySelectorAll('tbody tr')).map(r=>({
                        cells: Array.from(r.querySelectorAll('td')).map(c=>(c.textContent||'').replace(/\s+/g,' ').trim()),
                        hasAddDate: !!r.querySelector('a.alien-addDate'),
                    })),
                });
            });
            return out;
        }""")

    before = dump_summary()
    log = print
    results = []
    for w in grp["workers"]:
        wp = w["work_permit"]; hd = w["hire_date"]
        ok = S._inform_set_hire_date(page, wp, hd, log=print)
        results.append({"wp": wp, "ok": ok})
        page.wait_for_timeout(1200)
    print("=== DATE RESULTS ===")
    print(json.dumps(results, ensure_ascii=False, indent=2))
    print("=== SUMMARY AFTER DATES ===")
    summ = dump_summary()
    print(json.dumps(summ, ensure_ascii=False, indent=2))
    page.screenshot(path=str(S.ROOT / "_flow_aftersave.png"), full_page=True)
    (S.ROOT / "_flow_out.json").write_text(json.dumps({"results":results,"summary":summ}, ensure_ascii=False, indent=2), encoding="utf-8")
    (S.ROOT / "_flow_before.json").write_text(json.dumps(before, ensure_ascii=False, indent=2), encoding="utf-8")
    page.wait_for_timeout(1500)
    browser.close()



