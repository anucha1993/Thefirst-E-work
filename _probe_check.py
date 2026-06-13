# -*- coding: utf-8 -*-
"""READ-ONLY probe: ตรวจว่าแรงงาน 3 คนถูกยื่นไปแล้วหรือยังค้างเป็นร่าง
- ไม่กดส่ง ไม่บันทึกอะไรทั้งสิ้น
- dump เมนูหน้าหลัก (หา 'รายการคำขอ/ติดตาม') + รายชื่อแรงงานทั้งหมดทุกสาขา
ผลออกไฟล์ _probe_out.json (อ่านผ่าน read_file)
"""
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

out = {"target_wps": WPS, "menu_links": [], "branches": [], "errors": []}


def dump_menu(page):
    return page.evaluate(r"""() => {
        const isVis = el => el && el.offsetParent !== null;
        return Array.from(document.querySelectorAll('a,button')).filter(isVis).map(a => ({
            txt: (a.textContent||'').replace(/\s+/g,' ').trim().slice(0,60),
            href: a.getAttribute('href') || '',
            id: a.id || '',
            onclick: (a.getAttribute('onclick')||'').slice(0,80)
        })).filter(x => x.txt && /คำขอ|ติดตาม|สถานะ|ร่าง|ประวัติ|รายการ|ดำเนินการ|งานของฉัน/.test(x.txt));
    }""")


def dump_worker_table(page):
    return page.evaluate(r"""() => {
        const isVis = el => el && el.offsetParent !== null;
        const tbls = Array.from(document.querySelectorAll('table')).filter(isVis);
        const out = [];
        for (const t of tbls) {
            const heads = Array.from(t.querySelectorAll('thead th, thead td')).map(h=>(h.textContent||'').replace(/\s+/g,' ').trim());
            const rows = [];
            for (const tr of Array.from(t.querySelectorAll('tbody tr'))) {
                const cells = Array.from(tr.querySelectorAll('td')).map(td=>(td.textContent||'').replace(/\s+/g,' ').trim());
                const cb = tr.querySelector('input[type=checkbox]');
                rows.push({cells, hasCheckbox: !!cb, cbDisabled: cb ? cb.disabled : null, cbChecked: cb ? cb.checked : null});
            }
            if (rows.length) out.push({heads, rowCount: rows.length, rows: rows.slice(0,40)});
        }
        return out;
    }""")


with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, args=["--ignore-certificate-errors", "--start-maximized"])
    ctx = browser.new_context(locale="th-TH", ignore_https_errors=True,
                              viewport={"width": 1920, "height": 1080}, accept_downloads=True)
    page = ctx.new_page()
    S.login(page, login_cfg)
    page.wait_for_timeout(1500)

    # 1) dump เมนูหน้าหลัก
    try:
        page.goto("https://eworkpermit.doe.go.th/", wait_until="domcontentloaded", timeout=30_000)
        page.wait_for_timeout(2500)
        out["home_url"] = page.url
        out["menu_links"] = dump_menu(page)
        # เปิดเมนูบริการเผื่อมีลิงก์ซ่อน
        page.evaluate(r"""() => { const a=document.querySelector('a.lang_menu_service'); if(a) a.click(); }""")
        page.wait_for_timeout(1200)
        more = dump_menu(page)
        seen = {(m["txt"], m["href"]) for m in out["menu_links"]}
        for m in more:
            if (m["txt"], m["href"]) not in seen:
                out["menu_links"].append(m)
    except Exception as e:
        out["errors"].append(f"menu: {e}")

    # 2) เปิดฟอร์ม → ค้นหานายจ้าง → ขั้นที่ 2 → dump รายชื่อแรงงานทุกสาขา
    try:
        S._open_inform_form(page, log=print)
        S._inform_search_employer(page, rec["emp_type"], rec["emp_id"], log=print)
        S._inform_consent_first(page, log=print)
        S._inform_click_next_step1(page, log=print)
        page.wait_for_timeout(1500)
        brs = S._inform_collect_branches(page)
        out["branch_count"] = len(brs)
        for b in brs:
            S._inform_select_branch(page, b["value"], log=print)
            page.wait_for_timeout(1500)
            tbl = dump_worker_table(page)
            # หาว่า WP เป้าหมายอยู่ในสาขานี้ไหม
            found = {}
            flat = json.dumps(tbl, ensure_ascii=False)
            for wp in WPS:
                found[wp] = wp in flat
            out["branches"].append({"branch": b.get("text", "")[:60], "value": b.get("value", ""),
                                     "found_targets": found, "tables": tbl})
    except Exception as e:
        out["errors"].append(f"form: {e}")

    (S.ROOT / "_probe_out.json").write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print("PROBE DONE -> _probe_out.json")
    page.wait_for_timeout(500)
    browser.close()
