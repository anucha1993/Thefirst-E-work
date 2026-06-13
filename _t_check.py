# -*- coding: utf-8 -*-
"""ตรวจว่าแรงงาน 3 คนยังเลือกได้ในฟอร์มไหม (รอบก่อนอาจสร้างร่างค้าง)"""
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

DUMP_LIST = r"""() => {
    const isVis = el => el && el.offsetParent !== null;
    const tbls = Array.from(document.querySelectorAll('table')).filter(isVis);
    const out = [];
    tbls.forEach((t, ti) => {
        const rows = Array.from(t.querySelectorAll('tbody tr')).map(r => {
            const cells = Array.from(r.querySelectorAll('td')).map(c=>(c.textContent||'').replace(/\s+/g,' ').trim());
            const cb = r.querySelector('input[type=checkbox]');
            return {cells, hasCb: !!cb, cbId: cb ? cb.id : '', checked: cb ? cb.checked : null};
        }).filter(r => r.cells.length);
        if (rows.length) out.push({tableIndex: ti, id: t.id, rowCount: rows.length, rows: rows.slice(0, 60)});
    });
    return out;
}"""

result = {"targets": WPS, "branches": []}
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
    page.wait_for_timeout(1500)
    brs = S._inform_collect_branches(page)
    result["branchCount"] = len(brs)
    for bi, b in enumerate(brs, start=1):
        S._inform_select_branch(page, b["value"], log=print)
        page.wait_for_timeout(2000)
        tables = page.evaluate(DUMP_LIST)
        # หา wp ที่ปรากฏในสาขานี้
        present = []
        for t in tables:
            for r in t["rows"]:
                joined = " ".join(r["cells"])
                for wp in WPS:
                    if wp in joined.replace(" ", ""):
                        present.append({"wp": wp, "checked": r.get("checked"), "cells": r["cells"][:6]})
        result["branches"].append({
            "index": bi, "text": b["text"][:80], "value": b["value"],
            "tables": tables, "targetsPresent": present,
        })
    # screenshot ของสาขาสุดท้าย
    page.screenshot(path=str(S.ROOT / "_check_step2.png"), full_page=True)
    # เช็กว่ามีเมนู/ลิงก์ร่างคำขอไหม
    drafts = page.evaluate(r"""() => {
        const isVis = el => el && el.offsetParent !== null;
        return Array.from(document.querySelectorAll('a,button')).filter(isVis)
            .map(a=>({txt:(a.textContent||'').replace(/\s+/g,' ').trim(), href:a.getAttribute('href')||''}))
            .filter(x=>/ร่าง|ฉบับร่าง|draft|ติดตาม|คำขอของฉัน|รายการคำขอ/i.test(x.txt));
    }""")
    result["draftLinks"] = drafts
    browser.close()

(S.ROOT / "_check_out.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
print("WROTE _check_out.json")
