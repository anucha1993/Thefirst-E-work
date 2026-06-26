# -*- coding: utf-8 -*-
"""ส่อง DOM แท็บ 'เอกสารตอบรับ' ของคำขอ multi-person (บต.55) ว่าโครงสร้างจริงเป็นอย่างไร"""
import json
from playwright.sync_api import sync_playwright
import scrape_wa as s

REQ_NO = "69155200181078"
OUT = "_probe_bt55_dom_out.json"

cfg = s.load_config(require_login=False)
accounts = s._read_login_accounts(s.ROOT / "UsernameLogin.xlsx")
acct = list(accounts.values())[0]
login_cfg = {
    "username": acct["username"], "password": acct["password"],
    "user_type": acct["type"], "method": acct.get("method") or "E-Workpermit",
}

dump = {}
with sync_playwright() as pw:
    browser = s._launch_chromium(pw, cfg, ["--ignore-certificate-errors", "--start-maximized"])
    ctx = browser.new_context(locale="th-TH", ignore_https_errors=True,
                              viewport={"width": 1920, "height": 1080}, accept_downloads=True)
    page = ctx.new_page()
    try:
        s.login(page, login_cfg)
        s.goto_tracking(page)
        page.wait_for_timeout(1500)
        s._search_request(page, REQ_NO)
        if not s._open_first_detail(page):
            print("เปิดรายละเอียดไม่ได้"); raise SystemExit(1)
        page.wait_for_timeout(1500)

        # รายชื่อแท็บทั้งหมด
        dump["tabs"] = page.evaluate(
            r"""() => [...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a')]
                .map(a => ({text:(a.innerText||'').trim().slice(0,60), href:a.getAttribute('href')||''}))
                .filter(t => t.text)"""
        )

        # คลิกแท็บเอกสารตอบรับ
        clicked = page.evaluate(
            r"""() => { const re=/เอกสารตอบรับ/;
                const f=[...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a')].find(a=>re.test(a.innerText||''));
                if(f){f.click();return true;} return false; }"""
        )
        dump["tab_clicked"] = clicked
        page.wait_for_timeout(2000)

        # ทุก onclick ที่มีคำว่า Document/Confirm/Download (ดูชื่อฟังก์ชันดาวน์โหลดจริง)
        dump["onclick_handlers"] = page.evaluate(
            r"""() => {
                const out=[]; const seen=new Set();
                document.querySelectorAll('[onclick]').forEach(el=>{
                    const oc=el.getAttribute('onclick')||'';
                    if(!/document|confirm|download|pdf|file|print/i.test(oc)) return;
                    const key=oc.slice(0,80); if(seen.has(key))return; seen.add(key);
                    out.push({tag:el.tagName, text:(el.innerText||'').trim().slice(0,40), onclick:oc.slice(0,160)});
                });
                return out;
            }"""
        )

        # ทุก element ที่ข้อความมี 'บต' (ดูว่า label จริงเขียนยังไง)
        dump["bt_texts"] = page.evaluate(
            r"""() => {
                const out=[]; const seen=new Set();
                const pane=document.querySelector('#tab-response')||document;
                pane.querySelectorAll('*').forEach(el=>{
                    const t=(el.innerText||'').trim();
                    if(!/บต/.test(t)) return;
                    if(t.length>80) return;  // เอาเฉพาะ node เล็ก ๆ ที่เป็น label
                    if(seen.has(t))return; seen.add(t);
                    const links=[...el.querySelectorAll('a,button,[onclick]')].map(b=>(b.getAttribute('onclick')||'').slice(0,60)).filter(Boolean);
                    out.push({text:t, tag:el.tagName, nlinks:links.length, links:links.slice(0,4)});
                });
                return out.slice(0,40);
            }"""
        )

        # มี #tab-response ไหม + จำนวน pane
        dump["panes"] = page.evaluate(
            r"""() => ({
                has_tab_response: !!document.querySelector('#tab-response'),
                tab_panes: [...document.querySelectorAll('.tab-pane, [id^="tab-"], [role=tabpanel]')]
                    .map(p=>({id:p.id, cls:p.className.slice(0,60), shown:p.offsetParent!==null}))
            })"""
        )

        # HTML ของ pane เอกสารตอบรับ (ตัดให้พอดู)
        html = page.evaluate(
            r"""() => { const p=document.querySelector('#tab-response')||document.body; return (p.innerHTML||'').slice(0,8000); }"""
        )
        dump["tab_response_html"] = html

        # ตาราง/แถวในแท็บนี้
        dump["rows"] = page.evaluate(
            r"""() => {
                const pane=document.querySelector('#tab-response')||document;
                return [...pane.querySelectorAll('tr')].map(tr=>(tr.innerText||'').replace(/\s+/g,' ').trim().slice(0,120)).filter(Boolean).slice(0,40);
            }"""
        )

        page.screenshot(path="_probe_bt55_dom.png", full_page=True)
    finally:
        json.dump(dump, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
        print("เขียน", OUT)
        ctx.close(); browser.close()
