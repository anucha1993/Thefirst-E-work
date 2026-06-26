# -*- coding: utf-8 -*-
"""เจาะโครงสร้าง #tab_default_4: จับคู่ลิงก์ GetDocumentConfirm กับ label เอกสาร (บต.55/บต.56)"""
import json
from playwright.sync_api import sync_playwright
import scrape_wa as s

REQ_NO = "69155200181078"
OUT = "_probe_bt55_dom2_out.json"

cfg = s.load_config(require_login=False)
acct = list(s._read_login_accounts(s.ROOT / "UsernameLogin.xlsx").values())[0]
login_cfg = {"username": acct["username"], "password": acct["password"],
             "user_type": acct["type"], "method": acct.get("method") or "E-Workpermit"}

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
        page.evaluate(
            r"""() => { const re=/เอกสารตอบรับ/;
                const f=[...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a')].find(a=>re.test(a.innerText||''));
                if(f) f.click(); }"""
        )
        page.wait_for_timeout(2000)

        # สำหรับแต่ละลิงก์ GetDocumentConfirm: ปีนหา ancestor ที่ innerText มี 'บต' แล้วเก็บ label
        dump["links"] = page.evaluate(
            r"""() => {
                const out=[];
                document.querySelectorAll('a[onclick*="GetDocumentConfirm"], button[onclick*="GetDocumentConfirm"], [onclick*="GetDocumentConfirm"]').forEach((a,idx)=>{
                    const oc=a.getAttribute('onclick')||'';
                    let label='', anc='', lvl=-1, node=a;
                    for(let i=0;i<8 && node;i++){
                        const t=(node.innerText||'').replace(/\s+/g,' ').trim();
                        if(/บต/.test(t)){ label=t.slice(0,120); anc=node.tagName+'.'+(node.className||'').slice(0,40); lvl=i; break; }
                        node=node.parentElement;
                    }
                    out.push({idx, onclick:oc.slice(0,90), label, ancestor:anc, climbed:lvl});
                });
                return out;
            }"""
        )

        # โครงสร้าง pane จริง #tab_default_4 (innerHTML แบบย่อ ตัด script/style)
        dump["pane_html"] = page.evaluate(
            r"""() => {
                const p=document.querySelector('#tab_default_4'); if(!p) return '(no #tab_default_4)';
                const c=p.cloneNode(true);
                c.querySelectorAll('script,style,link,meta').forEach(e=>e.remove());
                return (c.innerHTML||'').replace(/\s+/g,' ').slice(0,12000);
            }"""
        )

        # นับด้วยตรรกะปัจจุบัน (pane #tab-response || document, el มี GetDocumentConfirm 1 อัน + text บต.55)
        dump["current_count"] = page.evaluate(
            r"""(pat) => {
                const re=new RegExp(pat);
                const pane=document.querySelector('#tab-response')||document;
                const seen=new Set(); let n=0;
                pane.querySelectorAll('*').forEach(el=>{
                    const txt=(el.innerText||'').trim();
                    if(!txt||!re.test(txt)) return;
                    const links=[...el.querySelectorAll('a,button,[onclick]')].filter(b=>(b.getAttribute('onclick')||'').includes('GetDocumentConfirm'));
                    if(links.length!==1) return;
                    const oc=links[0].getAttribute('onclick')||''; if(!oc||seen.has(oc))return; seen.add(oc); n++;
                });
                return n;
            }""",
            r"บต\.?\s*55",
        )
    finally:
        json.dump(dump, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
        print("เขียน", OUT, "| current_count =", dump.get("current_count"))
        ctx.close(); browser.close()
