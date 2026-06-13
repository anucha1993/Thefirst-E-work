"""Probe:
1) Tracking page: หา select 'รายการคำขอ' (id/name/options)
2) Detail page: tab 'คำขออนุญาต' (tab anchor + content)
3) ดำเนินการแก้ไข button -> หน้า edit -> บันทึกเพิ่มเติม
"""
import os, json
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

load_dotenv()
U = os.environ["EWP_USERNAME"]; P = os.environ["EWP_PASSWORD"]

def login(p):
    p.goto("https://eworkpermit.doe.go.th/Login", wait_until="domcontentloaded")
    p.locator("select").nth(1).select_option(label="ผู้กระทำการแทน")
    p.wait_for_timeout(500)
    p.evaluate("""() => { const r=document.querySelector('input[type=radio][name=radio][value="2"]'); if(r){r.checked=true;r.dispatchEvent(new Event('change',{bubbles:true}));r.click();}}""")
    p.get_by_role("heading", name="E-Workpermit", exact=True).first.click()
    p.wait_for_timeout(500)
    p.locator("#employer_login").fill(U); p.locator("#password_login").fill(P); p.locator("#validate_login").click()
    p.wait_for_url(lambda u: "/Login" not in u, timeout=30000)

with sync_playwright() as pw:
    b = pw.chromium.launch(headless=False)
    ctx = b.new_context(locale="th-TH", viewport={"width":1500,"height":900})
    p = ctx.new_page()
    login(p)

    # ----- 1) Tracking page -----
    p.evaluate("openpageTracking && openpageTracking()")
    p.wait_for_url("**/Permit/Tracking", timeout=20000)
    p.wait_for_timeout(3000)

    # WA filter
    p.evaluate("""() => { const cb=document.getElementById('WA'); if(cb){cb.checked=true; if(typeof GetDataRequestFormEtrackingForAlien==='function') GetDataRequestFormEtrackingForAlien(); else cb.click();}}""")
    p.wait_for_timeout(2500)

    # หา select ทั้งหมดในหน้านี้
    sels = p.evaluate(r"""() => {
       const arr=[];
       document.querySelectorAll('select').forEach(s => {
          arr.push({
             id: s.id, name: s.name, cls: s.className,
             options: Array.from(s.options).slice(0,30).map(o => ({val: o.value, text: o.text.trim().slice(0,200)})),
             labelText: (() => {
                const l = s.closest('div,form,fieldset');
                return l ? (l.innerText||'').replace(/\s+/g,' ').slice(0,200) : '';
             })(),
          });
       });
       return arr;
    }""")
    Path("_probe_tracking_selects.json").write_text(json.dumps(sels, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"[+] selects on tracking: {len(sels)}")

    # ----- 2) Detail page tabs -----
    URL = "https://eworkpermit.doe.go.th/RequestForm41/DetailRequest41?user_id=303816&status=WA&group_id=69113100022770&form_type=MT_41_4_59"
    p.goto(URL, wait_until="domcontentloaded")
    p.wait_for_timeout(4000)
    tabs = p.evaluate(r"""() => {
       const out=[];
       document.querySelectorAll('a[data-toggle="tab"], a[role="tab"], .nav-link').forEach(a => {
          out.push({text: (a.innerText||'').trim().slice(0,80), href: a.getAttribute('href')||'', onclick: a.getAttribute('onclick')||''});
       });
       return out;
    }""")
    Path("_probe_detail_tabs.json").write_text(json.dumps(tabs, indent=2, ensure_ascii=False), encoding="utf-8")

    # คลิก tab "คำขออนุญาต" (น่าจะ tab_default_2_3 หรือ similar)
    # ลองหาทุก tab ที่มีคำว่า "คำขออนุญาต"
    target = p.evaluate(r"""() => {
       const a = Array.from(document.querySelectorAll('a')).find(x => /คำขออนุญาต/.test(x.innerText||''));
       if (!a) return null;
       a.click();
       return {text: a.innerText.trim(), href: a.getAttribute('href')};
    }""")
    print("[+] tab clicked:", target)
    p.wait_for_timeout(2500)
    # capture content of that tab (active pane)
    pane = p.evaluate(r"""() => {
       const panes = document.querySelectorAll('.tab-pane.active, .tab-content .active');
       const arr=[];
       panes.forEach(pn => {
          arr.push({
             id: pn.id, cls: pn.className,
             textPreview: (pn.innerText||'').replace(/\s+/g,'\n').slice(0,4000),
          });
       });
       return arr;
    }""")
    Path("_probe_kham_kor_pane.json").write_text(json.dumps(pane, indent=2, ensure_ascii=False), encoding="utf-8")

    # ----- 3) ปุ่ม ดำเนินการแก้ไข -----
    fix_btn = p.evaluate(r"""() => {
       const btns = Array.from(document.querySelectorAll('button, a, input[type=button], input[type=submit]'))
          .filter(b => /ดำเนินการแก้ไข/.test(b.innerText||b.value||''));
       return btns.map(b => ({tag: b.tagName, text: (b.innerText||b.value||'').trim(), onclick: b.getAttribute('onclick')||'', href: b.getAttribute('href')||''}));
    }""")
    Path("_probe_fix_btn.json").write_text(json.dumps(fix_btn, indent=2, ensure_ascii=False), encoding="utf-8")
    print("[+] fix button(s):", fix_btn)

    # คลิกปุ่ม ดำเนินการแก้ไข (เปิด popup ใหม่ก็ดักไว้)
    new_pages_before = len(ctx.pages)
    p.evaluate(r"""() => {
       const b = Array.from(document.querySelectorAll('button, a')).find(x => /ดำเนินการแก้ไข/.test(x.innerText||''));
       if (b) b.click();
    }""")
    p.wait_for_timeout(4500)

    # ตรวจหน้าใหม่ที่อาจเปิดขึ้น
    pages = ctx.pages
    print(f"[+] pages now: {len(pages)} (was {new_pages_before})")
    target_page = pages[-1]  # ใช้หน้าใหม่สุด (หรือเดิมถ้าไม่มี popup)
    try:
        target_page.wait_for_load_state("domcontentloaded", timeout=10000)
    except Exception: pass
    target_page.wait_for_timeout(3000)
    edit_info = target_page.evaluate(r"""() => {
       const out = {url: location.href, title: document.title};
       // หา "บันทึกเพิ่มเติม" และ list ข้างใต้
       const allText = document.body.innerText.replace(/\s+/g,' ').slice(0,5000);
       out.bodyPreview = allText;
       // เก็บ list items ทั้งหมด
       out.lists = [];
       document.querySelectorAll('ul, ol').forEach(ul => {
          const items = Array.from(ul.querySelectorAll('li')).map(li => (li.innerText||'').trim()).filter(Boolean);
          if (items.length) out.lists.push(items);
       });
       // หา block ที่มีคำว่า บันทึกเพิ่มเติม
       const block = Array.from(document.querySelectorAll('div,section,form'))
          .find(d => /บันทึกเพิ่มเติม/.test(d.innerText||'') && (d.innerText||'').length < 3000);
       out.notesBlock = block ? block.innerText : '';
       return out;
    }""")
    Path("_probe_edit_page.json").write_text(json.dumps(edit_info, indent=2, ensure_ascii=False), encoding="utf-8")
    print("[+] edit page url:", edit_info.get("url"))

    ctx.close(); b.close()
print("done")
