"""Scope-probe the ข้อมูลคนต่างด้าว tab pane only."""
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
    # ไปยัง detail page โดยตรง
    URL = "https://eworkpermit.doe.go.th/RequestForm41/DetailRequest41?user_id=303816&status=WA&group_id=69113100022770&form_type=MT_41_4_59"
    p.goto(URL, wait_until="domcontentloaded")
    p.wait_for_timeout(5000)
    # คลิก tab ข้อมูลคนต่างด้าว
    p.evaluate("""() => {
       const a = Array.from(document.querySelectorAll('a[href="#tab_default_2_1"]'))[0];
       if (a) a.click();
    }""")
    p.wait_for_timeout(2000)
    data = p.evaluate(r"""() => {
      const pane = document.querySelector('#tab_default_2_1');
      if (!pane) return {err: 'pane not found'};
      const pairs = [];
      pane.querySelectorAll('.row').forEach(r => {
         const cols = r.querySelectorAll(':scope > div');
         if (cols.length >= 2) {
           const k = (cols[0].innerText||'').trim();
           const v = (cols[cols.length-1].innerText||'').trim();
           if (k && v && k!==v && k.length<120 && v.length<400) pairs.push([k,v]);
         }
      });
      // also try labels and following spans
      const labels = [];
      pane.querySelectorAll('label, .form-label, dt').forEach(el => {
         labels.push((el.innerText||'').trim().slice(0,80));
      });
      return {
        text: pane.innerText.replace(/\s{2,}/g,'\n').slice(0,6000),
        pairs,
        labels,
        rowCount: pane.querySelectorAll('.row').length,
      };
    }""")
    Path("_probe_alien_tab.json").write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    # หมายเหตุ
    note = p.evaluate(r"""() => {
       const notes = [];
       document.querySelectorAll('span.lang_text_reason, .div_info, .note-card, .alert').forEach(n => {
          const t = (n.innerText||'').trim();
          if (t) notes.push({cls: n.className, txt: t.replace(/\s+/g,' ').slice(0,600)});
       });
       return notes;
    }""")
    Path("_probe_note.json").write_text(json.dumps(note, indent=2, ensure_ascii=False), encoding="utf-8")
    ctx.close(); b.close()
print("done")
