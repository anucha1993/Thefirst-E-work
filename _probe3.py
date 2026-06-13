"""Probe: filter by WA, open first detail, dump structure of detail page (หมายเหตุ + ข้อมูลคนต่างด้าว)."""
import os, json, io
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

load_dotenv()
U = os.environ["EWP_USERNAME"]; P = os.environ["EWP_PASSWORD"]

OUT = Path("_probe3_out.json")
result = {}

def dump():
    OUT.write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")

def login(p):
    p.goto("https://eworkpermit.doe.go.th/Login", wait_until="domcontentloaded")
    p.locator("select").nth(1).select_option(label="ผู้กระทำการแทน")
    p.wait_for_timeout(500)
    p.evaluate("""() => {
        const r = document.querySelector('input[type=radio][name=radio][value="2"]');
        if (r) { r.checked = true; r.dispatchEvent(new Event('change',{bubbles:true})); r.click(); }
    }""")
    p.get_by_role("heading", name="E-Workpermit", exact=True).first.click()
    p.wait_for_timeout(500)
    p.locator("#employer_login").fill(U)
    p.locator("#password_login").fill(P)
    p.locator("#validate_login").click()
    p.wait_for_url(lambda u: "/Login" not in u, timeout=30000)

with sync_playwright() as pw:
    b = pw.chromium.launch(headless=False)
    ctx = b.new_context(locale="th-TH", viewport={"width":1500,"height":900})
    p = ctx.new_page()
    login(p)
    p.evaluate("openpageTracking && openpageTracking()")
    p.wait_for_url("**/Permit/Tracking", timeout=20000)
    p.wait_for_timeout(2500)

    # คลิก checkbox WA (span ซ้อนทับ → ใช้ JS)
    print("[+] ติ๊ก checkbox รอยื่นเอกสารเพิ่มเติม (#WA)")
    p.evaluate("""() => {
        const cb = document.getElementById('WA');
        if (cb) { cb.checked = true; if (typeof GetDataRequestFormEtrackingForAlien === 'function') GetDataRequestFormEtrackingForAlien(); }
    }""")
    p.wait_for_timeout(3500)

    rows_total = p.evaluate("""() => {
      const t = document.querySelector('.dataTables_info');
      return t ? t.innerText : null;
    }""")
    result["pagination"] = rows_total
    dump()

    # เก็บ openDetail args ของแถวแรก
    detail_args = p.evaluate("""() => {
      const a = document.querySelector('table tbody tr a[onclick*="openDetail"]');
      if (!a) return null;
      const m = a.getAttribute('onclick').match(/openDetail\\(([^)]+)\\)/);
      return m ? m[1] : null;
    }""")
    result["detail_args_row1"] = detail_args
    dump()

    # คลิก ดูรายละเอียด แถวแรก
    p.locator('table tbody tr a[onclick*="openDetail"]').first.click()
    p.wait_for_timeout(5000)
    pages = ctx.pages
    result["page_urls"] = [pg.url for pg in pages]
    dump()

    # ใช้หน้าสุดท้ายเป็น detail
    dp = pages[-1]
    dp.bring_to_front()
    dp.wait_for_load_state("domcontentloaded")
    dp.wait_for_timeout(3000)
    dp.screenshot(path="screenshots/detail.png", full_page=True)

    info = dp.evaluate(r"""() => {
      const tabs = Array.from(document.querySelectorAll('[role=tab], .nav-link, .nav-tabs a, a[href*="#"]'))
        .map(t => ({txt: (t.innerText||'').trim().slice(0,80), href: t.getAttribute('href'), cls: t.className}))
        .filter(t => t.txt);
      const notes = Array.from(document.querySelectorAll('*')).filter(e => {
          const t = (e.innerText||'').trim();
          return t && t.includes('หมายเหตุ') && t.length < 800 && e.children.length < 8;
      }).slice(0,5).map(e => ({tag: e.tagName, cls: e.className, txt: e.innerText.replace(/\s+/g,' ').slice(0,700)}));
      return {url: location.href, title: document.title, tabs, notes};
    }""")
    result["detail_tabs_notes"] = info
    dump()

    # คลิก tab "ข้อมูลคนต่างด้าว"
    try:
        dp.get_by_text("ข้อมูลคนต่างด้าว", exact=False).first.click()
        dp.wait_for_timeout(2500)
    except Exception as e:
        result["tab_click_err"] = str(e)
    dp.screenshot(path="screenshots/detail_alien.png", full_page=True)

    alien = dp.evaluate(r"""() => {
      // capture label-value pairs (typical: dt/dd or two-column row)
      const out = {pairs: [], headings: [], html_snippet: ''};
      document.querySelectorAll('h1,h2,h3,h4,h5').forEach(h => {
        const t = (h.innerText||'').trim();
        if (t) out.headings.push(t.slice(0,100));
      });
      // จับคู่: row -> label/value
      document.querySelectorAll('.row').forEach(r => {
        const cols = r.querySelectorAll(':scope > div');
        if (cols.length >= 2) {
          const k = (cols[0].innerText||'').trim();
          const v = (cols[cols.length-1].innerText||'').trim();
          if (k && v && k !== v && k.length < 80 && v.length < 300) out.pairs.push([k, v]);
        }
      });
      // dt/dd
      document.querySelectorAll('dt').forEach(dt => {
        const dd = dt.nextElementSibling;
        if (dd && dd.tagName === 'DD') out.pairs.push([dt.innerText.trim(), dd.innerText.trim()]);
      });
      out.html_snippet = document.body.innerText.replace(/\s{2,}/g,'\n').slice(0, 3000);
      return out;
    }""")
    result["alien_tab_data"] = alien
    dump()
    ctx.close(); b.close()
