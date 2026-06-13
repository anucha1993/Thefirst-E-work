"""Probe: หน้า 'จัดการบัญชี' → tab 'ข้อมูลคนต่างด้าว'
รัน 1 ครั้งเพื่อ dump URL/selector ที่ใช้ในหน้านี้ จะนำผลไปใช้เขียน scraper ต่อ
"""
import json, os, sys
from pathlib import Path
from dotenv import load_dotenv

# ใช้ login() เดิมจาก scrape_wa
sys.path.insert(0, str(Path(__file__).parent))
from scrape_wa import login, LOGIN_URL
from playwright.sync_api import sync_playwright

load_dotenv()
cfg = {
    "username": os.environ["EWP_USERNAME"],
    "password": os.environ["EWP_PASSWORD"],
    "user_type": os.environ.get("EWP_USER_TYPE", "employer"),
    "method": os.environ.get("EWP_LOGIN_METHOD", "username"),
}
OUT = Path(__file__).parent / "screenshots"
OUT.mkdir(exist_ok=True)

with sync_playwright() as pw:
    b = pw.chromium.launch(headless=False)
    ctx = b.new_context(locale="th-TH", ignore_https_errors=True)
    p = ctx.new_page()
    print("[1] login ...")
    login(p, cfg)
    print(f"    → after login: {p.url}")
    p.wait_for_timeout(2000)
    p.screenshot(path=str(OUT / "acc_01_after_login.png"), full_page=True)

    # ขั้นที่ 1: หาปุ่ม "จัดการบัญชี" หรือ navigate ไปหน้า profile
    print("[2] หา/คลิก 'จัดการบัญชี'...")
    info = p.evaluate(
        r"""() => {
          const all = Array.from(document.querySelectorAll('a, button, [onclick]'));
          const found = [];
          for (const el of all) {
            const t = (el.innerText || '').replace(/\s+/g,' ').trim();
            if (/จัดการบัญชี|ข้อมูลส่วนตัว|ข้อมูลคนต่างด้าว/.test(t)) {
              found.push({
                tag: el.tagName,
                text: t.slice(0,80),
                onclick: el.getAttribute('onclick') || '',
                href: el.getAttribute('href') || '',
                id: el.id,
                cls: el.className,
              });
            }
          }
          return {url: location.href, candidates: found.slice(0, 30)};
        }"""
    )
    print(json.dumps(info, indent=2, ensure_ascii=False))

    # ลองคลิก "จัดการบัญชี"
    clicked = p.evaluate(
        r"""() => {
          const all = Array.from(document.querySelectorAll('a, button, [onclick], h2, h3, h4, .card, .card-body'));
          for (const el of all) {
            const t = (el.innerText || '').replace(/\s+/g,' ').trim();
            if (/^จัดการบัญชี/.test(t) && t.length < 80) { el.click(); return t; }
          }
          return null;
        }"""
    )
    print(f"    clicked: {clicked}")
    p.wait_for_timeout(3500)
    print(f"    → after click: {p.url}")
    p.screenshot(path=str(OUT / "acc_02_after_click_account.png"), full_page=True)

    # ขั้นที่ 2: หา tab "ข้อมูลคนต่างด้าว"
    print("[3] หา tab 'ข้อมูลคนต่างด้าว'...")
    tabs = p.evaluate(
        r"""() => {
          const all = Array.from(document.querySelectorAll('a, button, [role="tab"], .nav-link, .nav-item'));
          return all.map(el => ({
            tag: el.tagName,
            text: (el.innerText||'').replace(/\s+/g,' ').trim().slice(0,60),
            href: el.getAttribute('href') || '',
            onclick: (el.getAttribute('onclick')||'').slice(0,120),
            id: el.id,
            cls: el.className,
          })).filter(x => x.text && x.text.length < 60);
        }"""
    )
    print(json.dumps(tabs, indent=2, ensure_ascii=False)[:5000])

    # คลิก tab ข้อมูลคนต่างด้าว
    p.evaluate(
        r"""() => {
          const all = Array.from(document.querySelectorAll('a, button, [role="tab"], .nav-link'));
          for (const el of all) {
            const t = (el.innerText || '').replace(/\s+/g,' ').trim();
            if (/^ข้อมูลคนต่างด้าว/.test(t) && t.length < 40) { el.click(); return; }
          }
        }"""
    )
    p.wait_for_timeout(3500)
    print(f"    → URL: {p.url}")
    p.screenshot(path=str(OUT / "acc_03_alien_tab.png"), full_page=True)

    # ขั้นที่ 3: ดู sub-tab + table + pagination
    print("[4] dump table/pagination/subtabs...")
    info2 = p.evaluate(
        r"""() => {
          const collect = (sel, n=20) => Array.from(document.querySelectorAll(sel))
            .map(e => ({text:(e.innerText||'').replace(/\s+/g,' ').trim().slice(0,80), id:e.id, cls:e.className}))
            .filter(x => x.text).slice(0,n);
          const tables = Array.from(document.querySelectorAll('table')).map(t => ({
            id: t.id, cls: t.className,
            headers: Array.from(t.querySelectorAll('thead th, thead td')).map(th => (th.innerText||'').trim()),
            firstRowCells: Array.from((t.querySelector('tbody tr')||{querySelectorAll:()=>[]}).querySelectorAll('td')).map(td => (td.innerText||'').replace(/\s+/g,' ').trim().slice(0,80)),
            rowCount: t.querySelectorAll('tbody tr').length,
          }));
          return {
            url: location.href,
            subtabs: collect('.nav-tabs a, .nav-pills a, [role="tab"], button[onclick*="tab"]'),
            paginate: collect('a,button', 60).filter(x => /ถัดไป|ย้อนกลับ|Next|Prev|>|</.test(x.text)),
            tables,
          };
        }"""
    )
    print(json.dumps(info2, indent=2, ensure_ascii=False)[:8000])

    input("\nPress Enter to close...")
    ctx.close(); b.close()
print("DONE — screenshots saved at screenshots/acc_*.png")
