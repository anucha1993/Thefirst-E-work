"""Probe: นำทางเมนู บต.44 (คำขอเปลี่ยนรายการในใบอนุญาตทำงาน) + dump ฟอร์มค้นหาคนต่างด้าว
รัน: .\.venv\Scripts\python.exe _probe_bt44.py
"""
import json
from playwright.sync_api import sync_playwright
from scrape_wa import load_config, login

cfg = load_config(require_login=True)

with sync_playwright() as pw:
    b = pw.chromium.launch(headless=False, args=["--ignore-certificate-errors", "--start-maximized"])
    ctx = b.new_context(locale="th-TH", ignore_https_errors=True,
                        viewport={"width": 1920, "height": 1080})
    page = ctx.new_page()
    login(page, cfg)
    page.wait_for_timeout(1500)

    page.goto("https://eworkpermit.doe.go.th/", wait_until="domcontentloaded", timeout=30_000)
    page.wait_for_timeout(2000)

    # 1) คลิก 'เมนูบริการ'
    clicked = page.evaluate(r"""() => {
        const a = document.querySelector('a.lang_menu_service')
          || Array.from(document.querySelectorAll('a,button'))
               .find(x => /เมนูบริการ/.test((x.textContent || '').trim()));
        if (a) { a.click(); return true; }
        return false;
    }""")
    print("คลิกเมนูบริการ:", clicked)
    page.wait_for_timeout(1200)

    # 2) dump รายการ category (openCity tabs) ทั้งหมด + ปุ่มที่มี onclick=openCity(...)
    cats = page.evaluate(r"""() => {
        const out = [];
        document.querySelectorAll('[onclick*="openCity"]').forEach(e => {
            out.push({tag: e.tagName, id: e.id, txt: (e.textContent||'').trim().slice(0,80),
                      onclick: e.getAttribute('onclick')});
        });
        return out;
    }""")
    print("\n=== CATEGORIES (openCity) ===")
    print(json.dumps(cats, ensure_ascii=False, indent=2))

    # 3) เปิดหมวด tab_CHANGE_REQ แล้วคลิกเมนู #CHANGE_EMPLOYER (ตาม id ตรง ๆ — ไม่ใช้ text)
    page.evaluate(r"""() => {
        try { if (typeof openCity === 'function') openCity('tab_CHANGE_REQ', new Event('click')); } catch (e) {}
    }""")
    page.wait_for_timeout(900)
    ok = page.evaluate(r"""() => {
        const t = document.querySelector('#CHANGE_EMPLOYER');
        if (t) { t.click(); return true; }
        return false;
    }""")
    print("\nคลิก #CHANGE_EMPLOYER:", ok)
    try:
        page.wait_for_url("**/WorkPermit**", timeout=30_000)
    except Exception as e:
        print("wait WorkPermit:", e)
    page.wait_for_timeout(3000)
    print("URL หลังคลิก:", page.url)
    page.screenshot(path="screenshots/bt44_form.png", full_page=True)

    # หัวข้อ/แบบฟอร์มที่แสดง (ยืนยันว่าเป็น บต.44 / เปลี่ยนรายการ)
    header = page.evaluate(r"""() => {
        const h = Array.from(document.querySelectorAll('h1,h2,h3,h4,h5,.card-title,.page-title,.title'))
          .map(e => (e.textContent||'').trim()).filter(t => t && t.length < 160).slice(0, 12);
        return h;
    }""")
    print("\n=== FORM HEADERS ===")
    print(json.dumps(header, ensure_ascii=False, indent=2))

    # 4) dump ปุ่มเปิด modal ค้นหาคนต่างด้าว
    form = page.evaluate(r"""() => {
        const searchBtn = Array.from(document.querySelectorAll('button,a'))
          .filter(e => /search_alien_modal|ค้นหาข้อมูลคนต่างด้าว/.test((e.getAttribute('onclick')||'')+(e.textContent||'')))
          .map(e => ({tag:e.tagName, id:e.id, txt:(e.textContent||'').trim().slice(0,60), onclick:(e.getAttribute('onclick')||'').slice(0,120)}));
        return {searchBtns: searchBtn};
    }""")
    print("\n=== SEARCH-ALIEN BUTTON ===")
    print(json.dumps(form, ensure_ascii=False, indent=2))

    # 5) เปิด modal ค้นหา แล้ว dump ฟิลด์ทั้งหมด (เทียบกับ บต.30)
    page.evaluate(r"""() => {
        const b = Array.from(document.querySelectorAll('button,a'))
          .find(e => /search_alien_modal\.show/.test(e.getAttribute('onclick') || ''));
        if (b) b.click();
    }""")
    page.wait_for_timeout(1500)
    fields = page.evaluate(r"""() => {
        const inModal = (e) => e.offsetParent !== null;
        const inputs = Array.from(document.querySelectorAll('input:not([type=hidden])'))
          .filter(inModal).map(e => ({id:e.id, name:e.name, type:e.type, ph:e.placeholder,
              label:(e.closest('.form-group,.mb-3,.row')?.querySelector('label')?.textContent||'').trim().slice(0,60)}));
        const selects = Array.from(document.querySelectorAll('select')).filter(inModal)
          .map(e => ({id:e.id, name:e.name, opts:Array.from(e.options).slice(0,6).map(o=>o.text.trim().slice(0,30)),
              label:(e.closest('.form-group,.mb-3,.row')?.querySelector('label')?.textContent||'').trim().slice(0,60)}));
        const btns = Array.from(document.querySelectorAll('button')).filter(inModal)
          .map(e => ({id:e.id, txt:(e.textContent||'').trim().slice(0,40), onclick:(e.getAttribute('onclick')||'').slice(0,80)}));
        return {inputs, selects, btns};
    }""")
    print("\n=== SEARCH-ALIEN MODAL FIELDS ===")
    print(json.dumps(fields, ensure_ascii=False, indent=2))
    page.screenshot(path="screenshots/bt44_search_modal.png", full_page=True)

    print("\n--- เปิดค้างไว้ให้ตรวจสอบ ---")
    input("กด Enter เพื่อปิด...")
    ctx.close(); b.close()
