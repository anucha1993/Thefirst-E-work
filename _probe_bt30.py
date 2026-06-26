"""Probe (read-only) — ค้นหาโครงสร้าง DOM ของโหมด บต.30
flow ที่ต้องการ:
  เมนูบริการ → การยื่นขอต่ออายุใบอนุญาตทำงาน
            → การยื่นคำขอต่ออายุใบอนุญาตทำงานของคนต่างด้าวที่ได้รับอนุญาตทำงานตาม MoU (MT_59_MOU_RENEWAL)
            → ฟอร์ม "ค้นหาข้อมูลคนต่างด้าว"

โปรแกรมนี้ "ไม่กดบันทึก / ไม่ส่งคำขอ" — แค่ login + เปิดเมนู + dump โครงสร้างฟอร์มลง JSON
ใช้บัญชีจาก UsernameLogin.xlsx (บัญชีแรก)

รัน: .\.venv\Scripts\python.exe _probe_bt30.py
"""
from __future__ import annotations

import json
from pathlib import Path

from playwright.sync_api import sync_playwright

from scrape_wa import login, _read_login_accounts, ROOT

OUT = ROOT / "_probe_bt30_out.json"
SHOTS = ROOT / "screenshots"
SHOTS.mkdir(exist_ok=True)

DUMP_JS = r"""() => {
  const txt = e => ((e.textContent || '').replace(/\s+/g, ' ').trim()).slice(0, 90);
  const vis = e => !!(e && e.offsetParent !== null);
  const collectInputs = () => Array.from(document.querySelectorAll(
      'input:not([type=hidden])')).map(e => ({
        id: e.id, name: e.name, type: e.type, placeholder: e.placeholder,
        visible: vis(e),
        label: (e.labels && e.labels[0] ? e.labels[0].textContent.trim()
               : (e.previousElementSibling ? e.previousElementSibling.textContent.trim() : '')).slice(0, 60),
      }));
  const collectSelects = () => Array.from(document.querySelectorAll('select')).map(e => ({
        id: e.id, name: e.name, visible: vis(e),
        opts: Array.from(e.options).slice(0, 60).map(o => ({ v: o.value, t: (o.textContent || '').trim().slice(0, 50) })),
      }));
  const collectButtons = () => Array.from(document.querySelectorAll('button, a.btn, a[onclick], input[type=button], input[type=submit]')).map(e => ({
        tag: e.tagName, id: e.id, cls: (e.className || '').toString().slice(0, 80),
        txt: txt(e), onclick: (e.getAttribute('onclick') || '').slice(0, 120), visible: vis(e),
      }));
  return {
    url: location.href,
    title: document.title,
    inputs: collectInputs(),
    selects: collectSelects(),
    buttons: collectButtons(),
    headings: Array.from(document.querySelectorAll('h1,h2,h3,h4,h5,legend')).map(e => txt(e)).filter(Boolean).slice(0, 40),
  };
}"""

# ดึงรายการเมนูทั้งหมดในเมนูบริการ (id ขึ้นต้น MT_ + ปุ่มหมวด openCity)
MENU_JS = r"""() => {
  const out = { categories: [], items: [] };
  // ปุ่มหมวด (tab) — มักมี onclick="openCity('tab_XXX', ...)"
  document.querySelectorAll('[onclick*="openCity"]').forEach(e => {
    out.categories.push({
      tag: e.tagName, id: e.id, txt: (e.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 80),
      onclick: (e.getAttribute('onclick') || '').slice(0, 120),
    });
  });
  // รายการคำขอ — element id ขึ้นต้น MT_ หรือ CHANGE_/REPLACE_/NAMELIST_
  document.querySelectorAll('[id^="MT_"], [id^="CHANGE_"], [id^="REPLACE_"], [id^="NAMELIST_"]').forEach(e => {
    out.items.push({
      id: e.id, txt: (e.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 100),
      onclick: (e.getAttribute('onclick') || '').slice(0, 160),
      visible: !!(e.offsetParent !== null),
    });
  });
  return out;
}"""


def main() -> None:
    accounts = _read_login_accounts(ROOT / "UsernameLogin.xlsx")
    if not accounts:
        raise SystemExit("ไม่พบบัญชีใน UsernameLogin.xlsx")
    acct = next(iter(accounts.values()))
    cfg = {
        "username": acct["username"],
        "password": acct["password"],
        "user_type": acct["type"],
        "method": acct.get("method") or "E-Workpermit",
    }
    print(f"[i] ใช้บัญชี: {cfg['username']} ({cfg['user_type']}, {cfg['method']})")

    report: dict = {"account": cfg["username"]}
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False, args=["--ignore-certificate-errors", "--start-maximized"])
        ctx = browser.new_context(locale="th-TH", ignore_https_errors=True, viewport={"width": 1600, "height": 950})
        page = ctx.new_page()
        try:
            login(page, cfg)
            page.wait_for_timeout(2000)

            # 1) เปิดหน้าหลัก + คลิก 'เมนูบริการ'
            page.goto("https://eworkpermit.doe.go.th/", wait_until="domcontentloaded", timeout=30_000)
            page.wait_for_timeout(2500)
            page.evaluate(r"""() => {
                const a = document.querySelector('a.lang_menu_service')
                  || Array.from(document.querySelectorAll('a,button')).find(x => /เมนูบริการ/.test((x.textContent||'').trim()));
                if (a) a.click();
            }""")
            page.wait_for_timeout(2000)
            page.screenshot(path=str(SHOTS / "bt30_01_menu.png"), full_page=True)

            menu = page.evaluate(MENU_JS)
            report["menu"] = menu
            print(f"[i] หมวด (categories): {len(menu['categories'])}")
            for c in menu["categories"]:
                print("   CAT:", c["id"], "|", c["txt"], "|", c["onclick"])
            print(f"[i] รายการ (items): {len(menu['items'])}")
            for it in menu["items"]:
                mark = "*" if "MT_59_MOU_RENEWAL" in it["id"] else " "
                print(f"  {mark} ITEM:", it["id"], "vis=", it["visible"], "|", it["txt"][:60])

            # 2) เปิดหมวด renewal + คลิก MT_59_MOU_RENEWAL
            #    ลองเปิดทุก openCity ที่ข้อความมีคำว่า 'ต่ออายุ' ก่อน แล้วค่อยคลิก item
            page.evaluate(r"""() => {
                const cat = Array.from(document.querySelectorAll('[onclick*="openCity"]'))
                  .find(e => /ต่ออายุ/.test(e.textContent||''));
                if (cat) cat.click();
            }""")
            page.wait_for_timeout(1500)
            page.screenshot(path=str(SHOTS / "bt30_02_renew_cat.png"), full_page=True)

            # capture onclick ของ MT_59_MOU_RENEWAL (ก่อนคลิก) เพื่อรู้ว่ามันเรียกฟังก์ชันอะไร
            target_info = page.evaluate(r"""() => {
                const t = document.querySelector('#MT_59_MOU_RENEWAL');
                if (!t) return null;
                return { id: t.id, txt: (t.textContent||'').trim().slice(0,120),
                         onclick: t.getAttribute('onclick')||'', href: t.href||'', visible: !!(t.offsetParent!==null) };
            }""")
            report["target_item"] = target_info
            print("[i] MT_59_MOU_RENEWAL =", json.dumps(target_info, ensure_ascii=False))

            # คลิกเข้า form
            page.evaluate(r"""() => {
                const t = document.querySelector('#MT_59_MOU_RENEWAL');
                if (t) t.click();
            }""")
            page.wait_for_timeout(4000)
            print("[i] URL หลังคลิก:", page.url)
            page.screenshot(path=str(SHOTS / "bt30_03_form.png"), full_page=True)

            # 3) dump ฟอร์มหน้าแรก (ค้นหาข้อมูลคนต่างด้าว)
            report["form_page"] = page.evaluate(DUMP_JS)
            report["form_url"] = page.url

            # ถ้ามี modal/section 'ค้นหาข้อมูลคนต่างด้าว' ลองหา + เปิด แล้ว dump อีกรอบ
            opened = page.evaluate(r"""() => {
                const btn = Array.from(document.querySelectorAll('button,a,div'))
                  .find(e => /ค้นหาข้อมูลคนต่างด้าว/.test((e.textContent||'').trim()) && e.offsetParent!==null);
                if (btn) { btn.click(); return (btn.textContent||'').trim().slice(0,60); }
                return '';
            }""")
            if opened:
                print("[i] คลิกเปิด:", opened)
                page.wait_for_timeout(2000)
                page.screenshot(path=str(SHOTS / "bt30_04_search_modal.png"), full_page=True)
                report["after_search_click"] = page.evaluate(DUMP_JS)

            OUT.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"\n[OK] เขียนผลลง {OUT.name} แล้ว — ดู screenshots/bt30_*.png ประกอบ")
            print("[i] (โปรแกรมไม่กดบันทึก/ไม่ส่งคำขอใดๆ)")
            input("กด Enter เพื่อปิดเบราว์เซอร์...")
        finally:
            ctx.close(); browser.close()


if __name__ == "__main__":
    main()
