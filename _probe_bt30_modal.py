"""Probe v2 (read-only) — เปิด modal 'ค้นหาข้อมูลคนต่างด้าว' ของฟอร์ม บต.30
แล้ว dump โครงสร้าง field ในกล่อง modal (รวม attribute ของช่องวันเกิด)
ไม่กดบันทึก/ไม่ส่งคำขอ

รัน: .venv\\Scripts\\python.exe _probe_bt30_modal.py
"""
from __future__ import annotations

import json
from pathlib import Path

from playwright.sync_api import sync_playwright

from scrape_wa import login, _read_login_accounts, ROOT

OUT = ROOT / "_probe_bt30_modal_out.json"
SHOTS = ROOT / "screenshots"
SHOTS.mkdir(exist_ok=True)

FORM_URL = "https://eworkpermit.doe.go.th/WorkPermit?user_type=alien&ft=RENEW_REQ&ut=alien&uti=3"

# dump ทุก field ที่อยู่ใน modal ที่กำลังเปิด (visible) — เก็บ attribute สำคัญ
MODAL_JS = r"""() => {
  const visible = el => !!(el && el.offsetParent !== null && el.getClientRects().length);
  // หา modal container ที่เปิดอยู่ (มีคำว่า ค้นหาข้อมูลคนต่างด้าว)
  let modal = null;
  document.querySelectorAll('div').forEach(d => {
    if (modal) return;
    const t = d.textContent || '';
    if (/ค้นหาข้อมูลคนต่างด้าว/.test(t) && visible(d) && d.querySelector('input,select')) {
      // เลือกตัวที่เล็กที่สุดที่ยังมี field (ใกล้เคียง modal box)
      modal = d;
    }
  });
  const scope = modal || document.body;
  const fld = (e) => ({
    tag: e.tagName, id: e.id, name: e.name, type: e.type,
    readonly: e.readOnly || e.hasAttribute('readonly'),
    disabled: e.disabled, cls: (e.className||'').toString().slice(0,90),
    maxlength: e.getAttribute('maxlength'), placeholder: e.placeholder,
    value: (e.value||'').slice(0,40), visible: visible(e),
    label: (() => {
      if (e.labels && e.labels[0]) return e.labels[0].textContent.replace(/\s+/g,' ').trim().slice(0,50);
      let p = e.closest('.form-group, .col, .row, td, div');
      let lab = p ? p.querySelector('label') : null;
      return lab ? lab.textContent.replace(/\s+/g,' ').trim().slice(0,50) : '';
    })(),
    opts: e.tagName === 'SELECT' ? Array.from(e.options).slice(0,8).map(o => o.textContent.trim().slice(0,24)) : undefined,
  });
  const inputs = Array.from(scope.querySelectorAll('input:not([type=hidden]), select, textarea')).map(fld);
  const buttons = Array.from(scope.querySelectorAll('button, a.btn, input[type=button], input[type=submit]'))
    .filter(visible).map(b => ({ id: b.id, txt: (b.textContent||b.value||'').replace(/\s+/g,' ').trim().slice(0,30), onclick: (b.getAttribute('onclick')||'').slice(0,80) }));
  // ช่องวันเกิด: เก็บ raw outerHTML สั้นๆ เพื่อดู datepicker
  const bd = document.querySelector('#birthDateCheck');
  return {
    modalFound: !!modal,
    inputs, buttons,
    birthDateCheckHTML: bd ? bd.outerHTML.slice(0, 400) : '(not found)',
    birthDateParent: bd && bd.parentElement ? bd.parentElement.outerHTML.slice(0, 600) : '',
  };
}"""


def main() -> None:
    accounts = _read_login_accounts(ROOT / "UsernameLogin.xlsx")
    acct = next(iter(accounts.values()))
    cfg = {"username": acct["username"], "password": acct["password"],
           "user_type": acct["type"], "method": acct.get("method") or "E-Workpermit"}
    print(f"[i] บัญชี: {cfg['username']}")
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False, args=["--ignore-certificate-errors", "--start-maximized"])
        ctx = browser.new_context(locale="th-TH", ignore_https_errors=True, viewport={"width": 1600, "height": 950})
        page = ctx.new_page()
        try:
            login(page, cfg)
            page.wait_for_timeout(1500)
            # นำทางผ่านเมนู (ให้ setFormTypeRenew ตั้งค่า session ก่อน)
            page.goto("https://eworkpermit.doe.go.th/", wait_until="domcontentloaded", timeout=30_000)
            page.wait_for_timeout(2000)
            page.evaluate(r"""() => {
                const a = document.querySelector('a.lang_menu_service')
                  || Array.from(document.querySelectorAll('a,button')).find(x => /เมนูบริการ/.test((x.textContent||'').trim()));
                if (a) a.click();
            }""")
            page.wait_for_timeout(1200)
            page.evaluate(r"""() => { try { openCity('tab_RENEW_REQ', new Event('click')); } catch(e){} }""")
            page.wait_for_timeout(800)
            page.evaluate(r"""() => { const t = document.querySelector('#MT_59_MOU_RENEWAL'); if (t) t.click(); }""")
            page.wait_for_url("**/WorkPermit**", timeout=30_000)
            page.wait_for_timeout(3500)
            print("[i] อยู่ที่ฟอร์ม:", page.url)

            # คลิกปุ่ม 'ค้นหาข้อมูลคนต่างด้าว' (search_alien_modal.show)
            opened = page.evaluate(r"""() => {
                const b = Array.from(document.querySelectorAll('button, a'))
                  .find(e => /search_alien_modal\.show/.test(e.getAttribute('onclick')||''));
                if (b) { b.click(); return true; }
                return false;
            }""")
            print("[i] เปิด search_alien_modal:", opened)
            page.wait_for_timeout(2000)
            page.screenshot(path=str(SHOTS / "bt30_modal_search.png"), full_page=True)

            data = page.evaluate(MODAL_JS)
            OUT.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"[OK] เขียน {OUT.name}")
            print("modalFound:", data["modalFound"])
            print("--- FIELDS ---")
            for f in data["inputs"]:
                print("  %-22s id=%-22s ro=%-5s vis=%-5s label=%s opts=%s" % (
                    f.get("tag"), f.get("id"), f.get("readonly"), f.get("visible"),
                    (f.get("label") or "")[:30], f.get("opts")))
            print("--- BUTTONS ---")
            for b in data["buttons"]:
                print("  id=%-24s txt=%-18s onclick=%s" % (b.get("id"), b.get("txt"), b.get("onclick")))
            print("--- birthDateCheck HTML ---")
            print(data["birthDateCheckHTML"])
            input("กด Enter เพื่อปิด...")
        finally:
            ctx.close(); browser.close()


if __name__ == "__main__":
    main()
