"""Probe (read-only) — ตรวจสอบรูปแบบปีของ datepicker ช่องวันเกิด (#birthDateCheck)
ว่าใช้ พ.ศ. หรือ ค.ศ. โดยคลิกเปิดปฏิทินแล้วอ่านปีที่แสดง + ทดลองตั้งค่า
ไม่กดบันทึก/ไม่ส่งคำขอ

รัน: .venv\\Scripts\\python.exe _probe_bt30_date.py
"""
from __future__ import annotations

import json
from pathlib import Path

from playwright.sync_api import sync_playwright

from scrape_wa import login, _read_login_accounts, ROOT

SHOTS = ROOT / "screenshots"


def main() -> None:
    accounts = _read_login_accounts(ROOT / "UsernameLogin.xlsx")
    acct = next(iter(accounts.values()))
    cfg = {"username": acct["username"], "password": acct["password"],
           "user_type": acct["type"], "method": acct.get("method") or "E-Workpermit"}
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False, args=["--ignore-certificate-errors", "--start-maximized"])
        ctx = browser.new_context(locale="th-TH", ignore_https_errors=True, viewport={"width": 1600, "height": 950})
        page = ctx.new_page()
        try:
            login(page, cfg)
            page.wait_for_timeout(1500)
            page.goto("https://eworkpermit.doe.go.th/", wait_until="domcontentloaded", timeout=30_000)
            page.wait_for_timeout(2000)
            page.evaluate(r"""() => { const a=document.querySelector('a.lang_menu_service'); if(a)a.click(); }""")
            page.wait_for_timeout(1000)
            page.evaluate(r"""() => { try { openCity('tab_RENEW_REQ', new Event('click')); } catch(e){} }""")
            page.wait_for_timeout(700)
            page.evaluate(r"""() => { const t=document.querySelector('#MT_59_MOU_RENEWAL'); if(t)t.click(); }""")
            page.wait_for_url("**/WorkPermit**", timeout=30_000)
            page.wait_for_timeout(3000)
            page.evaluate(r"""() => { const b=Array.from(document.querySelectorAll('button,a')).find(e=>/search_alien_modal\.show/.test(e.getAttribute('onclick')||'')); if(b)b.click(); }""")
            page.wait_for_timeout(1500)

            # คลิกช่องวันเกิดเพื่อเปิดปฏิทิน
            page.locator("#birthDateCheck").click()
            page.wait_for_timeout(1200)
            page.screenshot(path=str(SHOTS / "bt30_datepicker_open.png"), full_page=True)

            # อ่านข้อความ header ของ widget ปฏิทินที่เปิดอยู่ (ทุกตัวที่ดูเหมือนปฏิทิน)
            info = page.evaluate(r"""() => {
                const out = { pickers: [], yearsSeen: [] };
                const cands = document.querySelectorAll(
                  '.datepicker, .ui-datepicker, .bootstrap-datetimepicker-widget, .flatpickr-calendar, [class*="datepicker"], [class*="calendar"]');
                cands.forEach(c => {
                  if (c.offsetParent === null) return;
                  const txt = (c.textContent||'').replace(/\s+/g,' ').trim().slice(0,200);
                  out.pickers.push({ cls: (c.className||'').toString().slice(0,80), txt });
                  (txt.match(/\b(19|20|25|26)\d{2}\b/g)||[]).forEach(y => out.yearsSeen.push(y));
                });
                return out;
            }""")
            print("=== DATEPICKER widgets ===")
            print(json.dumps(info, ensure_ascii=False, indent=2))

            # ทดลองตั้งค่าแบบ ค.ศ. แล้วอ่านกลับ
            for test_val in ["01/01/1995", "01/01/2538"]:
                got = page.evaluate(r"""(v) => {
                    const t = document.getElementById('birthDateCheck');
                    t.removeAttribute('readonly'); t.value = v;
                    t.dispatchEvent(new Event('input',{bubbles:true}));
                    t.dispatchEvent(new Event('change',{bubbles:true}));
                    t.dispatchEvent(new Event('blur',{bubbles:true}));
                    if (window.jQuery) { try { jQuery(t).trigger('change'); } catch(e){} }
                    return t.value;
                }""", test_val)
                err = page.evaluate(r"""() => { const e=document.querySelector('#birthDateCheck-error'); return e?e.textContent.trim():''; }""")
                print(f"set {test_val!r} -> value now {got!r}  error={err!r}")
                page.wait_for_timeout(500)

            input("กด Enter เพื่อปิด... (ดูปฏิทินบนหน้าจอว่าปีเป็น พ.ศ. หรือ ค.ศ.)")
        finally:
            ctx.close(); browser.close()


if __name__ == "__main__":
    main()
