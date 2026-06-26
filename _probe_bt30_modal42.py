"""Probe โมดอล 'เพิ่มเอกสาร' (4.2 เอกสารอื่นๆ) — เจาะลึกว่าเปิดอย่างไร.

เดินครบ Step1 + Step2 → หน้าแนบเอกสาร แล้ว:
  1) ตรวจ global function: typeof OncOpenModalAddFile + .toString() + รายชื่อ fn ที่เกี่ยวกับ modal/addfile
  2) คลิกปุ่ม 'เพิ่มเอกสาร' จริง (ไม่ eval) → รอ → dump โมดอลที่โผล่ (inputs/file/select/ปุ่มบันทึก)
ไม่ส่งคำขอ — แค่ dump แล้วหยุด
รัน: .venv\\Scripts\\python.exe _probe_bt30_modal42.py
"""
from __future__ import annotations

import json
from pathlib import Path

from playwright.sync_api import sync_playwright

from scrape_wa import (
    login, _read_login_accounts, _read_bt30_excel,
    _open_bt30_form, _bt30_fill_search_one, _bt30_do_step2, ROOT,
)

OUT = ROOT / "_probe_bt30_modal42_out.json"
SHOTS = ROOT / "screenshots" / "bt30_modal42"
SHOTS.mkdir(parents=True, exist_ok=True)
result: dict = {}


def _save() -> None:
    OUT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")


INVESTIGATE_JS = r"""() => {
  const fns = Object.getOwnPropertyNames(window)
    .filter(n => /modal|addfile|add_file|adddoc|add_doc/i.test(n) && typeof window[n] === 'function');
  let src = 'NOT_GLOBAL_FUNCTION';
  try { if (typeof OncOpenModalAddFile === 'function') src = OncOpenModalAddFile.toString().slice(0, 1200); } catch (e) { src = 'ERR:' + e.message; }
  const modals = Array.from(document.querySelectorAll('[id*=odal],[class*=odal],[role=dialog]'))
    .map(e => ({ id: e.id, cls: (e.className || '').toString().slice(0, 50),
       disp: getComputedStyle(e).display }));
  return { fns, src, modals };
}"""

MODAL_DUMP_JS = r"""() => {
  const vis = el => !!(el && getComputedStyle(el).display !== 'none' && el.getClientRects().length);
  const txt = e => (e.textContent || '').replace(/\s+/g, ' ').trim();
  const cand = Array.from(document.querySelectorAll('.modal, [role=dialog], .swal2-popup, .ui-dialog'))
    .filter(vis);
  const m = cand[cand.length - 1];
  if (!m) {
    // เผื่อ modal เป็น element อื่น: หา element ที่เพิ่ง display ขึ้นและมี file input
    const anyFile = Array.from(document.querySelectorAll('input[type=file]'))
      .filter(f => vis(f.closest('[class*=odal]') || f)).map(f => ({ id: f.id, name: f.name }));
    return { found: false, nCand: cand.length, anyVisibleFile: anyFile };
  }
  const title = txt(m.querySelector('.modal-title, .modal-header, h1,h2,h3,h4,h5, .swal2-title') || { textContent: '' });
  const inputs = Array.from(m.querySelectorAll('input, select, textarea')).map(e => ({
    tag: e.tagName, type: e.type || '', id: e.id, name: e.name,
    placeholder: e.placeholder || '', accept: e.accept || '', visible: vis(e),
    cls: (e.className || '').toString().slice(0, 50),
    onchange: (e.getAttribute('onchange') || '').slice(0, 120),
    optionCount: e.tagName === 'SELECT' ? e.options.length : undefined,
    options: e.tagName === 'SELECT' ? Array.from(e.options).slice(0, 10).map(o => txt(o).slice(0, 50)) : undefined,
  }));
  const buttons = Array.from(m.querySelectorAll('button, a.btn, label.btn, label[for], input[type=button], input[type=submit]'))
    .filter(vis).map(b => ({ id: b.id, htmlFor: b.getAttribute('for') || '', txt: txt(b).slice(0, 30),
       onclick: (b.getAttribute('onclick') || '').slice(0, 130) }));
  return { found: true, id: m.id, cls: (m.className || '').toString().slice(0, 80),
    title: title.slice(0, 140), inputs, buttons, modalText: txt(m).slice(0, 500) };
}"""


def main() -> None:
    accounts = _read_login_accounts(ROOT / "UsernameLogin.xlsx")
    acct = next(iter(accounts.values()))
    cfg = {"username": acct["username"], "password": acct["password"],
           "user_type": acct["type"], "method": acct.get("method") or "E-Workpermit"}
    rec = _read_bt30_excel(ROOT / "from_bt30.xlsx")[0]
    print(f"[i] บัญชี: {cfg['username']} | คนต่างด้าว: {rec['name']}")

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False, args=["--ignore-certificate-errors", "--start-maximized"])
        ctx = browser.new_context(locale="th-TH", ignore_https_errors=True, viewport={"width": 1600, "height": 950})
        page = ctx.new_page()
        try:
            login(page, cfg)
            page.wait_for_timeout(1500)
            if not _open_bt30_form(page, log=print):
                print("[!] เปิดฟอร์มไม่สำเร็จ"); return
            res = _bt30_fill_search_one(page, rec, SHOTS, log=print)
            print("[i] Step1:", res.get("status"))
            if res.get("status") != "SUCCESS":
                print("[!] Step1 ไม่สำเร็จ — หยุด"); return
            page.wait_for_timeout(1200)
            r2 = _bt30_do_step2(page, rec, SHOTS, log=print)
            print("[i] Step2:", r2.get("step2_status"))
            page.wait_for_timeout(4000)

            inv = page.evaluate(INVESTIGATE_JS)
            result["investigate"] = inv
            _save()
            print("\n[1] global fns:", inv["fns"])
            print("[1] OncOpenModalAddFile src:\n", inv["src"][:1000])
            print("[1] modal-ish elements:")
            for m in inv["modals"]:
                print("     id=%-26s disp=%-8s cls=%s" % ((m["id"] or "")[:26], m["disp"], m["cls"]))

            # คลิกปุ่ม 'เพิ่มเอกสาร' จริง
            print("\n[2] คลิกปุ่ม 'เพิ่มเอกสาร' จริง...")
            clicked = page.evaluate(r"""() => {
              const b = Array.from(document.querySelectorAll('button, a, label'))
                .find(e => /เพิ่มเอกสาร/.test((e.textContent || '').trim())
                      || /OncOpenModalAddFile/.test(e.getAttribute('onclick') || ''));
              if (b) { b.click(); return (b.textContent||'').trim().slice(0,30); }
              return null;
            }""")
            print("[2] clicked:", clicked)
            page.wait_for_timeout(4000)
            modal = page.evaluate(MODAL_DUMP_JS)
            result["add_modal"] = modal
            _save()
            page.screenshot(path=str(SHOTS / "add_modal.png"), full_page=True)
            if not modal.get("found"):
                print("[2] ไม่พบโมดอล nCand=%s anyVisibleFile=%s" % (modal.get("nCand"), modal.get("anyVisibleFile")))
            else:
                print("[2] โมดอล id=%s cls=%s title=%s" % (modal.get("id"), modal.get("cls"), modal.get("title")))
                print("[2] inputs:")
                for inp in modal.get("inputs", []):
                    print("      <%s type=%s id=%s name=%s accept=%s ph='%s' vis=%s onchange=%s>" % (
                        inp["tag"], inp.get("type"), inp.get("id"), inp.get("name"),
                        inp.get("accept"), inp.get("placeholder"), inp.get("visible"), (inp.get("onchange") or "")[:50]))
                    if inp.get("options"):
                        print("        options:", inp["options"])
                print("[2] buttons:")
                for b in modal.get("buttons", []):
                    print("      id=%-18s for=%-12s txt=%-18s onclick=%s" % (
                        (b["id"] or "")[:18], (b["htmlFor"] or "")[:12], (b["txt"] or "")[:18], b["onclick"][:70]))
                print("[2] modalText:", modal.get("modalText", "")[:300])

            print("\n[done] dump → ", OUT)
        except Exception as e:
            print("[ERR]", e)
            page.screenshot(path=str(SHOTS / "error.png"), full_page=True)
            _save()
        finally:
            input("กด Enter เพื่อปิดเบราว์เซอร์...")
            ctx.close()
            browser.close()


if __name__ == "__main__":
    main()
