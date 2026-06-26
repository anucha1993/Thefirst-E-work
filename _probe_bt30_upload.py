"""Probe การอัปโหลดไฟล์ + โมดอล 'เพิ่มเอกสาร' (4.2) บนหน้า 'แนบเอกสาร' (Step 3).

เดินครบ Step1 + Step2 → หน้าแนบเอกสาร แล้ว:
  1) set_input_files ที่ #file_name_33 (สำเนาหนังสือเดินทาง) ด้วยไฟล์ทดสอบ
     แล้ว dump สถานะแถวหลังอัปโหลด (ชื่อไฟล์/ปุ่มลบ/preview/loading) เพื่อหา success indicator
  2) คลิก OncOpenModalAddFile() (ปุ่ม 'เพิ่มเอกสาร') แล้ว dump โครงสร้างโมดอล
     (input ทั้งหมด, file input, ปุ่มบันทึก/ตกลง) เพื่อ map การแนบ 'เอกสารอื่นๆ' ทีละไฟล์

ไม่กด 'ถัดไป' / ไม่ส่งคำขอ — แค่ dump แล้วหยุด
รัน: .venv\\Scripts\\python.exe _probe_bt30_upload.py
"""
from __future__ import annotations

import json
from pathlib import Path

from playwright.sync_api import sync_playwright

from scrape_wa import (
    login, _read_login_accounts, _read_bt30_excel,
    _open_bt30_form, _bt30_fill_search_one, _bt30_do_step2, ROOT,
)

OUT = ROOT / "_probe_bt30_upload_out.json"
SHOTS = ROOT / "screenshots" / "bt30_upload"
SHOTS.mkdir(parents=True, exist_ok=True)

TEST_FILE = ROOT / "Files" / "1.พาส.pdf"

result: dict = {}


def _save() -> None:
    OUT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")


# dump แถวของ file input ตาม id ที่ระบุ (หลังอัปโหลด) เพื่อหา success indicator
ROW_DUMP_JS = r"""(fid) => {
  const vis = el => !!(el && el.offsetParent !== null && el.getClientRects().length);
  const txt = e => (e.textContent || e.value || '').replace(/\s+/g, ' ').trim();
  const inp = document.getElementById(fid);
  if (!inp) return { found: false };
  // หา container แถว
  let row = inp;
  for (let up = 0; up < 8 && row; up++) {
    row = row.parentElement;
    if (row && txt(row).length > 12 && /เอกสาร|สำเนา|หนังสือเดินทาง/.test(txt(row))) break;
  }
  const links = row ? Array.from(row.querySelectorAll('a, button, img, span, i'))
    .filter(vis).map(e => ({ tag: e.tagName, txt: txt(e).slice(0, 40),
       id: e.id, cls: (e.className||'').toString().slice(0,50),
       href: (e.getAttribute('href')||'').slice(0,60),
       onclick: (e.getAttribute('onclick')||'').slice(0,80),
       src: (e.getAttribute('src')||'').slice(0,60) }))
    .filter(e => e.txt || e.onclick || e.href || e.src) : [];
  return { found: true, rowText: row ? txt(row).slice(0, 240) : '', links };
}"""


# dump โมดอลที่เปิดอยู่ (เพิ่มเอกสาร)
MODAL_DUMP_JS = r"""() => {
  const vis = el => !!(el && el.offsetParent !== null && el.getClientRects().length);
  const txt = e => (e.textContent || '').replace(/\s+/g, ' ').trim();
  // หาโมดอลที่กำลังแสดง
  const modals = Array.from(document.querySelectorAll('.modal, [role=dialog], .swal2-popup, .ui-dialog'))
    .filter(vis);
  const m = modals[modals.length - 1];
  if (!m) return { found: false, nModals: modals.length };
  const title = txt(m.querySelector('.modal-title, .modal-header, h1,h2,h3,h4,h5, .swal2-title') || {});
  const inputs = Array.from(m.querySelectorAll('input, select, textarea')).map(e => ({
    tag: e.tagName, type: e.type || '', id: e.id, name: e.name,
    placeholder: e.placeholder || '', cls: (e.className||'').toString().slice(0,60),
    accept: e.accept || '', visible: vis(e),
    onchange: (e.getAttribute('onchange')||'').slice(0,120),
    optionCount: e.tagName === 'SELECT' ? e.options.length : undefined,
    options: e.tagName === 'SELECT' ? Array.from(e.options).slice(0,8).map(o => txt(o).slice(0,40)) : undefined,
  }));
  const buttons = Array.from(m.querySelectorAll('button, a.btn, label.btn, label[for], input[type=button], input[type=submit]'))
    .filter(vis).map(b => ({ id: b.id, htmlFor: b.getAttribute('for')||'', txt: txt(b).slice(0,30),
       onclick: (b.getAttribute('onclick')||'').slice(0,130) }));
  return { found: true, id: m.id, cls: (m.className||'').toString().slice(0,80),
    title: title.slice(0, 120), inputs, buttons, modalText: txt(m).slice(0, 400) };
}"""


def main() -> None:
    accounts = _read_login_accounts(ROOT / "UsernameLogin.xlsx")
    acct = next(iter(accounts.values()))
    cfg = {"username": acct["username"], "password": acct["password"],
           "user_type": acct["type"], "method": acct.get("method") or "E-Workpermit"}
    recs = _read_bt30_excel(ROOT / "from_bt30.xlsx")
    rec = recs[0]
    print(f"[i] บัญชี: {cfg['username']} | คนต่างด้าว: {rec['name']}")
    print(f"[i] ไฟล์ทดสอบ: {TEST_FILE} (exists={TEST_FILE.exists()})")

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
            print("[i] Step1:", res.get("status"), res.get("note", "")[:80])
            if res.get("status") != "SUCCESS":
                print("[!] Step1 ไม่สำเร็จ — หยุด"); return
            page.wait_for_timeout(1200)

            r2 = _bt30_do_step2(page, rec, SHOTS, log=print)
            print("[i] Step2:", r2.get("step2_status"), r2.get("step2_note", "")[:140])
            page.wait_for_timeout(4000)
            page.screenshot(path=str(SHOTS / "attach_before.png"), full_page=True)

            # ---- 1) ทดสอบอัปโหลดที่ file_name_33 ----
            before = page.evaluate(ROW_DUMP_JS, "file_name_33")
            result["row_before"] = before
            print("\n[1] ก่อนอัปโหลด row33:", (before.get("rowText") or "")[:120])
            if TEST_FILE.exists():
                try:
                    page.set_input_files("#file_name_33", str(TEST_FILE))
                    print("[1] set_input_files #file_name_33 OK — รอ AJAX อัปโหลด...")
                    page.wait_for_timeout(6000)
                except Exception as e:
                    print("[1][ERR] set_input_files:", e)
            after = page.evaluate(ROW_DUMP_JS, "file_name_33")
            result["row_after"] = after
            print("[1] หลังอัปโหลด row33:", (after.get("rowText") or "")[:160])
            for l in after.get("links", []):
                print("      link:", l["tag"], "|", l["txt"][:30], "| onclick=", l["onclick"][:50], "| src=", l["src"][:40])
            page.screenshot(path=str(SHOTS / "attach_after_upload.png"), full_page=True)
            _save()

            # ---- 2) เปิดโมดอล 'เพิ่มเอกสาร' (4.2) ----
            print("\n[2] คลิก 'เพิ่มเอกสาร' (OncOpenModalAddFile())...")
            try:
                page.evaluate("() => { if (typeof OncOpenModalAddFile === 'function') OncOpenModalAddFile(); }")
            except Exception as e:
                print("[2][ERR] eval OncOpenModalAddFile:", e)
                try:
                    page.get_by_text("เพิ่มเอกสาร", exact=False).first.click()
                except Exception as e2:
                    print("[2][ERR] click เพิ่มเอกสาร:", e2)
            page.wait_for_timeout(2500)
            modal = page.evaluate(MODAL_DUMP_JS)
            result["add_modal"] = modal
            _save()
            page.screenshot(path=str(SHOTS / "add_modal.png"), full_page=True)
            if not modal.get("found"):
                print("[2] ไม่พบโมดอล (nModals=%s)" % modal.get("nModals"))
            else:
                print("[2] โมดอล id=%s title=%s" % (modal.get("id"), modal.get("title")))
                print("[2] inputs:")
                for inp in modal.get("inputs", []):
                    print("      <%s type=%s id=%s name=%s accept=%s ph='%s' vis=%s onchange=%s>" % (
                        inp["tag"], inp.get("type"), inp.get("id"), inp.get("name"),
                        inp.get("accept"), inp.get("placeholder"), inp.get("visible"),
                        (inp.get("onchange") or "")[:50]))
                    if inp.get("options"):
                        print("        options:", inp["options"])
                print("[2] buttons:")
                for b in modal.get("buttons", []):
                    print("      id=%-20s for=%-12s txt=%-16s onclick=%s" % (
                        (b["id"] or "")[:20], (b["htmlFor"] or "")[:12], (b["txt"] or "")[:16], b["onclick"][:70]))

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
