"""Debug ละเอียดการบันทึกโมดอล 'เพิ่มเอกสาร' (4.2) — ดูทีละสเต็ปว่าพังตรงไหน.
เดินถึงหน้าแนบเอกสาร (เรียก sub-func ตรงๆ ข้ามการอัปโหลด 11 ช่อง) แล้ว:
  เปิดโมดอล → แนบไฟล์ → สังเกต file_name_select/preview → กรอกชื่อ → ยืนยัน → จับ error/รายการ
รัน: .venv\\Scripts\\python.exe _probe_bt30_modal42b.py
"""
from __future__ import annotations

import json
from pathlib import Path

from playwright.sync_api import sync_playwright

from scrape_wa import (
    login, _read_login_accounts, _read_bt30_excel, _open_bt30_form,
    _bt30_fill_search_one, _bt30_consent_next, _bt30_fill_address,
    _bt30_fill_staypermit, _bt30_step2_next, _bt30_fill_page2, _bt30_page2_next,
    ROOT,
)

SHOTS = ROOT / "screenshots" / "bt30_modal42b"
SHOTS.mkdir(parents=True, exist_ok=True)
OUT = ROOT / "_probe_bt30_modal42b_out.json"
log = print
trace: list = []


def snap(page, tag):
    d = page.evaluate(r"""() => {
      const txt = e => (e ? (e.textContent || '').replace(/\s+/g,' ').trim() : '');
      const val = id => { const e = document.getElementById(id); return e ? (e.value || '') : '__none__'; };
      const modal = document.getElementById('AddFileOrther');
      const swal = document.querySelector('.swal2-popup');
      const showbox = document.getElementById('ShowtxtBox_File');
      const fileInputs = modal ? Array.from(modal.querySelectorAll('input[type=file]')).map(e => ({id:e.id, name:e.name})) : [];
      const textInputs = modal ? Array.from(modal.querySelectorAll('input[type=text]')).map(e => ({id:e.id, val:e.value})) : [];
      const selHidden = modal ? Array.from(modal.querySelectorAll('input[type=hidden]')).map(e => ({id:e.id, val:e.value})) : [];
      return {
        modalShown: modal ? (modal.classList.contains('show')) : false,
        fileInputs, textInputs, selHidden,
        checklen: document.querySelectorAll('#ShowtxtBox_File .checklength').length,
        showboxText: txt(showbox).slice(0, 200),
        swalText: swal ? txt(swal).slice(0, 200) : '',
        modalBodySnippet: modal ? (modal.querySelector('.modal-body') ? modal.querySelector('.modal-body').innerHTML.replace(/\s+/g,' ').slice(0, 600) : '') : '',
      };
    }""")
    trace.append({"tag": tag, **d})
    OUT.write_text(json.dumps(trace, ensure_ascii=False, indent=2), encoding="utf-8")
    log(f"\n=== {tag} ===")
    log("  modalShown:", d["modalShown"], "| checklen:", d["checklen"])
    log("  fileInputs:", d["fileInputs"])
    log("  textInputs:", d["textInputs"])
    log("  selHidden:", d["selHidden"])
    if d["swalText"]:
        log("  SWAL:", d["swalText"])
    if d["showboxText"]:
        log("  showbox:", d["showboxText"])
    return d


def main():
    accounts = _read_login_accounts(ROOT / "UsernameLogin.xlsx")
    acct = next(iter(accounts.values()))
    cfg = {"username": acct["username"], "password": acct["password"],
           "user_type": acct["type"], "method": acct.get("method") or "E-Workpermit"}
    rec = _read_bt30_excel(ROOT / "from_bt30.xlsx")[0]
    fp = rec["doc_others"][0]
    log(f"[i] ไฟล์ทดสอบ 4.2: {fp}")

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False, args=["--ignore-certificate-errors", "--start-maximized"])
        ctx = browser.new_context(locale="th-TH", ignore_https_errors=True, viewport={"width": 1600, "height": 950})
        page = ctx.new_page()
        try:
            login(page, cfg); page.wait_for_timeout(1500)
            if not _open_bt30_form(page, log=log):
                log("[!] เปิดฟอร์มไม่สำเร็จ"); return
            r = _bt30_fill_search_one(page, rec, SHOTS, log=log)
            if r.get("status") != "SUCCESS":
                log("[!] Step1 ไม่สำเร็จ"); return
            page.wait_for_timeout(1200)
            _bt30_consent_next(page, log=log)
            _bt30_fill_address(page, rec, log=log)
            _bt30_fill_staypermit(page, rec, log=log)
            _bt30_step2_next(page, log=log)
            page.wait_for_timeout(1500)
            _bt30_fill_page2(page, rec, log=log)
            _bt30_page2_next(page, log=log)
            page.wait_for_timeout(3000)
            log("[i] อยู่หน้าแนบเอกสารแล้ว — เริ่ม debug โมดอล 4.2")

            # เปิดโมดอล
            page.evaluate(r"""() => {
              const b = Array.from(document.querySelectorAll('button,a,label'))
                .find(e => /เพิ่มเอกสาร/.test((e.textContent||'').trim()) || /OncOpenModalAddFile\(/.test(e.getAttribute('onclick')||''));
              if (b) b.click();
            }""")
            page.wait_for_timeout(1500)
            snap(page, "หลังเปิดโมดอล")

            # แนบไฟล์ที่ input เฉพาะ (onchange=AddNewFileOrther)
            try:
                page.locator("#AddFileOrther input[name='file_name_orther']").last.set_input_files(fp)
                log("[i] set_input_files ที่ input[name=file_name_orther] แล้ว")
            except Exception as e:
                log("[ERR] set_input_files:", e)
            # สังเกตหลังแนบไฟล์
            for i in range(8):
                page.wait_for_timeout(800)
                d = snap(page, f"หลังแนบไฟล์ +{(i+1)*0.8:.1f}s")
                if any(s.get("val") for s in d["selHidden"]):
                    log("  >> file_name_select มีค่าแล้ว")
                    break
            page.screenshot(path=str(SHOTS / "after_setfile.png"), full_page=True)

            # กรอกชื่อเอกสาร
            try:
                page.locator("#AddFileOrther input[type=text]").last.fill(Path(fp).stem)
                log("[i] กรอกชื่อเอกสารแล้ว:", Path(fp).stem)
            except Exception as e:
                log("[ERR] fill name:", e)
            page.wait_for_timeout(500)
            snap(page, "หลังกรอกชื่อ")

            # ยืนยัน
            page.evaluate("() => { const b=document.getElementById('buttonSaveFileSelect'); if(b) b.click(); }")
            log("[i] คลิกยืนยัน (SaveAddFileOrther) แล้ว")
            for i in range(10):
                page.wait_for_timeout(700)
                d = snap(page, f"หลังยืนยัน +{(i+1)*0.7:.1f}s")
                if d["checklen"] > 0:
                    log("  >> รายการเพิ่มขึ้นแล้ว! checklen =", d["checklen"]); break
            page.screenshot(path=str(SHOTS / "after_confirm.png"), full_page=True)

            log("\n[done] trace →", OUT)
        except Exception as e:
            log("[ERR]", e)
            page.screenshot(path=str(SHOTS / "error.png"), full_page=True)
        finally:
            input("กด Enter เพื่อปิดเบราว์เซอร์...")
            ctx.close(); browser.close()


if __name__ == "__main__":
    main()
