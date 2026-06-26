"""Probe หน้า 'แนบเอกสาร' (Step 3) — เดินครบ Step1 + Step2 (รวมหน้า 2/2) แล้ว dump:
  - รายการเอกสารแต่ละประเภท (label/heading)
  - input[type=file] ทุกตัว (id, name, accept, label/row ที่ใกล้ที่สุด)
  - ปุ่มต่างๆ (แนบไฟล์/เลือกไฟล์/อัปโหลด/ลบ/ถัดไป/ย้อนกลับ) + onclick
  - โครงสร้างแถวเอกสาร (card/row) เพื่อ map ประเภท → file input

ไม่แนบไฟล์/ไม่ส่งคำขอ — แค่ dump แล้วหยุด
รัน: .venv\\Scripts\\python.exe _probe_bt30_attach.py
"""
from __future__ import annotations

import json
from pathlib import Path

from playwright.sync_api import sync_playwright

from scrape_wa import (
    login, _read_login_accounts, _read_bt30_excel,
    _open_bt30_form, _bt30_fill_search_one, _bt30_do_step2, ROOT,
)

OUT = ROOT / "_probe_bt30_attach_out.json"
SHOTS = ROOT / "screenshots" / "bt30_attach"
SHOTS.mkdir(parents=True, exist_ok=True)

result: dict = {}


def _save() -> None:
    OUT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")


# dump รายการเอกสาร: หาแต่ละแถว/การ์ดที่มี input[type=file] แล้วเก็บ label + ปุ่มในแถวนั้น
ATTACH_DUMP_JS = r"""() => {
  const vis = el => !!(el && el.offsetParent !== null && el.getClientRects().length);
  const txt = e => (e.textContent || e.value || '').replace(/\s+/g, ' ').trim();

  // 1) file inputs ทั้งหมด (รวมที่ซ่อน)
  const fileInputs = Array.from(document.querySelectorAll('input[type=file]')).map((e, idx) => {
    // หา container แถว เพื่อดึง label ของประเภทเอกสาร
    let row = e.closest('.row, .card, .list-group-item, li, tr, .form-group, .mb-3, .doc-item, .attachment, .file-item');
    let block = e;
    for (let up = 0; up < 8 && block; up++) {
      block = block.parentElement;
      if (block && txt(block).length > 12 && /เอกสาร|สำเนา|ใบ|รูปถ่าย|หนังสือ|บัตร|อนุญาต/.test(txt(block))) { row = block; break; }
    }
    return {
      idx, id: e.id, name: e.name, accept: e.accept, multiple: e.multiple,
      disabled: e.disabled, visible: vis(e),
      cls: (e.className || '').toString().slice(0, 60),
      onchange: (e.getAttribute('onchange') || '').slice(0, 120),
      rowText: row ? txt(row).slice(0, 160) : '',
      rowButtons: row ? Array.from(row.querySelectorAll('button, a.btn, label.btn, label[for], input[type=button]'))
        .filter(vis).map(b => ({ id: b.id, htmlFor: b.getAttribute('for') || '', txt: txt(b).slice(0, 30),
          onclick: (b.getAttribute('onclick') || '').slice(0, 120) })) : [],
    };
  });

  // 2) heading ของแต่ละประเภทเอกสาร
  const headings = Array.from(document.querySelectorAll(
      'h1,h2,h3,h4,h5,h6,legend,.card-header,.panel-heading,.box-title,label,strong,b,td,.doc-name,.file-title,.list-group-item'))
    .filter(vis).map(e => txt(e)).filter(t => t.length > 6
       && /เอกสาร|สำเนา|ใบรับรอง|ใบอนุญาต|รูปถ่าย|หนังสือ|บัตรประจำตัว|อนุญาตให้เข้า|สัญญาจ้าง|บต\.|มอบอำนาจ|อากรแสตมป์/.test(t))
    .map(t => t.slice(0, 170));

  // 3) ปุ่มหลัก (ถัดไป/ย้อนกลับ/บันทึกร่าง/อัปโหลด)
  const buttons = Array.from(document.querySelectorAll('button, a.btn, a[onclick], input[type=button], input[type=submit], label.btn'))
    .filter(vis).map(b => ({
      id: b.id, txt: txt(b).slice(0, 36), htmlFor: b.getAttribute('for') || '',
      onclick: (b.getAttribute('onclick') || '').slice(0, 130), cls: (b.className || '').toString().slice(0, 60),
    })).filter(b => b.txt || b.onclick || b.htmlFor);

  return { url: location.href, nFileInputs: fileInputs.length, fileInputs, headings: [...new Set(headings)], buttons };
}"""


def main() -> None:
    accounts = _read_login_accounts(ROOT / "UsernameLogin.xlsx")
    acct = next(iter(accounts.values()))
    cfg = {"username": acct["username"], "password": acct["password"],
           "user_type": acct["type"], "method": acct.get("method") or "E-Workpermit"}
    recs = _read_bt30_excel(ROOT / "from_bt30.xlsx")
    rec = recs[0]
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
            print("[i] Step1:", res.get("status"), res.get("note", "")[:80])
            if res.get("status") != "SUCCESS":
                print("[!] Step1 ไม่สำเร็จ — หยุด"); return
            page.wait_for_timeout(1200)

            # เดิน 2.1-2.7 → ไปหน้า 'แนบเอกสาร'
            r2 = _bt30_do_step2(page, rec, SHOTS, log=print)
            print("[i] Step2:", r2.get("step2_status"), r2.get("step2_note", "")[:140])
            page.wait_for_timeout(4000)

            page.screenshot(path=str(SHOTS / "attach_overview.png"), full_page=True)
            dump = page.evaluate(ATTACH_DUMP_JS)
            result["attach"] = dump
            _save()

            print("\n[OK] URL:", dump["url"])
            print("[OK] จำนวน input[type=file]:", dump["nFileInputs"])
            print("\n[OK] headings (ประเภทเอกสาร):")
            for h in dump["headings"]:
                print("   -", h[:120])
            print("\n[OK] file inputs:")
            for f in dump["fileInputs"]:
                print("  #%-2d id=%-26s name=%-20s accept=%-18s mult=%-5s vis=%-5s" % (
                    f["idx"], (f["id"] or "")[:26], (f["name"] or "")[:20], (f["accept"] or "")[:18],
                    f["multiple"], f["visible"]))
                if f["rowText"]:
                    print("       row:", f["rowText"][:130])
                for b in f["rowButtons"]:
                    print("       btn id=%-18s for=%-14s txt=%-14s onclick=%s" % (
                        (b["id"] or "")[:18], (b["htmlFor"] or "")[:14], (b["txt"] or "")[:14], b["onclick"][:70]))
            print("\n[OK] ปุ่มหลัก:")
            for b in dump["buttons"]:
                if any(k in (b["txt"] + b["onclick"] + b["id"]) for k in ("ถัดไป", "ย้อนกลับ", "บันทึก", "อัปโหลด", "upload", "Upload", "Next", "Step", "แนบ", "เลือกไฟล์")):
                    print("   id=%-26s for=%-12s txt=%-16s onclick=%s" % (
                        (b["id"] or "")[:26], (b["htmlFor"] or "")[:12], (b["txt"] or "")[:16], b["onclick"][:70]))

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
