"""Probe Step 2 (read-only-ish) — เดินจาก Step 1 ต่อไปยังหน้า "ถัดไป" ของ บต.30
แล้ว dump โครงสร้าง:
  - 2.1 consent checkbox + ปุ่ม 'ถัดไป'
  - หน้า detail: section ที่อยู่ในประเทศไทย / เอกสารการได้รับอนุญาตให้อยู่ฯ / ตม.
  - ปุ่ม 'แก้ไขข้อมูล' ของแต่ละ section + modal 'ข้อมูลเพิ่มเติม' (fields/options)

ไม่กดส่งคำขอจริง (ไม่ submit ปลายทาง) — แค่เปิด modal เพื่อ dump แล้วปิด
รัน: .venv\\Scripts\\python.exe _probe_bt30_step2.py
"""
from __future__ import annotations

import json
from pathlib import Path

from playwright.sync_api import sync_playwright

from scrape_wa import (
    login, _read_login_accounts, _read_bt30_excel,
    _open_bt30_form, _bt30_fill_search_one, ROOT,
)

OUT = ROOT / "_probe_bt30_step2_out.json"
SHOTS = ROOT / "screenshots" / "bt30_step2"
SHOTS.mkdir(parents=True, exist_ok=True)

result: dict = {}


def _save() -> None:
    OUT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")


# dump ทุกปุ่ม/checkbox/heading ที่มองเห็นในหน้า (ภาพรวม)
PAGE_OVERVIEW_JS = r"""() => {
  const vis = el => !!(el && el.offsetParent !== null && el.getClientRects().length);
  const txt = e => (e.textContent || e.value || '').replace(/\s+/g, ' ').trim();
  const headings = Array.from(document.querySelectorAll(
      'h1,h2,h3,h4,h5,legend,.card-header,.panel-heading,.box-title,fieldset>legend,strong,b,.section-title'))
    .filter(vis).map(e => txt(e).slice(0, 90)).filter(t => t.length > 2);
  const buttons = Array.from(document.querySelectorAll('button, a.btn, a[onclick], input[type=button], input[type=submit]'))
    .filter(vis).map(b => ({
      id: b.id, txt: txt(b).slice(0, 40), onclick: (b.getAttribute('onclick') || '').slice(0, 120),
      cls: (b.className || '').toString().slice(0, 60),
    })).filter(b => b.txt || b.onclick);
  const checks = Array.from(document.querySelectorAll('input[type=checkbox]')).map(c => ({
    id: c.id, name: c.name, checked: c.checked, visible: vis(c),
    label: (() => {
      if (c.labels && c.labels[0]) return txt(c.labels[0]).slice(0, 160);
      const p = c.closest('label, .form-check, .checkbox, div, td');
      return p ? txt(p).slice(0, 160) : '';
    })(),
  }));
  return { url: location.href, headings: [...new Set(headings)], buttons, checks };
}"""

# dump fields ของ modal ที่กำลังเปิด (visible) — generic
MODAL_DUMP_JS = r"""() => {
  const vis = el => !!(el && el.offsetParent !== null && el.getClientRects().length);
  const txt = e => (e.textContent || e.value || '').replace(/\s+/g, ' ').trim();
  // หา modal ที่เปิดอยู่
  let modal = Array.from(document.querySelectorAll('.modal.show, .modal.in, .modal[style*="display: block"], .swal2-popup'))
    .filter(vis).pop();
  if (!modal) {
    // fallback: div ที่มี input และ visible และมีคำว่า ข้อมูลเพิ่มเติม / บันทึก
    modal = Array.from(document.querySelectorAll('div')).filter(d =>
      vis(d) && d.querySelector('input,select') && /ข้อมูลเพิ่มเติม|บันทึก/.test(d.textContent || '')
    ).sort((a, b) => (a.textContent || '').length - (b.textContent || '').length)[0];
  }
  const scope = modal || document.body;
  const fld = e => ({
    tag: e.tagName, id: e.id, name: e.name, type: e.type,
    readonly: e.readOnly || e.hasAttribute('readonly'), disabled: e.disabled,
    cls: (e.className || '').toString().slice(0, 70),
    placeholder: e.placeholder, value: (e.value || '').slice(0, 50), visible: vis(e),
    label: (() => {
      if (e.labels && e.labels[0]) return txt(e.labels[0]).slice(0, 60);
      const p = e.closest('.form-group, .col, .row, td, .mb-3, div');
      const lab = p ? p.querySelector('label') : null;
      return lab ? txt(lab).slice(0, 60) : '';
    })(),
    opts: e.tagName === 'SELECT' ? Array.from(e.options).slice(0, 12).map(o => txt(o).slice(0, 40)) : undefined,
    nopts: e.tagName === 'SELECT' ? e.options.length : undefined,
  });
  const title = (() => {
    const t = scope.querySelector('.modal-title, .swal2-title, h4, h5, legend');
    return t ? txt(t).slice(0, 80) : '';
  })();
  const inputs = Array.from(scope.querySelectorAll('input:not([type=hidden]), select, textarea')).map(fld);
  const buttons = Array.from(scope.querySelectorAll('button, a.btn, input[type=button], input[type=submit]'))
    .filter(vis).map(b => ({ id: b.id, txt: txt(b).slice(0, 30), onclick: (b.getAttribute('onclick') || '').slice(0, 120) }));
  return { title, found: !!modal, inputs, buttons };
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
            print("[i] Step1 fill result:", res.get("status"), res.get("note", "")[:80])
            page.wait_for_timeout(1500)

            # ---- Phase A: หน้าหลัง บันทึก (หา consent checkbox + ถัดไป) ----
            page.screenshot(path=str(SHOTS / "A_after_save.png"), full_page=True)
            result["phaseA_afterSave"] = page.evaluate(PAGE_OVERVIEW_JS)
            _save()
            print(f"[OK] phaseA headings: {result['phaseA_afterSave']['headings'][:8]}")
            print(f"[OK] phaseA checks: {len(result['phaseA_afterSave']['checks'])}")
            for c in result["phaseA_afterSave"]["checks"]:
                print("   chk id=%-20s vis=%-5s label=%s" % (c["id"], c["visible"], c["label"][:90]))
            print("[OK] phaseA buttons with 'ถัดไป/ยินยอม':")
            for b in result["phaseA_afterSave"]["buttons"]:
                if any(k in (b["txt"] + b["onclick"]) for k in ("ถัดไป", "ยินยอม", "next", "Next", "gonext", "Submit", "consent")):
                    print("   btn id=%-22s txt=%-14s onclick=%s" % (b["id"], b["txt"], b["onclick"]))

            # ---- Phase B: ติ๊ก consent + กดถัดไป ----
            checked = page.evaluate(r"""() => {
              const vis = el => !!(el && el.offsetParent !== null);
              const boxes = Array.from(document.querySelectorAll('input[type=checkbox]')).filter(vis);
              // เลือกตัวที่ label มีคำว่า 'ขอรับรอง' หรือ 'ยินยอม' หรือ 'ความประสงค์'
              let target = boxes.find(c => {
                const p = c.closest('label, .form-check, .checkbox, div, td');
                const t = p ? (p.textContent || '') : '';
                return /ขอรับรอง|ยินยอม|ความประสงค์|ดำเนินการแทน/.test(t);
              }) || boxes[0];
              if (target) { if (!target.checked) target.click(); return {id: target.id, checked: target.checked}; }
              return null;
            }""")
            print("[i] ติ๊ก consent:", checked)
            page.wait_for_timeout(800)
            page.screenshot(path=str(SHOTS / "B_consent_checked.png"), full_page=True)

            clicked_next = page.evaluate(r"""() => {
              const vis = el => !!(el && el.offsetParent !== null);
              const txt = e => (e.textContent || e.value || '').replace(/\s+/g, ' ').trim();
              const btns = Array.from(document.querySelectorAll('button, a.btn, input[type=button], input[type=submit]')).filter(vis);
              let b = btns.find(x => /^ถัดไป|ถัดไป$|ถัดไป/.test(txt(x)) && !/ค้นหา/.test(txt(x)));
              if (!b) b = btns.find(x => /gonext|next/i.test(x.getAttribute('onclick') || x.id || ''));
              if (b) { b.click(); return {id: b.id, txt: txt(b), onclick: (b.getAttribute('onclick')||'').slice(0,80)}; }
              return null;
            }""")
            print("[i] กดถัดไป:", clicked_next)
            result["nextButton_2_1"] = clicked_next
            _save()
            page.wait_for_timeout(4000)

            # ---- Phase C: หน้า detail (sections + edit buttons) ----
            page.screenshot(path=str(SHOTS / "C_detail_page.png"), full_page=True)
            overview = page.evaluate(PAGE_OVERVIEW_JS)
            result["phaseC_detailPage"] = overview
            _save()
            print("\n[OK] phaseC URL:", overview["url"])
            print("[OK] phaseC headings:")
            for h in overview["headings"]:
                print("   -", h)
            print("[OK] phaseC ปุ่ม 'แก้ไข':")
            edit_btns = [b for b in overview["buttons"] if "แก้ไข" in b["txt"] or "แก้ไข" in b["onclick"] or "edit" in (b["onclick"].lower())]
            for b in edit_btns:
                print("   btn id=%-26s txt=%-16s onclick=%s" % (b["id"], b["txt"], b["onclick"]))

            # ---- Phase D: เปิด edit modal แต่ละอันแล้ว dump ----
            result["phaseD_modals"] = []
            # ใช้ index ของ edit button ในหน้า (re-query ทุกครั้งกัน stale)
            n_edit = page.evaluate(r"""() => {
              const vis = el => !!(el && el.offsetParent !== null);
              const txt = e => (e.textContent || e.value || '').replace(/\s+/g,' ').trim();
              return Array.from(document.querySelectorAll('button, a.btn, a[onclick]')).filter(vis)
                .filter(b => /แก้ไข/.test(txt(b)) || /edit/i.test(b.getAttribute('onclick')||'')).length;
            }""")
            print(f"\n[i] พบปุ่มแก้ไข {n_edit} ปุ่ม — เปิดทีละอัน")
            for idx in range(n_edit):
                info = page.evaluate(r"""(idx) => {
                  const vis = el => !!(el && el.offsetParent !== null);
                  const txt = e => (e.textContent || e.value || '').replace(/\s+/g,' ').trim();
                  const btns = Array.from(document.querySelectorAll('button, a.btn, a[onclick]')).filter(vis)
                    .filter(b => /แก้ไข/.test(txt(b)) || /edit/i.test(b.getAttribute('onclick')||''));
                  const b = btns[idx];
                  if (!b) return null;
                  // หา heading ของ section ที่ใกล้ที่สุดด้านบน
                  let sec = '';
                  let node = b;
                  for (let up = 0; up < 8 && node; up++) {
                    node = node.parentElement;
                    if (!node) break;
                    const hd = node.querySelector('h3,h4,h5,legend,.card-header,.panel-heading,.box-title,strong');
                    if (hd && txt(hd)) { sec = txt(hd).slice(0, 90); break; }
                  }
                  b.click();
                  return { idx, btnId: b.id, btnTxt: txt(b).slice(0,30), onclick: (b.getAttribute('onclick')||'').slice(0,120), section: sec };
                }""", idx)
                page.wait_for_timeout(2200)
                dump = page.evaluate(MODAL_DUMP_JS)
                try:
                    page.screenshot(path=str(SHOTS / f"D_modal_{idx}.png"), full_page=True)
                except Exception:
                    pass
                entry = {"trigger": info, "modal": dump}
                result["phaseD_modals"].append(entry)
                _save()
                print(f"\n--- modal #{idx} (section={info.get('section') if info else '?'}) title='{dump.get('title')}' found={dump.get('found')} ---")
                for f in dump["inputs"]:
                    print("   %-8s id=%-26s ro=%-5s vis=%-5s label=%-26s opts=%s" % (
                        f.get("tag"), f.get("id"), f.get("readonly"), f.get("visible"),
                        (f.get("label") or "")[:26], (f.get("opts") or "")))
                print("   buttons:", [(b["id"], b["txt"]) for b in dump["buttons"]])
                # ปิด modal (ปุ่มปิด/ยกเลิก/close) โดยไม่กดบันทึก
                page.evaluate(r"""() => {
                  const vis = el => !!(el && el.offsetParent !== null);
                  const txt = e => (e.textContent || e.value || '').replace(/\s+/g,' ').trim();
                  const modal = Array.from(document.querySelectorAll('.modal.show, .modal.in, .modal[style*="display: block"]')).filter(vis).pop();
                  const scope = modal || document;
                  const closeBtns = Array.from(scope.querySelectorAll('button, a, .close, [data-dismiss=modal], [data-bs-dismiss=modal]')).filter(vis);
                  let c = closeBtns.find(x => /ปิด|ยกเลิก|close|×|✕/i.test(txt(x)) || x.getAttribute('data-dismiss') || x.getAttribute('data-bs-dismiss'));
                  if (c) c.click();
                }""")
                page.wait_for_timeout(1200)

            print(f"\n[OK] เขียนผลทั้งหมดที่ {OUT.name}")
            input("กด Enter เพื่อปิดเบราว์เซอร์...")
        finally:
            _save()
            ctx.close(); browser.close()


if __name__ == "__main__":
    main()
