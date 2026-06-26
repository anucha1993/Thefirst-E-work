"""Probe หน้า 2/2 ของ 'กรอกข้อมูลคำขอ' (Step 3) — เดินครบ Step1 + Step2 (2.1-2.5)
แล้ว dump โครงสร้างหน้า 2/2:
  - หัวข้อ 'โดยจะมาทำงาน' → สถานที่ทำงาน/สาขา, ประเภทกิจการ
  - หัวข้อ 'เอกสารแสดงการอนุญาตหรือการรับรอง' → เลขที่, ออกให้โดย, วันที่ออกเอกสาร, วันที่เอกสารหมดอายุ
  - ปุ่ม 'แก้ไขข้อมูล' ของแต่ละ section + modal (ถ้ามี)
  - ปุ่ม 'ถัดไป' / 'ย้อนกลับ'

ไม่กดส่งคำขอจริง — แค่ dump แล้วหยุด
รัน: .venv\\Scripts\\python.exe _probe_bt30_page2.py
"""
from __future__ import annotations

import json
from pathlib import Path

from playwright.sync_api import sync_playwright

from scrape_wa import (
    login, _read_login_accounts, _read_bt30_excel,
    _open_bt30_form, _bt30_fill_search_one, _bt30_do_step2, ROOT,
)

OUT = ROOT / "_probe_bt30_page2_out.json"
SHOTS = ROOT / "screenshots" / "bt30_page2"
SHOTS.mkdir(parents=True, exist_ok=True)

result: dict = {}


def _save() -> None:
    OUT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")


# dump ทุก input/select/textarea + ปุ่ม + heading ที่มองเห็น พร้อม label ที่ใกล้ที่สุด
PAGE_DUMP_JS = r"""() => {
  const vis = el => !!(el && el.offsetParent !== null && el.getClientRects().length);
  const txt = e => (e.textContent || e.value || '').replace(/\s+/g, ' ').trim();
  const headings = Array.from(document.querySelectorAll(
      'h1,h2,h3,h4,h5,legend,.card-header,.panel-heading,.box-title,fieldset>legend,strong,b,.section-title,label'))
    .filter(vis).map(e => txt(e).slice(0, 110)).filter(t => t.length > 2);
  const fld = e => ({
    tag: e.tagName, id: e.id, name: e.name, type: e.type,
    readonly: e.readOnly || e.hasAttribute('readonly'), disabled: e.disabled,
    cls: (e.className || '').toString().slice(0, 70),
    placeholder: e.placeholder, value: (e.value || '').slice(0, 70), visible: vis(e),
    label: (() => {
      if (e.labels && e.labels[0]) return txt(e.labels[0]).slice(0, 80);
      const p = e.closest('.form-group, .col, .row, td, .mb-3, div');
      const lab = p ? p.querySelector('label') : null;
      return lab ? txt(lab).slice(0, 80) : '';
    })(),
    opts: e.tagName === 'SELECT' ? Array.from(e.options).slice(0, 20).map(o => txt(o).slice(0, 50)) : undefined,
    nopts: e.tagName === 'SELECT' ? e.options.length : undefined,
  });
  const inputs = Array.from(document.querySelectorAll('input:not([type=hidden]), select, textarea'))
    .filter(vis).map(fld);
  const buttons = Array.from(document.querySelectorAll('button, a.btn, a[onclick], input[type=button], input[type=submit]'))
    .filter(vis).map(b => ({
      id: b.id, txt: txt(b).slice(0, 40), onclick: (b.getAttribute('onclick') || '').slice(0, 140),
      cls: (b.className || '').toString().slice(0, 60),
    })).filter(b => b.txt || b.onclick);
  return { url: location.href, headings: [...new Set(headings)], inputs, buttons };
}"""

MODAL_DUMP_JS = r"""() => {
  const vis = el => !!(el && el.offsetParent !== null && el.getClientRects().length);
  const txt = e => (e.textContent || e.value || '').replace(/\s+/g, ' ').trim();
  let modal = Array.from(document.querySelectorAll('.modal.show, .modal.in, .modal[style*="display: block"], .swal2-popup'))
    .filter(vis).pop();
  const scope = modal || document.body;
  const fld = e => ({
    tag: e.tagName, id: e.id, name: e.name, type: e.type,
    readonly: e.readOnly || e.hasAttribute('readonly'), disabled: e.disabled,
    cls: (e.className || '').toString().slice(0, 70),
    placeholder: e.placeholder, value: (e.value || '').slice(0, 60), visible: vis(e),
    label: (() => {
      if (e.labels && e.labels[0]) return txt(e.labels[0]).slice(0, 70);
      const p = e.closest('.form-group, .col, .row, td, .mb-3, div');
      const lab = p ? p.querySelector('label') : null;
      return lab ? txt(lab).slice(0, 70) : '';
    })(),
    opts: e.tagName === 'SELECT' ? Array.from(e.options).slice(0, 16).map(o => txt(o).slice(0, 45)) : undefined,
    nopts: e.tagName === 'SELECT' ? e.options.length : undefined,
  });
  const title = (() => {
    const t = scope.querySelector('.modal-title, .swal2-title, h4, h5, legend');
    return t ? txt(t).slice(0, 90) : '';
  })();
  const inputs = Array.from(scope.querySelectorAll('input:not([type=hidden]), select, textarea')).map(fld);
  const buttons = Array.from(scope.querySelectorAll('button, a.btn, input[type=button], input[type=submit]'))
    .filter(vis).map(b => ({ id: b.id, txt: txt(b).slice(0, 30), onclick: (b.getAttribute('onclick') || '').slice(0, 140) }));
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
            print("[i] Step1:", res.get("status"), res.get("note", "")[:80])
            if res.get("status") != "SUCCESS":
                print("[!] Step1 ไม่สำเร็จ — หยุด"); return
            page.wait_for_timeout(1200)

            # เดิน 2.1-2.5 → ไปหน้า 2/2
            r2 = _bt30_do_step2(page, rec, SHOTS, log=print)
            print("[i] Step2:", r2.get("step2_status"), r2.get("step2_note", "")[:120])
            page.wait_for_timeout(3500)

            # ---- dump หน้า 2/2 ----
            page.screenshot(path=str(SHOTS / "page2_overview.png"), full_page=True)
            dump = page.evaluate(PAGE_DUMP_JS)
            result["page2"] = dump
            _save()
            print("\n[OK] URL:", dump["url"])
            print("[OK] headings:")
            for h in dump["headings"]:
                print("   -", h)
            print("\n[OK] inputs/selects:")
            for f in dump["inputs"]:
                print("   %-8s id=%-32s ro=%-5s label=%s val=%s" % (
                    f["tag"], f["id"], f["readonly"], (f["label"] or "")[:45], (f["value"] or "")[:30]))
                if f.get("opts"):
                    print("        opts:", f["opts"][:10])
            print("\n[OK] buttons:")
            for b in dump["buttons"]:
                if any(k in (b["txt"] + b["onclick"]) for k in ("แก้ไข", "ถัดไป", "ย้อนกลับ", "edit", "Edit", "Modal", "บันทึก")):
                    print("   id=%-28s txt=%-16s onclick=%s" % (b["id"], b["txt"], b["onclick"]))

            # ---- เปิด edit modal แต่ละอันแล้ว dump ----
            result["modals"] = []
            n_edit = page.evaluate(r"""() => {
              const vis = el => !!(el && el.offsetParent !== null);
              const txt = e => (e.textContent || e.value || '').replace(/\s+/g,' ').trim();
              return Array.from(document.querySelectorAll('button, a.btn, a[onclick]')).filter(vis)
                .filter(b => /แก้ไข/.test(txt(b)) || /edit|Modal/i.test(b.getAttribute('onclick')||'')).length;
            }""")
            print(f"\n[i] พบปุ่มแก้ไข {n_edit} ปุ่ม — เปิดทีละอัน")
            for idx in range(n_edit):
                info = page.evaluate(r"""(idx) => {
                  const vis = el => !!(el && el.offsetParent !== null);
                  const txt = e => (e.textContent || e.value || '').replace(/\s+/g,' ').trim();
                  const btns = Array.from(document.querySelectorAll('button, a.btn, a[onclick]')).filter(vis)
                    .filter(b => /แก้ไข/.test(txt(b)) || /edit|Modal/i.test(b.getAttribute('onclick')||''));
                  const b = btns[idx];
                  if (!b) return null;
                  let sec = '';
                  let node = b;
                  for (let up = 0; up < 9 && node; up++) {
                    node = node.parentElement;
                    if (!node) break;
                    const hd = node.querySelector('h3,h4,h5,legend,.card-header,.panel-heading,.box-title,strong');
                    if (hd && txt(hd)) { sec = txt(hd).slice(0, 100); break; }
                  }
                  b.click();
                  return { idx, sec, id: b.id, txt: txt(b).slice(0, 30), onclick: (b.getAttribute('onclick')||'').slice(0,140) };
                }""", idx)
                if not info:
                    continue
                page.wait_for_timeout(1800)
                # บังคับแสดง modal เผื่อ init ไม่โชว์เอง
                page.evaluate(r"""() => {
                  if (window.jQuery) {
                    jQuery('.modal').each(function(){
                      const t = (this.textContent||'');
                      if (/ข้อมูลเพิ่มเติม|เอกสาร|อนุญาต|รับรอง|กิจการ|สถานที่/.test(t) && !/ข่าวสาร|ประชาสัมพันธ์/.test(t)) {
                        try { jQuery(this).modal('show'); } catch(e){}
                        this.style.display='block';
                      }
                    });
                  }
                }""")
                page.wait_for_timeout(800)
                md = page.evaluate(MODAL_DUMP_JS)
                md["opened_by"] = info
                result["modals"].append(md)
                _save()
                print(f"\n[OK] modal #{idx} sec='{info['sec']}' opened_by id={info['id']} onclick={info['onclick'][:80]}")
                print("     title:", md["title"], "found:", md["found"])
                for f in md["inputs"]:
                    if f.get("visible"):
                        print("     %-8s id=%-30s ro=%-5s label=%s" % (f["tag"], f["id"], f["readonly"], (f["label"] or "")[:45]))
                        if f.get("opts"):
                            print("          opts:", f["opts"][:10])
                for b in md["buttons"]:
                    if any(k in (b["txt"] + b["onclick"]) for k in ("บันทึก", "Save", "ปิด", "Close")):
                        print("     btn id=%-24s txt=%-12s onclick=%s" % (b["id"], b["txt"], b["onclick"][:70]))
                # ปิด modal
                page.evaluate(r"""() => {
                  if (window.jQuery) { try { jQuery('.modal').modal('hide'); } catch(e){} }
                  document.querySelectorAll('.modal-backdrop').forEach(x=>x.remove());
                  document.body.classList.remove('modal-open');
                  document.body.style.removeProperty('overflow');
                }""")
                page.wait_for_timeout(700)

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
