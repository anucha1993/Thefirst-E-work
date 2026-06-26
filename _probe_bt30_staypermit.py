"""Probe StayPermit modal (2.3/2.4) ของ บต.30 — dump fields ของ modal ข้อมูลเพิ่มเติม
(เอกสารการได้รับอนุญาตให้อยู่ในราชอาณาจักร + การตรวจลงตรา + วันเดินทางมาถึง)
ปิด popup ข่าวสารก่อน แล้วเปิด ModalEditAlienComponentStayPermit
ไม่กดบันทึก/ไม่ submit

รัน: .venv\\Scripts\\python.exe _probe_bt30_staypermit.py
"""
from __future__ import annotations

import json
from pathlib import Path

from playwright.sync_api import sync_playwright

from scrape_wa import (
    login, _read_login_accounts, _read_bt30_excel,
    _open_bt30_form, _bt30_fill_search_one, ROOT,
)

OUT = ROOT / "_probe_bt30_staypermit_out.json"
SHOTS = ROOT / "screenshots" / "bt30_step2"
SHOTS.mkdir(parents=True, exist_ok=True)
result: dict = {}

MODAL_DUMP_JS = r"""(modalId) => {
  const vis = el => !!(el && el.offsetParent !== null && el.getClientRects().length);
  const txt = e => (e.textContent || e.value || '').replace(/\s+/g, ' ').trim();
  let modal = modalId ? document.getElementById(modalId) : null;
  if (!modal) {
    modal = Array.from(document.querySelectorAll('.modal.show, .modal.in, .modal[style*="display: block"]'))
      .filter(vis).filter(m => !/ข่าวสาร|ประชาสัมพันธ์/.test(m.textContent||'')).pop();
  }
  const scope = modal || document.body;
  const fld = e => ({
    tag: e.tagName, id: e.id, name: e.name, type: e.type,
    readonly: e.readOnly || e.hasAttribute('readonly'), disabled: e.disabled,
    cls: (e.className || '').toString().slice(0, 60),
    placeholder: e.placeholder, value: (e.value || '').slice(0, 50), visible: vis(e),
    label: (() => {
      if (e.labels && e.labels[0]) return txt(e.labels[0]).slice(0, 70);
      const p = e.closest('.form-group, .col, .row, td, .mb-3, div');
      const lab = p ? p.querySelector('label') : null;
      return lab ? txt(lab).slice(0, 70) : '';
    })(),
    opts: e.tagName === 'SELECT' ? Array.from(e.options).slice(0, 30).map(o => txt(o).slice(0, 70)) : undefined,
    nopts: e.tagName === 'SELECT' ? e.options.length : undefined,
  });
  const title = (() => { const t = scope.querySelector('.modal-title, h4, h5, legend'); return t ? txt(t).slice(0, 90) : ''; })();
  const inputs = Array.from(scope.querySelectorAll('input:not([type=hidden]), select, textarea')).map(fld);
  const buttons = Array.from(scope.querySelectorAll('button, a.btn, input[type=button], input[type=submit]'))
    .filter(vis).map(b => ({ id: b.id, txt: txt(b).slice(0, 30), onclick: (b.getAttribute('onclick') || '').slice(0, 120) }));
  return { title, found: !!modal, modalId: modal ? modal.id : '', inputs, buttons };
}"""


def _dismiss_news(page) -> None:
    """ปิด popup ข่าวสารประชาสัมพันธ์ / modal บัง"""
    for _ in range(3):
        page.keyboard.press("Escape")
        page.wait_for_timeout(300)
    page.evaluate(r"""() => {
      const vis = el => !!(el && el.offsetParent !== null);
      document.querySelectorAll('.modal.show, .modal.in, .modal[style*="display: block"]').forEach(m => {
        if (/ข่าวสาร|ประชาสัมพันธ์/.test(m.textContent || '')) {
          const c = m.querySelector('[data-dismiss=modal],[data-bs-dismiss=modal],.close,.btn-close,button');
          if (c) c.click();
          m.classList.remove('show','in'); m.style.display='none';
        }
      });
      document.querySelectorAll('.modal-backdrop').forEach(b => b.remove());
      document.body.classList.remove('modal-open');
      document.body.style.overflow='';
    }""")
    page.wait_for_timeout(500)


def _gentle_hide_news(page) -> None:
    """ซ่อนเฉพาะ popup ข่าวสาร — ไม่กด Escape / ไม่ลบ backdrop ของ modal อื่น"""
    page.evaluate(r"""() => {
      document.querySelectorAll('.modal.show, .modal.in, .modal[style*="display: block"]').forEach(m => {
        if (/ข่าวสาร|ประชาสัมพันธ์/.test(m.textContent || '')) {
          const c = m.querySelector('[data-dismiss=modal],[data-bs-dismiss=modal],.close,.btn-close');
          if (c) c.click();
          m.classList.remove('show','in'); m.style.display='none';
        }
      });
    }""")
    page.wait_for_timeout(400)


def main() -> None:
    accounts = _read_login_accounts(ROOT / "UsernameLogin.xlsx")
    acct = next(iter(accounts.values()))
    cfg = {"username": acct["username"], "password": acct["password"],
           "user_type": acct["type"], "method": acct.get("method") or "E-Workpermit"}
    rec = _read_bt30_excel(ROOT / "from_bt30.xlsx")[0]
    print(f"[i] บัญชี {cfg['username']} | {rec['name']}")

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
            page.wait_for_timeout(1200)
            # 2.1 consent + ถัดไป
            page.evaluate(r"""() => { const c=document.getElementById('check_truth'); if(c && !c.checked) c.click(); }""")
            page.wait_for_timeout(500)
            page.evaluate(r"""() => { const b=document.getElementById('gonextSubmit'); if(b) b.click(); }""")
            page.wait_for_timeout(4500)
            print("[i] detail url:", page.url)
            _dismiss_news(page)
            page.screenshot(path=str(SHOTS / "SP_detail.png"), full_page=True)

            # เปิด StayPermit modal (เรียก init ตรง ๆ)
            opened = page.evaluate(r"""() => {
              try {
                if (typeof ModalEditAlienComponentStayPermit !== 'undefined') {
                  ModalEditAlienComponentStayPermit.init({event: new Event('click'), form_type_id:'MT_59', alien_emp_relate_id:''});
                  return 'init-called';
                }
              } catch(e) { return 'err:'+e.message; }
              // fallback: คลิกปุ่มแก้ไขที่ onclick มี StayPermit
              const b = Array.from(document.querySelectorAll('button,a')).find(x => /StayPermit/.test(x.getAttribute('onclick')||''));
              if (b) { b.click(); return 'clicked-btn'; }
              return 'not-found';
            }""")
            print("[i] เปิด StayPermit:", opened)
            page.wait_for_timeout(4500)
            _gentle_hide_news(page)
            # บังคับแสดง modal staypermitModal ผ่าน Bootstrap (init เติมข้อมูลแล้วแต่ไม่ show)
            page.evaluate(r"""() => {
              try { if (window.jQuery && jQuery('#staypermitModal').length) jQuery('#staypermitModal').modal('show'); } catch(e){}
              const m = document.getElementById('staypermitModal');
              if (m) { m.classList.add('show'); m.style.display='block'; m.removeAttribute('aria-hidden'); }
            }""")
            page.wait_for_timeout(1500)
            _gentle_hide_news(page)
            page.wait_for_timeout(600)
            # debug: รายงาน modal ที่เปิดอยู่
            result["openModals"] = page.evaluate(r"""() => {
              const vis = el => !!(el && el.offsetParent !== null);
              return Array.from(document.querySelectorAll('.modal'))
                .filter(m => m.classList.contains('show') || m.style.display==='block')
                .map(m => ({ id: m.id, vis: vis(m), title: (m.querySelector('.modal-title,h4,h5')||{}).textContent?.replace(/\s+/g,' ').trim().slice(0,60) }));
            }""")
            print("[i] open modals:", result["openModals"])
            page.screenshot(path=str(SHOTS / "SP_modal.png"), full_page=True)

            dump = page.evaluate(MODAL_DUMP_JS, "staypermitModal")
            result["staypermit"] = {"opened": opened, "modal": dump}
            OUT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"\n[OK] modal title='{dump.get('title')}' id='{dump.get('modalId')}' found={dump.get('found')}")
            print("--- FIELDS ---")
            for f in dump["inputs"]:
                print("  %-8s id=%-30s ro=%-5s vis=%-5s label=%-30s nopts=%s" % (
                    f.get("tag"), f.get("id"), f.get("readonly"), f.get("visible"),
                    (f.get("label") or "")[:30], f.get("nopts")))
                if f.get("opts"):
                    print("        opts:", f["opts"][:14])
            print("--- BUTTONS ---")
            for b in dump["buttons"]:
                print("  id=%-26s txt=%-16s onclick=%s" % (b.get("id"), b.get("txt"), b.get("onclick")))
            print(f"\n[OK] เขียน {OUT.name}")
            input("กด Enter เพื่อปิด...")
        finally:
            OUT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
            ctx.close(); browser.close()


if __name__ == "__main__":
    main()
