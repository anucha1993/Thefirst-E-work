"""Probe: หา DOM ของช่อง 'วันที่ยื่นคำขอ' / 'ถึง' บนหน้า e-Tracking
รัน: python _probe_date_filter.py
จะเปิด browser แบบเห็นหน้า → login → เปิด tracking → dump input ที่มี placeholder 'วว/ดด/ปปปป'
พร้อมทั้ง label, parent HTML, และปุ่มค้นหา (ถ้ามี)
"""
from __future__ import annotations
import json
import os
import sys
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()

from scrape_wa import login, goto_tracking  # noqa: E402


def main() -> None:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
        "headless": False,
    }
    print(f"[i] login as {cfg['username']!r}")
    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False,
                               args=["--ignore-certificate-errors", "--start-maximized"])
        ctx = b.new_context(locale="th-TH", ignore_https_errors=True,
                            viewport={"width": 1600, "height": 1000})
        page = ctx.new_page()
        login(page, cfg)
        print(f"[i] login ok — url: {page.url}")
        goto_tracking(page)
        page.wait_for_timeout(3000)
        print(f"[i] tracking url: {page.url}")

        # ---- 1) หาทุก input ที่มี placeholder 'วว/ดด/ปปปป' ----
        info = page.evaluate(
            r"""() => {
              const out = {inputs: [], searchButtons: [], dateScripts: []};
              const inputs = Array.from(document.querySelectorAll(
                'input[placeholder="วว/ดด/ปปปป"]'
              ));
              inputs.forEach((inp, i) => {
                const rec = {
                  index: i,
                  id: inp.id || '',
                  name: inp.name || '',
                  className: inp.className || '',
                  type: inp.type || '',
                  value: inp.value || '',
                  readonly: !!inp.readOnly,
                  visible: inp.offsetParent !== null,
                  outerHTML: inp.outerHTML.slice(0, 500),
                };
                // label ที่เกี่ยวข้อง
                let lbl = null;
                if (inp.id) lbl = document.querySelector('label[for="'+inp.id+'"]');
                if (!lbl) {
                  const par = inp.closest('.form-floating') || inp.parentElement;
                  if (par) lbl = par.querySelector('label');
                }
                rec.labelText = lbl ? (lbl.textContent||'').trim() : '';
                rec.labelHTML = lbl ? lbl.outerHTML.slice(0, 300) : '';
                // parent container
                const par2 = inp.parentElement;
                rec.parentTag = par2 ? par2.tagName : '';
                rec.parentClass = par2 ? par2.className : '';
                rec.parentHTML = par2 ? par2.outerHTML.slice(0, 800) : '';
                // grandparent
                const gp = par2 ? par2.parentElement : null;
                rec.grandparentClass = gp ? gp.className : '';
                // datepicker instance (bootstrap-datepicker)
                if (window.jQuery) {
                  try {
                    const $inp = window.jQuery(inp);
                    rec.jq_datepicker = !!$inp.data('datepicker');
                    rec.jq_flatpickr = !!inp._flatpickr;
                    rec.jq_events = Object.keys($inp.data() || {});
                  } catch(e) { rec.jq_err = String(e); }
                }
                if (inp._flatpickr) rec.flatpickr = true;
                out.inputs.push(rec);
              });

              // ---- 2) หา ปุ่ม 'ค้นหา' ----
              const btns = Array.from(document.querySelectorAll(
                'button, a, input[type=button], input[type=submit]'
              ));
              btns.forEach(b => {
                const txt = ((b.innerText || b.value || '') + '').trim();
                if (/ค้นหา|Search|search|กรอง|filter/i.test(txt)) {
                  out.searchButtons.push({
                    tag: b.tagName,
                    text: txt.slice(0, 60),
                    id: b.id || '',
                    onclick: (b.getAttribute('onclick') || '').slice(0, 200),
                    outerHTML: b.outerHTML.slice(0, 400),
                  });
                }
              });

              // ---- 3) หา code ที่อ่านค่าวันที่ใน GetDataRequestFormEtrackingForAlien ----
              if (typeof GetDataRequestFormEtrackingForAlien === 'function') {
                out.getDataSource = GetDataRequestFormEtrackingForAlien.toString().slice(0, 3000);
              }

              // ---- 4) list hidden inputs ที่อาจเก็บวันที่จริง ----
              const allInputs = Array.from(document.querySelectorAll('input'));
              out.hiddenLikelyDate = allInputs
                .filter(i => {
                  const n = (i.id || i.name || '').toLowerCase();
                  return /date|submit|start|end|from|to/.test(n)
                    && i.type !== 'checkbox' && i.type !== 'radio';
                })
                .slice(0, 30)
                .map(i => ({
                  id: i.id, name: i.name, type: i.type,
                  value: i.value, className: i.className,
                  outerHTML: i.outerHTML.slice(0, 300),
                }));

              return out;
            }"""
        )
        Path("_probe_date_filter_out.json").write_text(
            json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"[✓] dumped → _probe_date_filter_out.json")
        print(f"    inputs found: {len(info.get('inputs', []))}")
        print(f"    search buttons: {len(info.get('searchButtons', []))}")
        print(f"    hidden date-like: {len(info.get('hiddenLikelyDate', []))}")
        for r in info.get("inputs", []):
            print(f"      [{r['index']}] id={r['id']!r} label={r['labelText']!r} "
                  f"value={r['value']!r} readonly={r['readonly']} "
                  f"dp={r.get('jq_datepicker')} fp={r.get('flatpickr', False)}")

        # ---- 5) ทดลอง set ค่า + trigger + ดูว่าตารางเปลี่ยนไหม ----
        print("\n[test] ทดลอง set วันที่ = 01/01/2026 → 15/07/2026 แล้วรอ 5 วิ...")
        test_result = page.evaluate(
            r"""({df, dt}) => {
              function findByLabel(txt) {
                const inputs = Array.from(document.querySelectorAll('input[placeholder="วว/ดด/ปปปป"]'));
                for (const inp of inputs) {
                  let lbl = null;
                  if (inp.id) lbl = document.querySelector('label[for="'+inp.id+'"]');
                  if (!lbl) {
                    const par = inp.closest('.form-floating') || inp.parentElement;
                    if (par) lbl = par.querySelector('label');
                  }
                  if (lbl && (lbl.textContent||'').trim() === txt) return inp;
                }
                return null;
              }
              const inp1 = findByLabel('วันที่ยื่นคำขอ');
              const inp2 = findByLabel('ถึง');
              const log = {inp1_found: !!inp1, inp2_found: !!inp2};
              if (inp1) {
                inp1.value = df;
                inp1.dispatchEvent(new Event('input',{bubbles:true}));
                inp1.dispatchEvent(new Event('change',{bubbles:true}));
                inp1.dispatchEvent(new Event('blur',{bubbles:true}));
                log.inp1_after = inp1.value;
              }
              if (inp2) {
                inp2.value = dt;
                inp2.dispatchEvent(new Event('input',{bubbles:true}));
                inp2.dispatchEvent(new Event('change',{bubbles:true}));
                inp2.dispatchEvent(new Event('blur',{bubbles:true}));
                log.inp2_after = inp2.value;
              }
              return log;
            }""",
            {"df": "01/01/2026", "dt": "15/07/2026"},
        )
        print(f"    set result: {test_result}")
        page.wait_for_timeout(2000)

        # อ่านค่าจริงหลัง set (เผื่อ datepicker เขียนทับกลับ)
        after = page.evaluate(
            r"""() => {
              const inputs = Array.from(document.querySelectorAll('input[placeholder="วว/ดด/ปปปป"]'));
              return inputs.map(i => ({id: i.id, value: i.value}));
            }"""
        )
        print(f"    values after 2s: {after}")

        print("\n[i] เปิด browser ค้างไว้ 30 วิ — ลองกดค้นหาเองแล้วสังเกต network/DOM")
        page.wait_for_timeout(30000)
        ctx.close(); b.close()


if __name__ == "__main__":
    main()
