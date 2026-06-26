"""READ-ONLY probe step 2: open Modal 1 (OpenModalPay1) and capture its DOM + JS sources.

Does NOT click the inner 'ชำระเงิน' — only opens the first modal and dumps it so we
can understand whether the next click commits a payment or just generates a bill.
"""
import json
from playwright.sync_api import sync_playwright

from scrape_wa import (
    load_config, login, goto_tracking, apply_wa_filter,
    collect_all_wa_rows, build_detail_url,
)

OUT = {}
cfg = load_config(require_login=True)
print(f"[probe2] login as {cfg['username']}")

with sync_playwright() as pw:
    b = pw.chromium.launch(headless=False, args=["--ignore-certificate-errors"])
    ctx = b.new_context(locale="th-TH", ignore_https_errors=True)
    page = ctx.new_page()

    login(page, cfg)
    goto_tracking(page)
    apply_wa_filter(page, "", status_ids=["WP"])
    page.wait_for_timeout(1200)
    rows = collect_all_wa_rows(page, log=print)
    target = next((r for r in rows if "รอจ่ายเงินค่าธรรมเนียม" in (r.get("statusText", "") or "")), None)
    if not target:
        print("[probe2] no รอจ่ายเงินค่าธรรมเนียม row found")
        raise SystemExit(0)
    url = build_detail_url(target)
    OUT["target_url"] = url
    print(f"[probe2] detail: {url}")
    page.goto(url, wait_until="domcontentloaded", timeout=40_000)
    page.wait_for_timeout(4000)
    try:
        page.locator("#sltLang").first.select_option(value="th")
        page.wait_for_timeout(1200)
    except Exception:
        pass
    # dismiss news modal if any
    page.evaluate(r"""() => { if (window.jQuery){ try{ jQuery('.modal').modal('hide'); }catch(e){} }
      document.querySelectorAll('.modal.show .btn-close,.modal.show [data-bs-dismiss],.modal.show [data-dismiss]').forEach(b=>{try{b.click();}catch(e){}}); }""")
    page.wait_for_timeout(600)
    # click การชำระเงิน tab
    page.evaluate(r"""() => { const a=Array.from(document.querySelectorAll('a[href^="#"]')).find(e=>/การชำระเงิน/.test((e.textContent||'').trim())); if(a)a.click(); }""")
    page.wait_for_timeout(2000)

    # capture JS source of payment fns
    OUT["fn_src"] = page.evaluate(
        r"""() => {
          const names = ['OpenModalPay1','OpenModalPay2','OpenModalPay3','SavePay','ConfirmPay',
            'PrintPayment','PrintFormPay','Openproofpayment','CreateBillPayment','PayFee','GenBillPayment'];
          const out = {};
          names.forEach(n => { try { if (typeof window[n]==='function') out[n]=window[n].toString().slice(0,2500); } catch(e){} });
          return out;
        }"""
    )
    print("[probe2] captured fn sources:", list(OUT["fn_src"].keys()))

    # click ชำระเงิน (OpenModalPay1) — opens Modal 1 only
    clicked = page.evaluate(
        r"""() => {
          const vis = e => { const r=e.getBoundingClientRect(); return r.width>0&&r.height>0; };
          const b = Array.from(document.querySelectorAll('button,a'))
            .find(e => vis(e) && /OpenModalPay1/.test(e.getAttribute('onclick')||''));
          if (b) { b.click(); return (b.getAttribute('onclick')||''); }
          return '';
        }"""
    )
    print(f"[probe2] clicked Pay btn onclick={clicked!r}")
    page.wait_for_timeout(3000)

    # dump all visible modals
    OUT["modals"] = page.evaluate(
        r"""() => {
          const vis = e => { const r=e.getBoundingClientRect(); const s=getComputedStyle(e);
            return r.width>0&&r.height>0&&s.display!=='none'&&s.visibility!=='hidden'; };
          return Array.from(document.querySelectorAll('.modal'))
            .filter(m => m.classList.contains('show') || vis(m))
            .map(m => ({
              id: m.id||'',
              title: (m.querySelector('.modal-title,.modal-header')||{}).innerText ? (m.querySelector('.modal-title,.modal-header').innerText||'').trim().slice(0,120):'',
              body_text: (m.querySelector('.modal-body')||m).innerText ? ((m.querySelector('.modal-body')||m).innerText||'').trim().replace(/\s+/g,' ').slice(0,600):'',
              buttons: Array.from(m.querySelectorAll('button,a.btn,a[onclick],input[type=button],input[type=submit]'))
                .filter(vis).map(e => ({text:(e.textContent||e.value||'').trim().replace(/\s+/g,' ').slice(0,60), onclick:(e.getAttribute('onclick')||'').slice(0,200), id:e.id||'', cls:(e.className||'').slice(0,80)}))
            }));
        }"""
    )
    print(f"[probe2] visible modals: {len(OUT['modals'])}")
    for m in OUT["modals"]:
        print(f"  MODAL id={m['id']} title={m['title']!r}")
        print(f"    body: {m['body_text'][:300]!r}")
        for bn in m["buttons"]:
            print(f"    BTN {bn['text']!r} id={bn['id']} onclick={bn['onclick']!r}")

    # capture source of any fn referenced by modal buttons
    extra = page.evaluate(
        r"""() => {
          const names = new Set();
          document.querySelectorAll('.modal.show button,.modal.show a').forEach(e=>{
            const oc=e.getAttribute('onclick')||''; const m=oc.match(/([A-Za-z_$][\w$]*)\s*\(/g);
            if(m) m.forEach(x=>names.add(x.replace(/\s*\($/,'')));
          });
          const out={};
          names.forEach(n=>{ try{ if(typeof window[n]==='function') out[n]=window[n].toString().slice(0,2500);}catch(e){} });
          return out;
        }"""
    )
    OUT["modal_btn_fn_src"] = extra
    print("[probe2] modal btn fn sources:", list(extra.keys()))

    page.screenshot(path="_probe_billpay_modal1.png", full_page=True)
    with open("_probe_billpay2_out.json", "w", encoding="utf-8") as f:
        json.dump(OUT, f, ensure_ascii=False, indent=2)
    print("[probe2] wrote _probe_billpay2_out.json — browser open 12s (NOT clicking inner ชำระเงิน)")
    page.wait_for_timeout(12_000)
