"""READ-ONLY probe step 3: capture #WP_Payment modal static DOM + JS sources of
sendPaymentStatus/createdatapaymentgroup + any print-bill function.

Does NOT call OpentWP()/sendPaymentStatus — only reads the hidden modal HTML and fn sources
so we can confirm the next step just generates a bill (no immediate charge).
"""
import json, re
from playwright.sync_api import sync_playwright

from scrape_wa import (
    load_config, login, goto_tracking, apply_wa_filter,
    collect_all_wa_rows, build_detail_url,
)

OUT = {}
cfg = load_config(require_login=True)
print(f"[probe3] login as {cfg['username']}")

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
    url = build_detail_url(target)
    print(f"[probe3] detail: {url}")
    page.goto(url, wait_until="domcontentloaded", timeout=40_000)
    page.wait_for_timeout(4000)
    try:
        page.locator("#sltLang").first.select_option(value="th"); page.wait_for_timeout(1000)
    except Exception:
        pass
    page.evaluate(r"""() => { if(window.jQuery){try{jQuery('.modal').modal('hide');}catch(e){}}
      document.querySelectorAll('.modal.show .btn-close,.modal.show [data-bs-dismiss],.modal.show [data-dismiss]').forEach(b=>{try{b.click();}catch(e){}}); }""")
    page.wait_for_timeout(500)
    page.evaluate(r"""() => { const a=Array.from(document.querySelectorAll('a[href^="#"]')).find(e=>/การชำระเงิน/.test((e.textContent||'').trim())); if(a)a.click(); }""")
    page.wait_for_timeout(1500)

    # static HTML of the WP_Payment modal (hidden in DOM)
    OUT["WP_Payment_html"] = page.evaluate(
        r"""() => { const m=document.getElementById('WP_Payment'); return m ? m.outerHTML.slice(0,6000) : '(#WP_Payment not found)'; }"""
    )
    # list all modal ids present
    OUT["all_modal_ids"] = page.evaluate(
        r"""() => Array.from(document.querySelectorAll('.modal')).map(m=>m.id||'(noid)')"""
    )
    print("[probe3] modal ids:", OUT["all_modal_ids"])

    # capture broad fn sources (payment + print/bill)
    OUT["fn_src"] = page.evaluate(
        r"""() => {
          const want = ['sendPaymentStatus','createdatapaymentgroup','createdatapaymentSingle',
            'OpentWP','printBillPayment','PrintBillPayment','printPayment','PrintPayment',
            'printFormPayment','PrintFormPayment','printpayment','genPdfPayment','downloadBill',
            'OpenBillPayment','printbill','Printbill'];
          const out = {};
          want.forEach(n => { try { if (typeof window[n]==='function') out[n]=window[n].toString().slice(0,3000); } catch(e){} });
          // also scan all global functions whose name matches print/bill/pay+pdf
          for (const k in window) {
            try {
              if (typeof window[k]==='function' && /print|bill|pdf|receipt/i.test(k) && !(k in out)) {
                out[k] = window[k].toString().slice(0,1500);
              }
            } catch(e){}
          }
          return out;
        }"""
    )
    print("[probe3] fn sources captured:", list(OUT["fn_src"].keys()))

    # buttons inside WP_Payment (even if hidden) — text + onclick
    OUT["WP_Payment_buttons"] = page.evaluate(
        r"""() => { const m=document.getElementById('WP_Payment'); if(!m) return [];
          return Array.from(m.querySelectorAll('button,a,input[type=button],input[type=submit]'))
            .map(e=>({text:(e.textContent||e.value||'').trim().replace(/\s+/g,' ').slice(0,60), onclick:(e.getAttribute('onclick')||'').slice(0,200), id:e.id||'', cls:(e.className||'').slice(0,90)}))
            .filter(b=>b.text||b.onclick||b.id); }"""
    )
    print(f"[probe3] WP_Payment buttons ({len(OUT['WP_Payment_buttons'])}):")
    for bn in OUT["WP_Payment_buttons"]:
        print(f"   BTN {bn['text']!r} id={bn['id']} onclick={bn['onclick']!r} cls={bn['cls']!r}")

    with open("_probe_billpay3_out.json", "w", encoding="utf-8") as f:
        json.dump(OUT, f, ensure_ascii=False, indent=2)
    print("[probe3] wrote _probe_billpay3_out.json — browser open 8s (NO payment triggered)")
    page.wait_for_timeout(8_000)
