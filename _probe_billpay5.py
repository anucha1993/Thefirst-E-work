"""LIVE download probe: run the full bill-payment flow on ONE request, download the
invoice PDF (pay-later bill — NO charge), dump its text + parsed fields to design naming.
"""
import io, json
from playwright.sync_api import sync_playwright
from scrape_wa import (load_config, login, goto_tracking, apply_wa_filter,
                       collect_all_wa_rows, build_detail_url,
                       _grab_pdf_after_click, _bt30_parse_payment_pdf)

cfg = load_config(require_login=True)
print(f"[probe5] login as {cfg['username']}")

with sync_playwright() as pw:
    b = pw.chromium.launch(headless=False, args=["--ignore-certificate-errors"])
    ctx = b.new_context(locale="th-TH", ignore_https_errors=True)
    page = ctx.new_page()
    login(page, cfg)
    goto_tracking(page)
    apply_wa_filter(page, "", status_ids=["WP"])
    page.wait_for_timeout(1000)
    rows = collect_all_wa_rows(page, log=print)
    target = next((r for r in rows if "รอจ่ายเงินค่าธรรมเนียม" in (r.get("statusText", "") or "")), None)
    print(f"[probe5] target group_id={target.get('group_id')} reqNo={target.get('reqNo')} requester={target.get('requester')}")
    page.goto(build_detail_url(target), wait_until="domcontentloaded", timeout=40_000)
    page.wait_for_timeout(3500)
    try:
        page.locator("#sltLang").first.select_option(value="th"); page.wait_for_timeout(900)
    except Exception:
        pass
    page.evaluate(r"""() => { if(window.jQuery){try{jQuery('.modal').modal('hide');}catch(e){}}
      document.querySelectorAll('.modal.show .btn-close,.modal.show [data-bs-dismiss],.modal.show [data-dismiss]').forEach(b=>{try{b.click();}catch(e){}}); }""")
    page.wait_for_timeout(500)

    # also capture alien name from ข้อมูลคนต่างด้าว tab BEFORE paying
    page.evaluate(r"""() => { const a=Array.from(document.querySelectorAll('a[href^="#"]')).find(e=>/ข้อมูลคนต่างด้าว/.test((e.textContent||'').trim())); if(a)a.click(); }""")
    page.wait_for_timeout(1500)
    alien_tab = page.evaluate(
        r"""() => { const el=document.getElementById('tab_default_2'); return el ? (el.innerText||'').replace(/\s+/g,' ').slice(0,800) : ''; }"""
    )
    print(f"[probe5] alien tab text: {alien_tab[:400]!r}")

    # go to payment tab
    page.evaluate(r"""() => { const a=Array.from(document.querySelectorAll('a[href^="#"]')).find(e=>/การชำระเงิน/.test((e.textContent||'').trim())); if(a)a.click(); }""")
    page.wait_for_timeout(1500)

    # 1) OpenModalPay1
    page.evaluate(r"""() => { const b=Array.from(document.querySelectorAll('button,a')).find(e=>/OpenModalPay1/.test(e.getAttribute('onclick')||'')); if(b)b.click(); }""")
    page.wait_for_selector("#exampleModal.show", timeout=8000)
    page.wait_for_timeout(1200)
    # 2) inner ชำระเงิน -> OpentWP
    page.evaluate(r"""() => { const b=document.getElementById('openpaymentdetail'); if(b)b.click(); }""")
    page.wait_for_selector("#WP_Payment.show", timeout=10000)
    page.wait_for_timeout(1800)
    refs = page.evaluate(
        r"""() => ({
          ref1:(document.getElementById('ref1')||{}).innerText||'',
          ref2:(document.getElementById('ref2')||{}).innerText||'',
          amount:(document.getElementById('lbl_price_payment_amount')||{}).innerText||'',
          due:(document.getElementById('payment_expired_time')||{}).innerText||'',
        })"""
    )
    print(f"[probe5] WP_Payment refs: {refs}")

    # 3) print bill -> capture popup PDF
    def _do_click():
        page.evaluate(r"""() => { const b=document.getElementById('button_payment'); if(b)b.click(); }""")

    body, err = _grab_pdf_after_click(page, _do_click, log=print)
    print(f"[probe5] pdf bytes={len(body) if body else 0} err={err!r}")
    if body and body[:5] == b"%PDF-":
        with open("_probe_billpay_invoice.pdf", "wb") as f:
            f.write(body)
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(body))
        full = "\n".join((pg.extract_text() or "") for pg in reader.pages)
        print(f"[probe5] PDF pages={len(reader.pages)}")
        print("===== PDF TEXT (per page) =====")
        for i, pg in enumerate(reader.pages):
            print(f"\n----- page {i+1} -----")
            print((pg.extract_text() or "")[:2500])
        parsed = _bt30_parse_payment_pdf(body)
        print("\n[probe5] _bt30_parse_payment_pdf =>", json.dumps(parsed, ensure_ascii=False))
        with open("_probe_billpay5_out.json", "w", encoding="utf-8") as f:
            json.dump({"target": {k: target.get(k, "") for k in ("group_id","reqNo","requester","statusText","user_id","form_type")},
                       "alien_tab": alien_tab, "refs": refs, "parsed": parsed, "pdf_text": full[:6000]}, f, ensure_ascii=False, indent=2)
        print("[probe5] wrote _probe_billpay5_out.json + _probe_billpay_invoice.pdf")
    else:
        print("[probe5] FAILED to get PDF")
    page.wait_for_timeout(4000)
