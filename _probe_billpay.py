"""READ-ONLY probe for the new Bill Payment mode.

Steps (no irreversible action — does NOT click 'ชำระเงิน'):
1. login (.env account)
2. goto /Permit/Tracking
3. filter status = WP (รอชำระเงิน)
4. collect rows, dump statusText to find 'รอจ่ายเงินค่าธรรมเนียม'
5. open first matching detail page, click tab 'การชำระเงิน',
   DUMP the payment-section DOM (sections, buttons, onclick) WITHOUT paying.
"""
import json
from playwright.sync_api import sync_playwright

from scrape_wa import (
    load_config, login, goto_tracking, apply_wa_filter,
    collect_all_wa_rows, build_detail_url,
)

OUT = {}

cfg = load_config(require_login=True)
print(f"[probe] login as {cfg['username']} / {cfg['user_type']} / {cfg['method']}")

with sync_playwright() as pw:
    b = pw.chromium.launch(headless=False, args=["--ignore-certificate-errors"])
    ctx = b.new_context(locale="th-TH", ignore_https_errors=True)
    page = ctx.new_page()

    login(page, cfg)
    goto_tracking(page)
    apply_wa_filter(page, "", status_ids=["WP"])
    page.wait_for_timeout(1500)

    rows = collect_all_wa_rows(page, log=print)
    print(f"[probe] total WP rows = {len(rows)}")
    OUT["rows"] = [
        {k: r.get(k, "") for k in
         ("group_id", "user_id", "status", "form_type", "institution_id", "id",
          "reqNo", "requester", "desc", "statusText")}
        for r in rows
    ]
    # show distinct statusText values
    seen_status = {}
    for r in rows:
        st = r.get("statusText", "")
        seen_status[st] = seen_status.get(st, 0) + 1
    OUT["statusText_counts"] = seen_status
    print("[probe] statusText counts:")
    for st, n in seen_status.items():
        print(f"   {n:>3}  {st!r}")

    # pick first row that mentions รอจ่ายเงินค่าธรรมเนียม, else first WP row
    target = None
    for r in rows:
        if "รอจ่ายเงินค่าธรรมเนียม" in (r.get("statusText", "") or ""):
            target = r
            break
    if target is None and rows:
        target = rows[0]

    if target:
        url = build_detail_url(target)
        OUT["target_row"] = {k: target.get(k, "") for k in
                             ("group_id", "user_id", "status", "form_type", "statusText", "reqNo")}
        OUT["target_url"] = url
        print(f"[probe] opening detail: {url}")
        page.goto(url, wait_until="domcontentloaded", timeout=40_000)
        page.wait_for_timeout(4000)
        # set Thai + dismiss any news modal
        try:
            page.locator("#sltLang").first.select_option(value="th")
            page.wait_for_timeout(1500)
        except Exception:
            pass
        page.evaluate(
            r"""() => {
              // close any visible bootstrap modal / news popup
              if (window.jQuery) { try { jQuery('.modal').modal('hide'); } catch(e){} }
              document.querySelectorAll('.modal.show .btn-close, .modal.show [data-bs-dismiss], .modal.show [data-dismiss]')
                .forEach(b => { try { b.click(); } catch(e){} });
            }"""
        )
        page.wait_for_timeout(800)

        # dump tabs
        tabs = page.evaluate(
            r"""() => {
              const vis = e => { const r = e.getBoundingClientRect(); return r.width>0 && r.height>0; };
              return Array.from(document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a, button[data-bs-toggle="tab"]'))
                .filter(vis)
                .map(a => ({text:(a.textContent||'').trim().slice(0,60), href:a.getAttribute('href')||'', onclick:(a.getAttribute('onclick')||'').slice(0,120), id:a.id||''}))
                .filter(t => t.text);
            }"""
        )
        OUT["tabs"] = tabs
        print(f"[probe] tabs found: {len(tabs)}")
        for t in tabs:
            print(f"   tab: {t['text']!r}  href={t['href']}  id={t['id']}  onclick={t['onclick']}")

        # click the การชำระเงิน tab
        clicked = page.evaluate(
            r"""() => {
              const vis = e => { const r = e.getBoundingClientRect(); return r.width>0 && r.height>0; };
              const a = Array.from(document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a, button[data-bs-toggle="tab"]'))
                .find(e => vis(e) && /การชำระเงิน|ชำระเงิน/.test((e.textContent||'').trim()));
              if (a) { a.click(); return (a.textContent||'').trim(); }
              return '';
            }"""
        )
        print(f"[probe] clicked payment tab = {clicked!r}")
        page.wait_for_timeout(2500)

        # dump the payment section: headings, buttons with text/onclick, section text
        pay = page.evaluate(
            r"""() => {
              const vis = e => { const r = e.getBoundingClientRect(); const s=getComputedStyle(e);
                return r.width>0 && r.height>0 && s.display!=='none' && s.visibility!=='hidden'; };
              const out = {};
              out.headings = Array.from(document.querySelectorAll('h1,h2,h3,h4,h5,h6,legend,.card-title,.card-header,b,strong'))
                .filter(vis).map(e => (e.textContent||'').trim().replace(/\s+/g,' ').slice(0,120)).filter(t=>t).slice(0,60);
              out.buttons = Array.from(document.querySelectorAll('button, a.btn, a[onclick], input[type="button"], input[type="submit"]'))
                .filter(vis)
                .map(e => ({text:(e.textContent||e.value||'').trim().replace(/\s+/g,' ').slice(0,60), onclick:(e.getAttribute('onclick')||'').slice(0,200), id:e.id||'', cls:(e.className||'').slice(0,80)}))
                .filter(b => b.text || b.onclick);
              // any element whose text mentions ค่าธรรมเนียมขออนุญาตทำงาน
              const fee = [];
              document.querySelectorAll('*').forEach(e => {
                const t=(e.childElementCount===0? (e.textContent||''):'' ).trim();
                if (/ค่าธรรมเนียมขออนุญาตทำงาน/.test(t)) fee.push(t.slice(0,120));
              });
              out.fee_labels = Array.from(new Set(fee)).slice(0,10);
              return out;
            }"""
        )
        OUT["payment_section"] = pay
        print(f"[probe] payment headings ({len(pay['headings'])}):")
        for h in pay["headings"]:
            print("   H:", h)
        print(f"[probe] buttons in payment view ({len(pay['buttons'])}):")
        for bn in pay["buttons"]:
            print(f"   BTN: {bn['text']!r}  id={bn['id']}  onclick={bn['onclick']!r}")
        print(f"[probe] fee labels: {pay['fee_labels']}")

        page.screenshot(path="_probe_billpay_payment.png", full_page=True)

    with open("_probe_billpay_out.json", "w", encoding="utf-8") as f:
        json.dump(OUT, f, ensure_ascii=False, indent=2)
    print("[probe] wrote _probe_billpay_out.json — leaving browser open 10s")
    page.wait_for_timeout(10_000)
