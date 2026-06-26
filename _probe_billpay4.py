"""READ-ONLY: capture UpdatePayment/startConnection/createdatapaymentgroup/backtoPayment_Modal sources."""
import json
from playwright.sync_api import sync_playwright
from scrape_wa import (load_config, login, goto_tracking, apply_wa_filter,
                       collect_all_wa_rows, build_detail_url)

cfg = load_config(require_login=True)
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
    page.goto(build_detail_url(target), wait_until="domcontentloaded", timeout=40_000)
    page.wait_for_timeout(3500)
    try:
        page.locator("#sltLang").first.select_option(value="th"); page.wait_for_timeout(800)
    except Exception:
        pass
    src = page.evaluate(
        r"""() => {
          const want = ['UpdatePayment','startConnection','createdatapaymentgroup','createdatapaymentSingle',
            'backtoPayment_Modal','show_loading','hide_loading'];
          const out = {};
          want.forEach(n => { try { if (typeof window[n]==='function') out[n]=window[n].toString().slice(0,3500); } catch(e){} });
          return out;
        }"""
    )
    for k, v in src.items():
        print(f"\n===== {k} =====\n{v}")
    with open("_probe_billpay4_out.json", "w", encoding="utf-8") as f:
        json.dump(src, f, ensure_ascii=False, indent=2)
    print("\n[probe4] wrote _probe_billpay4_out.json")
