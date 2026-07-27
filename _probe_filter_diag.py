"""Diagnose: check if apply_wa_filter actually leaves WP checkbox checked"""
import os, sys, json
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()

from scrape_wa import login, goto_tracking, apply_wa_filter, _wait_datatable_idle  # noqa: E402

cfg = {
    "username": os.getenv("EWP_USERNAME",""), "password": os.getenv("EWP_PASSWORD",""),
    "user_type": os.getenv("EWP_USER_TYPE","ผู้กระทำการแทน"),
    "method": os.getenv("EWP_LOGIN_METHOD","E-Workpermit"),
    "headless": False,
}

with sync_playwright() as pw:
    b = pw.chromium.launch(headless=False,
                           args=["--ignore-certificate-errors","--start-maximized"])
    ctx = b.new_context(locale="th-TH", ignore_https_errors=True, viewport={"width":1600,"height":1000})
    page = ctx.new_page()

    # capture ajax
    ajax = []
    def on_req(req):
        u = req.url or ""
        if "url_GetDataRequestFormEtracking" in u or "GetTracking" in u or "GetData" in u:
            try: post = req.post_data or ""
            except Exception: post = ""
            ajax.append({"url": u, "method": req.method, "post_len": len(post),
                         "post_data": post[:800]})
    page.on("request", on_req)

    login(page, cfg)
    goto_tracking(page)
    page.wait_for_timeout(3000)

    # ทดสอบ: ติ๊ก WP อย่างเดียว
    print("\n=== 1) apply_wa_filter(WP) ===")
    ajax.clear()
    apply_wa_filter(page, "", status_ids=["WP"])
    _wait_datatable_idle(page, max_wait_ms=15000)
    state = page.evaluate(r"""() => {
        const out = {};
        ['WP','WCOSNA','WA','AP','SS'].forEach(id => {
            const cb = document.getElementById(id);
            out[id] = cb ? {checked: cb.checked, value: cb.value, name: cb.name} : null;
        });
        // active jQ:checked selector
        try {
            const $ = window.jQuery;
            out.jq_checked = $("[name = 'checkbox_formstatus_alien']:checked").map(function(){return $(this).val();}).get().toString();
        } catch(e) { out.jq_err = String(e); }
        out.records = window.jQuery('#datatableE_Tracking').DataTable().page.info().recordsTotal;
        return out;
    }""")
    print("  checkbox state:")
    for k in ['WP','WCOSNA','WA','AP','SS']:
        print(f"    {k}: {state.get(k)}")
    print(f"  jq_checked (submitted to server): {state.get('jq_checked')!r}")
    print(f"  records: {state.get('records')}")
    print(f"  ajax calls: {len(ajax)}")
    for a in ajax:
        print(f"    · {a['method']} {a['url'][:120]}")
        if a['post_data']:
            # ดึงเฉพาะ status_id/form_status_id
            import re
            m = re.search(r"status_id[^&]*", a['post_data'])
            print(f"      status_id: {m.group() if m else '(not found)'}")
            m2 = re.search(r"search_start_date[^&]*", a['post_data'])
            m3 = re.search(r"search_end_date[^&]*", a['post_data'])
            print(f"      date: {m2.group() if m2 else ''} / {m3.group() if m3 else ''}")

    print("\n=== 2) apply_wa_filter(WA) ===")
    ajax.clear()
    apply_wa_filter(page, "", status_ids=["WA"])
    _wait_datatable_idle(page, max_wait_ms=15000)
    state = page.evaluate(r"""() => {
        const $ = window.jQuery;
        return {
            jq_checked: $("[name = 'checkbox_formstatus_alien']:checked").map(function(){return $(this).val();}).get().toString(),
            records: $('#datatableE_Tracking').DataTable().page.info().recordsTotal,
        };
    }""")
    print(f"  jq_checked: {state.get('jq_checked')!r}")
    print(f"  records: {state.get('records')}")

    page.wait_for_timeout(10000)
    ctx.close(); b.close()
