"""ตรวจ profile หลังจาก login — หาว่ามี profile กี่ตัว + ตัวไหนถูก select ตอนเข้า tracking"""
import os, sys, json
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()

from scrape_wa import login, goto_tracking, apply_wa_filter, _wait_datatable_idle, collect_all_wa_rows  # noqa: E402

cfg = {
    "username": os.getenv("EWP_USERNAME",""),
    "password": os.getenv("EWP_PASSWORD",""),
    "user_type": os.getenv("EWP_USER_TYPE","ผู้กระทำการแทน"),
    "method": os.getenv("EWP_LOGIN_METHOD","E-Workpermit"),
    "headless": False,
}

with sync_playwright() as pw:
    b = pw.chromium.launch(headless=False,
                           args=["--ignore-certificate-errors","--start-maximized"])
    ctx = b.new_context(locale="th-TH", ignore_https_errors=True,
                        viewport={"width":1600,"height":1000})
    page = ctx.new_page()

    ajax_hits = []
    def on_req(req):
        u = req.url or ""
        if "GetDataRequestFormEtracking" in u or "GetProfile" in u or "ChangeProfile" in u:
            try: post = req.post_data or ""
            except Exception: post = ""
            ajax_hits.append({"url": u, "method": req.method, "post": post[:800]})
    page.on("request", on_req)

    login(page, cfg)
    goto_tracking(page)
    page.wait_for_timeout(4000)

    # อ่าน core_datas (global var) + profile list
    info = page.evaluate(r"""() => {
        const out = {};
        try {
            if (typeof core_datas !== 'undefined') {
                out.core_datas = JSON.parse(JSON.stringify(core_datas));
            }
        } catch(e) { out.core_datas_err = String(e); }
        // list profile chooser (ถ้ามีในหน้า)
        const profileEls = document.querySelectorAll('[onclick*="ChangeProfile"], .profile-item, [data-profile]');
        out.profileButtons = [];
        profileEls.forEach(el => {
            out.profileButtons.push({
                tag: el.tagName,
                text: (el.innerText||'').trim().slice(0, 100),
                onclick: (el.getAttribute('onclick') || '').slice(0, 200),
                dataProfile: el.getAttribute('data-profile') || '',
            });
        });
        // ค่า Filter_request_submit options (submitters)
        const sel = document.getElementById('Filter_request_submit');
        if (sel) {
            out.filter_submit_options = Array.from(sel.options).map(o => ({v: o.value, t: (o.textContent||'').trim().slice(0, 80)}));
            out.filter_submit_selected = sel.value;
        }
        return out;
    }""")

    cd = info.get("core_datas", {}) or {}
    print("=== core_datas ===")
    print(f"  user_id: {cd.get('user_id')}")
    cur = cd.get('current_profile') or {}
    print(f"  current_profile.profile_id: {cur.get('profile_id')}")
    print(f"  current_profile.name: {cur.get('firstname_th')} {cur.get('lastname_th')}  |  {cur.get('company_th')}")
    profiles = cd.get('profiles') or []
    print(f"\n  profiles ({len(profiles)}):")
    for p in profiles:
        print(f"    id={p.get('profile_id')} type={p.get('profile_type_id')} name={p.get('firstname_th','') or p.get('company_th','')}")

    print(f"\n  Filter_request_submit options: {info.get('filter_submit_options')}")
    print(f"  Filter_request_submit selected: {info.get('filter_submit_selected')!r}")

    # เก็บทั้งหมดใน json
    Path("_probe_profile.json").write_text(
        json.dumps({"info": info, "ajax": ajax_hits}, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    # ทดสอบ: filter WP + count + list first row
    print("\n=== filter=WP ===")
    apply_wa_filter(page, "", status_ids=["WP"])
    _wait_datatable_idle(page, max_wait_ms=30000)
    total = page.evaluate("() => window.jQuery('#datatableE_Tracking').DataTable().page.info().recordsTotal")
    print(f"  recordsTotal: {total}")
    rows = collect_all_wa_rows(page)
    print(f"  collected: {len(rows)}")
    from collections import Counter
    counts = Counter(str(r.get("status") or "?").upper() for r in rows)
    print(f"  sub-status: {dict(counts)}")

    page.wait_for_timeout(5000)
    ctx.close(); b.close()
