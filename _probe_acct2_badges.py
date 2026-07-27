"""ตรวจสถานะทั้งหมดของบัญชี 2 (1509900662001)"""
import sys, tempfile
from pathlib import Path
from openpyxl import Workbook, load_workbook
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
from scrape_wa import login, goto_tracking, apply_wa_filter, _wait_datatable_idle  # noqa: E402

wb0 = load_workbook("UsernameLogin.xlsx", data_only=True)
ws0 = wb0.active
row2 = [c.value for c in ws0[2]]
cfg = {"username": str(row2[0]), "password": str(row2[1]),
       "user_type": str(row2[2]), "method": str(row2[3]),
       "headless": False}

with sync_playwright() as pw:
    b = pw.chromium.launch(headless=False,
                           args=["--ignore-certificate-errors", "--start-maximized"])
    ctx = b.new_context(locale="th-TH", ignore_https_errors=True,
                        viewport={"width":1600,"height":1000})
    page = ctx.new_page()
    login(page, cfg)
    goto_tracking(page)
    page.wait_for_timeout(3500)

    # อ่าน badge ทั้งหมด
    badges = page.evaluate(r"""() => {
        const out = {};
        ['WP','WCOSNA','WA','AP','SS'].forEach(sid => {
            const el = document.getElementById('number_' + sid);
            out[sid] = el ? el.textContent.trim() : '(missing)';
        });
        const total = document.getElementById('numberform');
        out['TOTAL'] = total ? total.textContent.trim() : '(missing)';
        return out;
    }""")
    print(f"Badges: {badges}")

    # count each status
    for s in ["WP","WCOSNA","WA","AP","SS"]:
        apply_wa_filter(page, "", status_ids=[s])
        _wait_datatable_idle(page, max_wait_ms=15000)
        n = page.evaluate("() => window.jQuery('#datatableE_Tracking').DataTable().page.info().recordsTotal")
        print(f"  {s:8s} → {n}")

    page.wait_for_timeout(8000)
    ctx.close(); b.close()
