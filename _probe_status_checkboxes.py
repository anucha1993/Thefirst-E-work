"""Probe: หา checkbox filter status ทั้งหมดในหน้า e-Tracking (WP, WP2, ...)"""
from __future__ import annotations
import os
import sys
import json
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()

from scrape_wa import login, goto_tracking, apply_wa_filter, _wait_datatable_idle  # noqa: E402


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
        "headless": False,
    }
    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False,
                               args=["--ignore-certificate-errors", "--start-maximized"])
        ctx = b.new_context(locale="th-TH", ignore_https_errors=True,
                            viewport={"width": 1600, "height": 1000})
        page = ctx.new_page()
        login(page, cfg)
        goto_tracking(page)
        page.wait_for_timeout(2500)

        # ---- 1) หา checkbox ทุกตัว + label ----
        info = page.evaluate(
            r"""() => {
              const out = {checkboxes: [], numberCounters: []};
              document.querySelectorAll('input[type="checkbox"]').forEach(cb => {
                const id = cb.id || '';
                if (!id) return;
                // label associate
                let lbl = document.querySelector('label[for="'+id+'"]');
                if (!lbl && cb.parentElement) lbl = cb.parentElement.querySelector('label') || cb.closest('label');
                const par = cb.parentElement;
                out.checkboxes.push({
                  id, name: cb.name || '', value: cb.value || '',
                  checked: cb.checked,
                  visible: cb.offsetParent !== null,
                  labelText: lbl ? (lbl.textContent||'').trim().replace(/\s+/g,' ').slice(0,120) : '',
                  parentText: par ? (par.textContent||'').trim().replace(/\s+/g,' ').slice(0,120) : '',
                });
              });
              // เก็บ counter span ที่ show 'number_XXX'
              document.querySelectorAll('[id^="number_"]').forEach(el => {
                out.numberCounters.push({
                  id: el.id,
                  text: (el.textContent||'').trim(),
                  parentText: (el.parentElement ? (el.parentElement.textContent||'').replace(/\s+/g,' ').trim().slice(0, 120) : ''),
                });
              });
              return out;
            }"""
        )
        Path("_probe_status_checkboxes.json").write_text(
            json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print("=== Filter Checkboxes ===")
        for cb in info["checkboxes"]:
            if not cb["visible"]:
                continue
            print(f"  id={cb['id']:15s} label={cb['labelText'] or cb['parentText']!r}")
        print("\n=== Number counters (badge with counts) ===")
        for c in info["numberCounters"]:
            print(f"  {c['id']:20s} = {c['text']!r}  | ctx: {c['parentText']}")

        # ---- 2) ทดสอบ apply_wa_filter ด้วย status_ids=['WP2'] ----
        print("\n=== Test filter WP2 ===")
        apply_wa_filter(page, "", status_ids=["WP2"])
        _wait_datatable_idle(page, max_wait_ms=15000)
        total = page.evaluate(
            "() => { try { return window.jQuery('#datatableE_Tracking').DataTable().page.info().recordsTotal; } catch(e) { return -1; } }"
        )
        print(f"  WP2 → {total} รายการ")

        # ---- 3) ทดสอบด้วย NCOS, APSS, CCOS ----
        for sid in ["NCOS", "APSS", "CCOS"]:
            apply_wa_filter(page, "", status_ids=[sid])
            _wait_datatable_idle(page, max_wait_ms=15000)
            total = page.evaluate(
                "() => { try { return window.jQuery('#datatableE_Tracking').DataTable().page.info().recordsTotal; } catch(e) { return -1; } }"
            )
            print(f"  {sid:8s} → {total} รายการ")

        page.wait_for_timeout(15000)
        ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
