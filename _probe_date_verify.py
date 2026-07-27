"""Test: verify apply_wa_filter with date_from/date_to actually works.
รัน: python _probe_date_verify.py
"""
from __future__ import annotations
import os
import sys
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()

from scrape_wa import login, goto_tracking, apply_wa_filter, _wait_datatable_idle  # noqa: E402


def main() -> None:
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
        print(f"[1] login as {cfg['username']!r}")
        login(page, cfg)
        goto_tracking(page)
        page.wait_for_timeout(2500)

        # ---- CASE A: ไม่กรองวันที่ + ทุก status ----
        print("\n[A] no date + all statuses")
        apply_wa_filter(page, "", status_ids=["WP", "WCOSNA", "WA", "AP", "SS"])
        _wait_datatable_idle(page, max_wait_ms=15000)
        totalA = page.evaluate(
            "() => window.jQuery('#datatableE_Tracking').DataTable().page.info().recordsTotal"
        )
        print(f"    → recordsTotal = {totalA}")
        valsA = page.evaluate(
            "() => ({s: document.getElementById('search_start_date').value, "
            "e: document.getElementById('search_end_date').value})"
        )
        print(f"    → date inputs: {valsA}")

        # ---- CASE B: กรองวันที่ 01/01/2026 → 30/06/2026 ----
        print("\n[B] date 01/01/2026 → 30/06/2026 + all statuses")
        apply_wa_filter(
            page, "", status_ids=["WP", "WCOSNA", "WA", "AP", "SS"],
            date_from="01/01/2026", date_to="30/06/2026",
        )
        _wait_datatable_idle(page, max_wait_ms=15000)
        totalB = page.evaluate(
            "() => window.jQuery('#datatableE_Tracking').DataTable().page.info().recordsTotal"
        )
        print(f"    → recordsTotal = {totalB}")
        valsB = page.evaluate(
            "() => ({s: document.getElementById('search_start_date').value, "
            "e: document.getElementById('search_end_date').value})"
        )
        print(f"    → date inputs: {valsB}")

        # ---- CASE C: กรองแคบ 01/07/2026 → 15/07/2026 (แค่ 2 อาทิตย์) ----
        print("\n[C] date 01/07/2026 → 15/07/2026")
        apply_wa_filter(
            page, "", status_ids=["WP", "WCOSNA", "WA", "AP", "SS"],
            date_from="01/07/2026", date_to="15/07/2026",
        )
        _wait_datatable_idle(page, max_wait_ms=15000)
        totalC = page.evaluate(
            "() => window.jQuery('#datatableE_Tracking').DataTable().page.info().recordsTotal"
        )
        print(f"    → recordsTotal = {totalC}")
        valsC = page.evaluate(
            "() => ({s: document.getElementById('search_start_date').value, "
            "e: document.getElementById('search_end_date').value})"
        )
        print(f"    → date inputs: {valsC}")

        print("\n[summary]")
        print(f"  A (no date)         → {totalA}")
        print(f"  B (H1 2026)         → {totalB}")
        print(f"  C (1-15 Jul 2026)   → {totalC}")
        if totalA and totalB and totalB <= totalA and totalC <= totalB:
            print("  ✓ filter ทำงาน (จำนวนลดลงเมื่อช่วงแคบลง)")
        else:
            print("  ✗ filter อาจไม่ทำงาน (จำนวนไม่ลด)")

        page.wait_for_timeout(15000)
        ctx.close(); b.close()


if __name__ == "__main__":
    main()
