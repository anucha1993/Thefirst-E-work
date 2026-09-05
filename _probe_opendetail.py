"""จับ openDetail() ตัวจริง + URL ที่ถูกต้องของแต่ละ form_type (โดยเฉพาะ CHANGE_44_22)
- login ด้วย .env (เบราว์เซอร์โผล่ให้แก้ reCAPTCHA เองถ้ามี)
- dump openDetail.toString() → เห็น mapping form_type → endpoint ทั้งหมด
- กรอง CHANGE_44_22 แล้วเรียก openDetail จริง จับ URL ปลายทาง (เทียบกับ build_detail_url)
ไม่พิมพ์รหัสผ่าน/ข้อมูลลับ
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import scrape_wa as s  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

OUT = s.ROOT / "_probe_opendetail_out.json"
TEST_FORM_TYPES = ["CHANGE_44_22", "CHANGE_44", "CHANGE_22"]


def main() -> int:
    cfg = s.load_config(require_login=True)
    cfg["headless"] = False
    dump: dict = {}

    with sync_playwright() as pw:
        browser = s._launch_chromium(pw, cfg, ["--disable-blink-features=AutomationControlled"])
        ctx = browser.new_context(
            locale="th-TH", timezone_id="Asia/Bangkok",
            viewport={"width": 1500, "height": 900},
        )
        page = ctx.new_page()
        try:
            print("[1] login...")
            s.login(page, cfg)
            print("[2] เปิด e-Tracking...")
            s.goto_tracking(page)
            if s._is_logged_out(page):
                s.ensure_session(page, cfg, log=print)
                s.goto_tracking(page)

            # ---- dump openDetail source ----
            src = page.evaluate(
                "() => (typeof openDetail==='function') ? openDetail.toString() : '(no openDetail)'"
            )
            dump["openDetail_src"] = src
            print("\n=== openDetail() source ===")
            print(src)
            print("=== end source ===\n")

            # ---- ทดสอบแต่ละ form_type: กรอง → เอาแถวแรก → เรียก openDetail จริง จับ URL ----
            dump["tests"] = {}
            for ft in TEST_FORM_TYPES:
                entry: dict = {"form_type": ft}
                try:
                    s.apply_wa_filter(page, ft, status_ids=["WP", "WCOSNA", "WA", "AP", "SS"])
                    rows = s.collect_all_wa_rows(page, log=lambda *_: None)
                    entry["n_rows"] = len(rows)
                    print(f"[{ft}] rows={len(rows)}")
                    if rows:
                        r = rows[0]
                        entry["row0"] = {k: r.get(k, "") for k in
                                         ("user_id", "group_id", "status", "form_type",
                                          "institution_id", "id", "reqNo")}
                        entry["build_detail_url"] = s.build_detail_url(r)
                        print(f"    build_detail_url : {entry['build_detail_url']}")
                        args = [r.get("user_id", ""), r.get("group_id", ""), r.get("status", ""),
                                r.get("form_type", ""), r.get("institution_id", ""), r.get("id", "")]
                        try:
                            with page.expect_navigation(timeout=15000, wait_until="domcontentloaded"):
                                page.evaluate(
                                    "(a)=>{ if(typeof openDetail==='function')"
                                    " openDetail(a[0],a[1],a[2],a[3],a[4],a[5]); }", args)
                            page.wait_for_timeout(1200)
                            entry["real_url"] = page.url
                            print(f"    REAL url (openDetail): {page.url}")
                            # กลับไป e-Tracking เพื่อ test ตัวถัดไป
                            s.goto_tracking(page)
                        except Exception as e:
                            entry["real_url_err"] = str(e).splitlines()[0][:200]
                            print(f"    openDetail nav err: {entry['real_url_err']}")
                            s.goto_tracking(page)
                except Exception as e:
                    entry["error"] = str(e).splitlines()[0][:200]
                    print(f"[{ft}] error: {entry['error']}")
                dump["tests"][ft] = entry
        finally:
            OUT.write_text(json.dumps(dump, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"\n[เสร็จ] เขียนผลลง {OUT.name}")
            page.wait_for_timeout(500)
            ctx.close()
            browser.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
