"""Interactive probe: user จะคลิกเลือก สาขา+วัน ในเบราว์เซอร์เอง
ผมจะ capture time-slot API + save ผลไว้ให้ดู
กด Enter ที่ terminal เมื่อคลิกเลือกครบ (สาขา + วัน) แล้ว
"""
from __future__ import annotations
import json, os, sys
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import login, ROOT, goto_tracking, apply_wa_filter, collect_all_wa_rows, build_detail_url

OUT_JSON = ROOT / "_probe_timeslot_interactive_out.json"


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }

    all_captured = []

    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, args=["--start-maximized"])
        ctx = b.new_context(locale="th-TH", viewport={"width": 1600, "height": 1000})

        def on_res(res):
            u = res.url or ""
            if "queue-be" in u:
                try: text = res.text()
                except Exception: text = ""
                all_captured.append({
                    "status": res.status,
                    "method": res.request.method,
                    "url": u[:280],
                    "body": text[:3000],
                })
        ctx.on("response", on_res)

        page = ctx.new_page()
        try:
            login(page, cfg)
            page.wait_for_timeout(1500)

            # เปิด e-Tracking + filter AP → หา record
            goto_tracking(page)
            page.wait_for_timeout(2500)
            apply_wa_filter(page, "", status_ids=["AP"])
            page.wait_for_timeout(2000)
            rows = collect_all_wa_rows(page, log=lambda *a: None)
            ap_rows = [r for r in rows if str(r.get("status") or "").upper() == "AP"]
            if not ap_rows:
                print("[!] ไม่มี AP record ใน account นี้")
                # ลอง APSS แทน
                goto_tracking(page)
                apply_wa_filter(page, "", status_ids=["AP", "SS"])
                page.wait_for_timeout(2000)
                rows = collect_all_wa_rows(page, log=lambda *a: None)
                if not rows:
                    print("[!] ไม่มี record เลย"); return 1
                ap_rows = rows

            print(f"AP/SS rows: {len(ap_rows)} — ใช้ record แรก")
            target = ap_rows[0]
            url = build_detail_url(target)
            print(f"target: {target.get('reqNo')}")
            page.goto(url, wait_until="domcontentloaded", timeout=25_000)
            page.wait_for_timeout(3000)
            page.evaluate("() => { if(window.jQuery) jQuery('a[href=\"#tab_default_5\"]').tab('show'); }")
            page.wait_for_timeout(3500)

            iframe_src = page.evaluate("() => (document.querySelector('#link_appointment') || {}).src || ''")

            # เปิด iframe URL ใน new tab เพื่อ user คลิกได้สะดวก
            qp = ctx.new_page()
            qp.goto(iframe_src, wait_until="networkidle", timeout=30_000)
            qp.wait_for_timeout(4000)

            print("\n" + "=" * 60)
            print("👉 กรุณาทำในเบราว์เซอร์ที่เปิดขึ้นมาล่าสุด:")
            print("   1. เลือก 'สถานที่' จาก dropdown")
            print("   2. กด 'เพิ่มวันนัดหมาย'")
            print("   3. เลือกวัน (ที่ปฏิทินโผล่มา)")
            print("   4. รอ 'เลือกช่วงเวลา' โชว์ time slot")
            print("   5. คลิก 1 time slot")
            print("   → แล้วกลับมาที่ terminal นี้ กด Enter")
            print("=" * 60)

            all_captured.clear()  # เริ่มบันทึกใหม่จากตรงนี้
            input("\n>> กด Enter หลังคลิกครบแล้ว...")

            # ตอนนี้ capture ทุก API request หลัง user ทำ
            print(f"\n[captured] {len(all_captured)} API calls:")

            # แยกกลุ่ม: calendar (สรุปวัน) vs time-slot (ต่อวัน)
            calendars = [c for c in all_captured if "calendar/get-data" in c["url"]]
            timeslots = [c for c in all_captured if "timeslot" in c["url"].lower() or "time-slot" in c["url"].lower()
                         or "schedule" in c["url"].lower() or "session" in c["url"].lower()
                         or "slot" in c["url"].lower() or "period" in c["url"].lower()]
            others = [c for c in all_captured
                      if c not in calendars and c not in timeslots
                      and "annotation" not in c["url"] and "dictionary" not in c["url"]
                      and "auth" not in c["url"] and "branch" not in c["url"]]

            print(f"\n=== calendar API calls ({len(calendars)}) ===")
            for c in calendars:
                print(f"  {c['status']} {c['method']} {c['url']}")

            print(f"\n=== time-slot / schedule / period API calls ({len(timeslots)}) ===")
            for c in timeslots:
                print(f"\n  {c['status']} {c['method']} {c['url']}")
                print(f"    body (first 1500): {c['body'][:1500]}")

            print(f"\n=== other queue-be calls ({len(others)}) ===")
            for c in others:
                print(f"\n  {c['status']} {c['method']} {c['url']}")
                if len(c['body']) < 800:
                    print(f"    body: {c['body'][:600]}")
                else:
                    print(f"    body (first 300): {c['body'][:300]}")

            OUT_JSON.write_text(json.dumps({
                "calendars": calendars,
                "timeslots": timeslots,
                "others": others,
                "all": all_captured,
            }, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"\n[done] all captured → {OUT_JSON}")

            input("\n Enter to close ...")
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
