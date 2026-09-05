"""Interactive probe v2:
- โปรแกรมเปิด iframe URL ให้แล้ว
- User ทำมือ:
  1. เลือก 'สถานที่' จาก dropdown (คลิกกล่อง แล้วเลือกสาขาใด ๆ)
  2. กด 'เพิ่มวันนัดหมาย' → ปฏิทินโผล่
  3. คลิก 1 วัน (สีอ่อน = ว่าง)
  4. รอเห็น 'ช่วงเวลา' หรือ time slots
- ระบบจะเก็บ API call ตลอด — user ไม่ต้องกลับมากด Enter
- Auto-print API calls ทุก 5 วิ หลังคลิก

Just close browser to end
"""
from __future__ import annotations
import json, os, sys, time
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import login, ROOT, goto_tracking, apply_wa_filter, collect_all_wa_rows, build_detail_url

OUT_JSON = ROOT / "_probe_timeslot_v2_out.json"


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }

    all_captured = []
    last_count = 0

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
                    "ts": time.time(),
                })
        ctx.on("response", on_res)

        page = ctx.new_page()
        try:
            login(page, cfg)
            page.wait_for_timeout(1500)

            goto_tracking(page)
            page.wait_for_timeout(2000)
            apply_wa_filter(page, "", status_ids=["AP"])
            page.wait_for_timeout(2000)
            rows = collect_all_wa_rows(page, log=lambda *a: None)
            ap_rows = [r for r in rows if str(r.get("status") or "").upper() == "AP"]
            if not ap_rows:
                print("[!] ไม่มี AP record"); return 1

            target = ap_rows[0]
            url = build_detail_url(target)
            print(f"target: {target.get('reqNo')}")
            page.goto(url, wait_until="domcontentloaded", timeout=25_000)
            page.wait_for_timeout(3000)
            page.evaluate("() => { if(window.jQuery) jQuery('a[href=\"#tab_default_5\"]').tab('show'); }")
            page.wait_for_timeout(3500)

            iframe_src = page.evaluate("() => (document.querySelector('#link_appointment') || {}).src || ''")

            qp = ctx.new_page()
            qp.goto(iframe_src, wait_until="networkidle", timeout=30_000)
            qp.wait_for_timeout(5000)

            print("\n" + "=" * 60)
            print("👉 กรุณาทำในเบราว์เซอร์ที่เปิดล่าสุด:")
            print("   1. คลิก dropdown 'กรุณาเลือกสถานที่' → เลือก 1 สาขา")
            print("   2. กด 'เพิ่มวันนัดหมาย' → ปฏิทินจะโผล่")
            print("   3. คลิก 1 วัน (สีน้ำเงินอ่อน = เปิด)")
            print("   4. รอเห็นช่วงเวลา (time slot)")
            print("")
            print("ระบบจะแสดง API calls ต่อเนื่อง 60 วิ")
            print("=" * 60)

            all_captured.clear()

            # poll ทุก 5 วิ 12 รอบ = 60 วิ
            for i in range(12):
                qp.wait_for_timeout(5000)
                new_calls = all_captured[last_count:]
                last_count = len(all_captured)
                if new_calls:
                    print(f"\n[+{5*(i+1)}s] +{len(new_calls)} new call(s):")
                    for c in new_calls:
                        url_s = c["url"]
                        # highlight ถ้าน่าจะเป็น time-slot
                        marker = "🎯" if any(k in url_s.lower() for k in
                                             ("timeslot", "time-slot", "period", "slot", "session",
                                              "hour", "schedule", "book/create")) else "  "
                        print(f"  {marker} {c['status']} {c['method']} {url_s[:200]}")
                        # print body if short
                        if len(c["body"]) < 500:
                            print(f"     body: {c['body'][:400]}")

            # สรุปสุดท้าย
            print(f"\n=== SUMMARY ({len(all_captured)} calls) ===")
            unique_endpoints = set()
            for c in all_captured:
                # เอาเฉพาะ path หลัง /doe-booking/
                u = c["url"]
                if "/doe-booking/" in u:
                    ep = u.split("/doe-booking/", 1)[1].split("?", 1)[0]
                    unique_endpoints.add(ep)
            print(f"\nUnique endpoints ({len(unique_endpoints)}):")
            for ep in sorted(unique_endpoints):
                print(f"  · /{ep}")

            OUT_JSON.write_text(json.dumps({
                "unique_endpoints": sorted(unique_endpoints),
                "all_captured": all_captured,
            }, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"\n[done] → {OUT_JSON}")

            input("\n Enter to close ...")
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
