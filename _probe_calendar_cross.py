"""Probe: ทดสอบว่า iframe token 1 อัน call calendar ของสาขาอื่นได้ไหม
- เปิด iframe จาก 1 record (Vararat account มี CCO และ PTE)
- call /calendar/get-data สำหรับ 3 branches ต่าง ๆ
"""
from __future__ import annotations
import json
import os
import sys
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import login, ROOT  # noqa: E402

# record ใด ๆ ที่ status=APSS (จองแล้ว) — เอาจาก probe ก่อนหน้า
TARGET_URL = "https://eworkpermit.doe.go.th/RenewMOU/DetailFormRenewMOU?user_id=303816&status=APSS&group_id=69123000059307&form_type=MT_59_MOU_RENEWAL"
OUT_JSON = ROOT / "_probe_calendar_cross_out.json"


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }
    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, args=["--start-maximized"])
        ctx = b.new_context(locale="th-TH", viewport={"width": 1600, "height": 1000})
        page = ctx.new_page()
        try:
            login(page, cfg)
            page.wait_for_timeout(1500)
            page.goto(TARGET_URL, wait_until="domcontentloaded", timeout=45_000)
            page.wait_for_timeout(4000)

            # trigger tab นัดหมาย
            page.evaluate("""() => {
                if (window.jQuery) jQuery('a[href="#tab_default_5"]').tab('show');
            }""")
            page.wait_for_timeout(3500)

            iframe_src = page.evaluate("() => document.querySelector('#link_appointment').src")
            print(f"[+] iframe src: {iframe_src[:120]}...")

            # เปิด iframe ใน new tab เพื่อทำงาน API ใน same origin
            qp = ctx.new_page()
            qp.goto(iframe_src, wait_until="networkidle", timeout=30_000)
            qp.wait_for_timeout(5000)

            # ทดสอบ 3 branches ต่าง ๆ + 2 เดือน
            TESTS = [
                ("CCO-SC-S-001", 2026, 8, "ฉะเชิงเทรา"),
                ("CCO-SC-S-001", 2026, 12, "ฉะเชิงเทรา (Dec)"),
                ("PTE-SC-M-001", 2026, 8, "ปทุมธานี"),
                ("PTE-SC-M-001", 2026, 12, "ปทุมธานี (Dec)"),
                ("MDH-OB-M-001", 2026, 8, "มุกดาหาร (แรกรับ)"),
                ("BKK-SC-L-001", 2026, 8, "กรุงเทพ (guess)"),
            ]
            results = []
            for br, y, m, name in TESTS:
                date_str = f"{y}-{m:02d}-01"
                url = f"/doe-booking/api/v2/calendar/get-data/{br}/{date_str}?month={m}&year={y}"
                print(f"\n[test] {name}  {br}  {y}-{m:02d}")

                # เรียกจาก origin queue-fe-uat
                res = qp.evaluate(r"""async (url) => {
                    try {
                        const r = await fetch(url, {credentials: 'include'});
                        const status = r.status;
                        const text = await r.text();
                        let parsed = null;
                        try { parsed = JSON.parse(text); } catch(e){}
                        return {status, ok: r.ok, textLen: text.length, parsed};
                    } catch(e) {
                        return {error: e.message};
                    }
                }""", url)
                print(f"  status={res.get('status')}  ok={res.get('ok')}  textLen={res.get('textLen')}")
                if res.get("parsed"):
                    parsed = res["parsed"]
                    data = parsed.get("data") if isinstance(parsed, dict) else None
                    if data:
                        print(f"  data: {len(data)} วัน")
                        # เอา 5 อันแรก
                        for d in data[:5]:
                            print(f"    · {d.get('date')}: open={d.get('open')}  "
                                  f"maxNormal={d.get('maxNormal')}  "
                                  f"left={d.get('count_left_Normal')}")
                    elif isinstance(parsed, dict) and (parsed.get("error") or parsed.get("message")):
                        print(f"  ⚠ API returned: {parsed}")
                results.append({"branch": br, "name": name, "year": y, "month": m,
                                "url": url, "result": res})

            # เพิ่มทดสอบ /branch/by-codes ด้วย (ไม่ต้อง branch code)
            print("\n[test] /branch/by-codes")
            res_br = qp.evaluate(r"""async () => {
                try {
                    const r = await fetch('/doe-booking/api/v1/branch/by-codes', {
                        method: 'POST',
                        credentials: 'include',
                        headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify([])
                    });
                    const status = r.status;
                    const text = await r.text();
                    return {status, ok: r.ok, textLen: text.length,
                            parsed: JSON.parse(text || '[]')};
                } catch(e) { return {error: e.message}; }
            }""")
            print(f"  status={res_br.get('status')} branches count={len(res_br.get('parsed') or [])}")
            if res_br.get("parsed"):
                print("  first 5 branches:")
                for br in (res_br["parsed"] or [])[:5]:
                    print(f"    · {br.get('branch_code_id'):20} {br.get('branch_name_th', '')[:60]}")

            OUT_JSON.write_text(
                json.dumps({"tests": results, "branches": res_br}, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            print(f"\n[done] → {OUT_JSON}")

            input("\n Enter to close ...")
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
