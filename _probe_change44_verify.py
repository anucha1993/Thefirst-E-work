"""ตรวจ end-to-end: CHANGE_44_22 หลังแก้ build_detail_url — เปิด detail ได้ + อ่าน tab/ข้อมูลไหม
ไม่พิมพ์รหัสผ่าน/ข้อมูลลับ
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import scrape_wa as s  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

FT = "CHANGE_44_22"
OUT = s.ROOT / "_probe_change44_out.json"


def main() -> int:
    cfg = s.load_config(require_login=True)
    cfg["headless"] = False
    dump: dict = {"form_type": FT, "rows": []}

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
            print("[2] เปิด e-Tracking + กรอง CHANGE_44_22...")
            s.goto_tracking(page)
            if s._is_logged_out(page):
                s.ensure_session(page, cfg, log=print)
                s.goto_tracking(page)
            s.apply_wa_filter(page, FT, status_ids=["WP", "WCOSNA", "WA", "AP", "SS"])
            rows = s.collect_all_wa_rows(page, log=print)
            SCAN_MAX = min(10, len(rows))
            print(f"[3] ได้ {len(rows)} แถว — จะสแกน {SCAN_MAX} แถว")

            def _is_ws(sk: str) -> bool:
                return ("นายจ้าง" not in sk) and any(
                    k in sk for k in ("ข้อมูลการขออนุญาต", "สถานประกอบการ", "สถานที่ทำงาน")
                )

            report_rows: list[dict] = []
            with_section: list[str] = []
            for i, row in enumerate(rows[:SCAN_MAX], 1):
                url = s.build_detail_url(row)
                detail = s.scrape_permit_detail(page, row, cfg=cfg, log=lambda *a, **k: None)
                tabs = s._resolve_detail_tabs(page)
                permit_tab = s._find_tab(tabs, "คำขออนุญาต")
                pane_data = page.evaluate(
                    s._PERMIT_TAB_JS,
                    {"paneId": permit_tab.get("pane", ""), "index": permit_tab.get("index", -1)},
                ) or {}
                pairs = pane_data.get("pairs") or []
                secs: dict[str, int] = {}
                for pr in pairs:
                    secs[pr.get("section", "")] = secs.get(pr.get("section", ""), 0) + 1
                has_ws = any(_is_ws(sk) for sk in secs)

                prov = detail.get("estab_province", "")
                company = (detail.get("company_main", "") or detail.get("employer_company", ""))[:40]
                flag = "  ★ มี section สถานประกอบการ" if has_ws else ""
                print(f"[{i}] {row.get('reqNo','')} | จังหวัด={prov or '(ว่าง)'} | บริษัท={company}{flag}")
                # วินิจฉัย: ถ้ามีที่อยู่ (มีคำว่า 'จังหวัด') นอก section นายจ้าง ให้โชว์
                if not has_ws:
                    for pr in pairs:
                        sk = pr.get("section", "")
                        val = pr.get("value", "") or ""
                        if "จังหวัด" in val and "นายจ้าง" not in sk:
                            print(f"      (ที่อยู่นอกนายจ้าง) [{sk or '-'}] {pr.get('label','') or '-'} = {val[:110]}")
                if has_ws:
                    with_section.append(row.get("reqNo", ""))
                    for pr in pairs:
                        if _is_ws(pr.get("section", "")):
                            print(f"      - [{pr.get('section','')}] {pr.get('label','')} = {pr.get('value','')[:110]}")

                report_rows.append({**row, **detail})
                dump["rows"].append({
                    "reqNo": row.get("reqNo", ""), "url": url,
                    "sections": secs, "has_section": has_ws, "detail": detail,
                })
                if len(with_section) >= 2:
                    print(f"    [พบครบ 2 record ที่มี section — หยุดสแกนที่แถว {i}]")
                    break

            print(f"\n[สรุป] มี section สถานประกอบการใน: {with_section or ('(ไม่พบใน %d แถวที่สแกน)' % len(report_rows))}")
            print("===== รายงาน 5 คอลัมน์ (ตามที่จะออก Excel) =====")
            for r in report_rows:
                company = r.get("company_main", "") or r.get("employer_company", "")
                prov = r.get("estab_province", "")
                found = r.get("approve_found", "") == "พบ"
                status_val = "อนุมัติคำขอ" if found else "ไม่พบการอนุมัติคำขอ"
                remark = " — ".join(
                    p for p in (
                        (r.get("approve_result", "") or "").strip(),
                        (r.get("approve_date", "") or "").strip(),
                    ) if p
                ) if found else ""
                print(f"  • {r.get('reqNo','')} | บริษัท={company} | จังหวัด={prov or '(ว่าง)'} | {status_val} | {remark}")
            s.REPORTS_DIR.mkdir(parents=True, exist_ok=True)
            xlsx = s.REPORTS_DIR / "_probe_change44_5col.xlsx"
            s.save_excel_permit_report(report_rows, xlsx)
            print(f"[Excel] บันทึก {xlsx}")
        finally:
            OUT.write_text(json.dumps(dump, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"\n[เสร็จ] เขียนผลลง {OUT.name}")
            try:
                ctx.close()
                browser.close()
            except Exception:
                pass
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
