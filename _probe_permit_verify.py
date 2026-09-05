"""ตรวจสอบจริง: โหมด permit_report อ่าน DOM ถูกไหม
- login ด้วย .env (เบราว์เซอร์โผล่ให้แก้ reCAPTCHA เองถ้ามี)
- ดึงรายการจริงจาก e-Tracking ไม่กี่รายการ
- เรียกฟังก์ชันจริง (_resolve_detail_tabs / _PERMIT_TAB_JS / scrape_permit_detail)
- dump: แท็บที่ resolve ได้, pairs ของแท็บคำขออนุญาต, ผลที่สกัดได้ → JSON + สรุปบนจอ
ไม่พิมพ์รหัสผ่าน/ข้อมูลลับ
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import scrape_wa as s  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

MAX_ROWS = 2
OUT = s.ROOT / "_probe_permit_verify_out.json"


def main() -> int:
    cfg = s.load_config(require_login=True)
    cfg["headless"] = False  # ให้เห็นเบราว์เซอร์ เผื่อ reCAPTCHA
    req_type = ""  # ดึงทุกประเภทคำขอ (default MT_59_MOU_RENEWAL คืน 0 แถว)
    status_ids = cfg.get("filter_status_ids") or ["WP", "WCOSNA", "WA", "AP", "SS"]

    print(f"[i] request_type={req_type or '(ทั้งหมด)'} | statuses={status_ids}")
    dump: dict = {"request_type": req_type, "statuses": status_ids, "rows": []}
    report_rows: list[dict] = []  # เก็บ row+detail เพื่อสร้างรายงาน 5 คอลัมน์จริง

    with sync_playwright() as pw:
        browser = s._launch_chromium(pw, cfg, ["--disable-blink-features=AutomationControlled"])
        ctx = browser.new_context(
            locale="th-TH", timezone_id="Asia/Bangkok",
            viewport={"width": 1500, "height": 900},
        )
        page = ctx.new_page()
        try:
            print("[1] login... (ถ้ามี reCAPTCHA ให้กดในเบราว์เซอร์ที่เปิดขึ้น)")
            s.login(page, cfg)
            print("[2] เปิด e-Tracking + ตั้งฟิลเตอร์...")
            s.goto_tracking(page)
            if s._is_logged_out(page):
                s.ensure_session(page, cfg, log=print)
                s.goto_tracking(page)

            s.apply_wa_filter(page, req_type, status_ids=status_ids)
            rows = s.collect_all_wa_rows(page, log=print)
            from collections import Counter
            ft_dist = Counter(r.get("form_type", "") for r in rows)
            print(f"[3] ได้ {len(rows)} แถว | form_types={dict(ft_dist)}")
            # เลือกฟอร์มที่ 'มี section สถานประกอบการ' ก่อน (เลี่ยง Change/Exit/Replace ที่ไม่มี)
            preferred = [r for r in rows if r.get("form_type", "") not in s.PATH_STYLE_FORM_TYPES]
            picks = (preferred or rows)[:MAX_ROWS]
            print(f"    จะตรวจ {len(picks)} แถว (เลี่ยงฟอร์ม Change ถ้ามีตัวเลือกอื่น)")

            for i, row in enumerate(picks, 1):
                url = s.build_detail_url(row)
                print(f"\n[{i}] {row.get('reqNo','')} | form_type={row.get('form_type','')}")
                print(f"    URL: {url}")

                # นำทางเข้า detail เอง เพื่อ dump แท็บ + pairs ดิบ
                ok = s._open_detail_direct(page, row)
                entry: dict = {
                    "reqNo": row.get("reqNo", ""),
                    "form_type": row.get("form_type", ""),
                    "url": url,
                    "open_ok": bool(ok),
                }
                if ok:
                    tabs = s._resolve_detail_tabs(page)
                    entry["tabs"] = tabs
                    status_tab = s._find_tab(tabs, "สถานะคำขอ")
                    permit_tab = s._find_tab(tabs, "คำขออนุญาต")
                    status_pane = status_tab.get("pane") or "tab_default_1"
                    permit_pane = permit_tab.get("pane") or "tab_default_2"
                    permit_index = permit_tab.get("index", -1)
                    entry["status_pane"] = status_pane
                    entry["permit_pane"] = permit_pane
                    print("    แท็บที่พบ: " + " | ".join(
                        f"{t.get('label','?')}→#{t.get('pane','?')}[{t.get('index','?')}]" for t in tabs if t.get("label")
                    ))
                    print(f"    → สถานะคำขอ=#{status_pane} | คำขออนุญาต=#{permit_pane}[{permit_index}]")

                    # คลิกแท็บคำขออนุญาต + ดึง pairs ดิบ (signature ใหม่: {paneId, index})
                    page.evaluate(
                        "(a)=>{let x=a.paneId?document.querySelector('a[href=\"#'+a.paneId+'\"]'):null;"
                        "if(x)x.click();}",
                        {"paneId": permit_pane, "index": permit_index},
                    )
                    page.wait_for_timeout(1000)
                    data = page.evaluate(s._PERMIT_TAB_JS, {"paneId": permit_pane, "index": permit_index}) or {}
                    pairs = data.get("pairs") or []
                    entry["pairs"] = pairs
                    entry["permit_raw"] = (data.get("raw") or "")[:2000]
                    # แสดงทุก section ที่พบ + label/value ที่เกี่ยวข้อง
                    secs: dict = {}
                    for p in pairs:
                        secs[p.get("section", "")] = secs.get(p.get("section", ""), 0) + 1
                    print(f"    pairs ทั้งหมด {len(pairs)} รายการ; sections={json.dumps(secs, ensure_ascii=False)}")
                    for p in pairs:
                        blob = f"{p.get('section','')}|{p.get('label','')}"
                        if any(k in blob for k in ("สถานประกอบการ", "จังหวัด", "นายจ้าง", "ผู้รับอนุญาต", "สถานที่ทำงาน", "บริษัท", "หน่วยงาน", "ที่อยู่")):
                            print(f"      [{p.get('section','')}] {p.get('label','')} = {p.get('value','')}")

                # เรียกฟังก์ชันจริงที่ใช้ใน production
                detail = s.scrape_permit_detail(page, row, cfg=cfg, log=print)
                entry["result"] = detail
                report_rows.append({**row, **detail})  # merge: reqNo จาก row + ผลสกัด
                print("    === ผลจาก scrape_permit_detail ===")
                print(f"    อนุมัติคำขอ : {detail.get('approve_found','')} | {detail.get('approve_date','')} | {detail.get('approve_result','')}")
                print(f"    บริษัท(ผู้รับอนุญาต) : {detail.get('estab_company','')}")
                print(f"    บริษัท(นายจ้าง)      : {detail.get('employer_company','')}")
                print(f"    จังหวัด      : {detail.get('estab_province','')}")
                print(f"    Section ใช้   : {detail.get('estab_section','')}")
                print(f"    tabs_read    : {detail.get('tabs_read','')}")
                if detail.get("scrape_errors"):
                    print(f"    errors       : {detail['scrape_errors']}")

                dump["rows"].append(entry)

            # สร้างรายงาน 5 คอลัมน์จริง (เหมือนที่จะออก Excel) + เซฟไฟล์ให้เปิดดู
            print("\n===== รายงาน 5 คอลัมน์ (ตามที่จะออก Excel) =====")
            for r in report_rows:
                company = r.get("company_main", "") or r.get("employer_company", "")
                province = r.get("estab_province", "")
                found = r.get("approve_found", "") == "พบ"
                status_val = "อนุมัติคำขอ" if found else "ไม่พบการอนุมัติคำขอ"
                remark = " — ".join(
                    p for p in (
                        (r.get("approve_result", "") or "").strip(),
                        (r.get("approve_date", "") or "").strip(),
                    ) if p
                ) if found else ""
                print(f"  • บริษัท={company} | จังหวัด={province or '(ว่าง)'} | เลขคำขอ={r.get('reqNo','')} | สถานะ={status_val} | หมายเหตุ={remark}")
            s.REPORTS_DIR.mkdir(parents=True, exist_ok=True)
            xlsx = s.REPORTS_DIR / "_probe_permit_5col.xlsx"
            s.save_excel_permit_report(report_rows, xlsx)
            print(f"[Excel] บันทึก {xlsx}")

            OUT.write_text(json.dumps(dump, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"\n[เสร็จ] เขียนผลลง {OUT.name}")
        finally:
            ctx.close(); browser.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
