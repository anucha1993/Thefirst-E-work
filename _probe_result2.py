"""ทดสอบดึงจริง 'โหมด results' (ใบแจ้งผล / ใบรับคำขอ / บต.50) สัก 2 record
โดยเรียกฟังก์ชันจริงใน scrape_wa: _process_one_result_by_ref (path เดียวกับ GUI โหมด results).

- login ครั้งเดียว (แก้ reCAPTCHA เอง ~120 วิ อย่าปิดเบราว์เซอร์กลางคัน)
- ค่าเริ่มต้น: record #1 = 69125200578664 (มี บต.50 แน่นอน) + auto-discover อีก 1 เลข (สถานะ AP/SS)
- override ได้: python _probe_result2.py <เลขคำขอ1> <เลขคำขอ2> ...
- ดาวน์โหลดครบ 3 เอกสาร: ใบแจ้งผล / ใบรับคำขอ / บต.50 อ.6
- เอาต์พุตลงโฟลเดอร์ _probe_result2_docs/<บริษัท>/ + รายงาน _probe_result2_report.xlsx
พิมพ์เฉพาะ username เท่านั้น — ไม่พิมพ์รหัสผ่าน
"""
from __future__ import annotations
import os
import shutil
import sys
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import (  # noqa: E402
    login, goto_tracking, apply_wa_filter, collect_all_wa_rows,
    _process_one_result_by_ref, _save_result_docs_report,
    _is_logged_out, ensure_session,
    RESULT_DOC_TYPES, ROOT,
)

FORCE_REQ = "69125200578664"          # เลขที่รู้ว่ามี บต.50 แน่นอน (ใช้เป็น record แรก)
DOC_KEYS = ["result_notice", "request_receipt", "bt50"]  # ทดสอบครบ 3 เอกสาร
N_RECORDS = 2


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
        "headless": False,
        "login_timeout_ms": 120_000,
    }
    if not cfg["username"] or not cfg["password"]:
        print("ERR: EWP_USERNAME/EWP_PASSWORD ไม่ถูกตั้งใน .env")
        return 1
    login_cfg = {
        "username": cfg["username"], "password": cfg["password"],
        "user_type": cfg["user_type"], "method": cfg["method"],
    }

    arg_reqs = [a.strip() for a in sys.argv[1:] if a.strip()]
    print(f"[i] บัญชี: {cfg['username']} ({cfg['user_type']})")
    print(f"[i] เอกสารที่จะทดสอบ: {', '.join(RESULT_DOC_TYPES[d]['label'] for d in DOC_KEYS)}")
    if arg_reqs:
        print(f"[i] ใช้เลขคำขอจาก argv: {arg_reqs}")
    else:
        print(f"[i] auto: record#1={FORCE_REQ} + auto-discover อีก {N_RECORDS - 1} เลข (AP/SS)")

    docs_dir = ROOT / "_probe_result2_docs"
    if docs_dir.exists():
        shutil.rmtree(docs_dir, ignore_errors=True)   # เริ่มสด ให้ดาวน์โหลดจริง (ไม่ข้ามด้วย resume)
    docs_dir.mkdir(parents=True, exist_ok=True)
    report_path = ROOT / "_probe_result2_report.xlsx"

    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, args=["--start-maximized"])
        ctx = b.new_context(locale="th-TH", viewport={"width": 1500, "height": 900},
                            ignore_https_errors=True, accept_downloads=True)
        page = ctx.new_page()
        try:
            print("[1] login... (แก้ reCAPTCHA ในเบราว์เซอร์)")
            login(page, login_cfg)
            print(f"[+] login OK ({page.url[:70]})")
            goto_tracking(page)
            if _is_logged_out(page):
                ensure_session(page, cfg)
                goto_tracking(page)
            page.wait_for_timeout(1500)

            # ---- เลือก 2 record ----
            recs: list[dict] = []          # {seq, req_no, username, prebuilt}
            if arg_reqs:
                for i, rq in enumerate(arg_reqs[:N_RECORDS], start=1):
                    recs.append({"seq": str(i), "req_no": rq,
                                 "username": cfg["username"], "prebuilt": None})
            else:
                print("[2] เก็บรายการคำขอ (สถานะ AP + SS) เพื่อเลือกเลขทดสอบ...")
                apply_wa_filter(page, "", status_ids=["AP", "SS"])
                rows = collect_all_wa_rows(page)
                by_req = {str(r.get("reqNo", "")).strip(): r for r in rows if r.get("reqNo")}
                print(f"[+] พบ {len(by_req)} คำขอ")
                # record#1 = FORCE_REQ (ใช้ prebuilt ถ้าเจอในรายการ ไม่งั้น fallback ค้นหา)
                recs.append({"seq": "1", "req_no": FORCE_REQ, "username": cfg["username"],
                             "prebuilt": by_req.get(FORCE_REQ)})
                # record#2 = เลขแรกที่ไม่ใช่ FORCE_REQ
                for rq, row in by_req.items():
                    if rq != FORCE_REQ:
                        recs.append({"seq": "2", "req_no": rq,
                                     "username": cfg["username"], "prebuilt": row})
                        break

            print(f"[3] จะทดสอบ {len(recs)} record: {[r['req_no'] for r in recs]}")

            results: list[dict] = []
            for r in recs:
                rec = {"seq": r["seq"], "name_eng": "", "req_no": r["req_no"],
                       "username": r["username"], "passport": "", "row_index": r["seq"]}
                print(f"\n[4] === record {r['seq']}: {r['req_no']} ===")
                res = _process_one_result_by_ref(
                    page, rec, login_cfg, docs_dir, DOC_KEYS,
                    log=print, prebuilt_row=r["prebuilt"],
                )
                results.append(res)

            _save_result_docs_report(results, report_path, DOC_KEYS, log=print)

            # ---- สรุป ----
            print("\n========== สรุปผล ==========")
            for res in results:
                print(f"• คำขอ {res.get('req_no')} | {res.get('name_eng') or '-'} "
                      f"| สถานะ={res.get('status')}")
                for dk in DOC_KEYS:
                    d = res["docs"].get(dk, {})
                    print(f"    - {RESULT_DOC_TYPES[dk]['label']}: {d.get('status')} "
                          f"{d.get('file') or ''} {('| ' + d['error']) if d.get('error') else ''}")
            print("\n---- ไฟล์ที่ได้จริงในโฟลเดอร์ ----")
            pdfs = sorted(docs_dir.rglob("*.pdf"))
            for p in pdfs:
                head = p.read_bytes()[:5]
                print(f"    {p.relative_to(docs_dir)}  ({p.stat().st_size // 1024} KB, header={head!r})")
            ok = sum(1 for res in results if res.get("status") in ("SUCCESS", "PARTIAL"))
            print(f"\nสำเร็จ/บางส่วน {ok}/{len(results)} record | ไฟล์ PDF {len(pdfs)} ไฟล์ | รายงาน: {report_path.name}")
            return 0 if ok else 3
        finally:
            page.wait_for_timeout(1200)
            ctx.close(); b.close()


if __name__ == "__main__":
    raise SystemExit(main())
