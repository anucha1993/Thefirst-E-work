"""รันดาวน์โหลดเอกสารผลอนุญาต (ใบแจ้งผล/ใบรับคำขอ/บต.50) หลายเลขคำขอรวด (login ครั้งเดียว)
กดสร้างเอกสารอัตโนมัติถ้ายังไม่เคยสร้าง (รอ ~1-2 นาที/เอกสาร) แล้วดาวน์โหลดต่อ
ใช้: python _run_result_docs_batch.py 69125300074577 69125300064821 ...
ไฟล์ที่ได้จะอยู่ใน reports/result_docs_manual/
"""
from __future__ import annotations
import os
import sys
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import (  # noqa: E402
    login, goto_tracking, _is_logged_out,
    _search_request, _open_first_detail, _extract_alien_eng_name,
    _download_response_doc_named, RESULT_DOC_TYPES, _receipt_safe_name,
)

OUT_DIR = Path(__file__).parent / "reports" / "result_docs_manual"
DOC_KEYS = ["result_notice", "request_receipt", "bt50"]
DEFAULT_REQS = [
    "69125300074577", "69125300064821", "69125300064776", "69125300064735",
    "69125300064668", "69125300064656", "69125300062460", "69125300062417",
    "69125300062377", "69125300062332", "69125300062305", "69125300062286",
    "69125300062264", "69125300062217",
]


def process_one(page, cfg, req_no: str, log=print) -> dict:
    """ประมวลผล 1 เลขคำขอ คืน {req_no, name_eng, docs:{dk:{status,file,error}}}"""
    goto_tracking(page)
    page.wait_for_timeout(1200)
    if _is_logged_out(page):
        log("     ⚠ session หมดอายุ — login ใหม่")
        login(page, cfg)
        goto_tracking(page)
        page.wait_for_timeout(1200)

    _search_request(page, req_no)
    if not _open_first_detail(page):
        return {"req_no": req_no, "name_eng": "", "status": "NOT_FOUND",
                "docs": {dk: {"status": "", "file": "", "error": ""} for dk in DOC_KEYS}}

    name_eng = _extract_alien_eng_name(page) or req_no
    name_safe = _receipt_safe_name(name_eng)
    docs: dict = {}
    for dk in DOC_KEYS:
        label = RESULT_DOC_TYPES[dk]["label"]
        log(f"   --- {label} ---")
        try:
            res = _download_response_doc_named(
                page, OUT_DIR, name_safe, RESULT_DOC_TYPES[dk], log=log, req_no=req_no,
            )
        except Exception as e:
            res = {"status": "FAIL", "file": "", "label": "", "error": str(e).splitlines()[0][:150]}
        docs[dk] = res
        log(f"     => {res['status']}  {res.get('file') or res.get('error','')}")
    statuses = [docs[dk]["status"] for dk in DOC_KEYS]
    if all(s in ("SUCCESS", "SKIP_EXISTS") for s in statuses):
        overall = "SUCCESS"
    elif any(s in ("SUCCESS", "SKIP_EXISTS") for s in statuses):
        overall = "PARTIAL"
    else:
        overall = "FAIL"
    return {"req_no": req_no, "name_eng": name_eng, "status": overall, "docs": docs}


def main() -> int:
    reqs = sys.argv[1:] if len(sys.argv) > 1 else DEFAULT_REQS
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }
    if not cfg["username"] or not cfg["password"]:
        print("ERR: EWP_USERNAME/EWP_PASSWORD ไม่ถูกตั้งใน .env")
        return 1
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"[i] บัญชี: {cfg['username']} | {len(reqs)} เลขคำขอ")
    print(f"[i] เอกสาร: {', '.join(RESULT_DOC_TYPES[d]['label'] for d in DOC_KEYS)}")

    results: list[dict] = []

    def _new_browser_ctx(pw):
        b = pw.chromium.launch(headless=False, args=["--start-maximized"])
        ctx = b.new_context(locale="th-TH", viewport={"width": 1500, "height": 900},
                            ignore_https_errors=True, accept_downloads=True)
        page = ctx.new_page()
        return b, ctx, page

    with sync_playwright() as pw:
        b, ctx, page = _new_browser_ctx(pw)
        try:
            print("[1] login... (แก้ reCAPTCHA)")
            login(page, cfg)
            print(f"[+] login OK ({page.url[:70]})")

            i = 1
            while i <= len(reqs):
                req_no = reqs[i - 1]
                print(f"\n{'=' * 70}\n[{i}/{len(reqs)}] เลขคำขอ {req_no}\n{'=' * 70}")
                try:
                    res = process_one(page, cfg, req_no, log=print)
                    results.append(res)
                    i += 1
                except Exception as e:
                    msg = str(e).splitlines()[0][:200]
                    # เบราว์เซอร์/หน้าเว็บถูกปิดกลางทาง (เคยพบเป็นระยะหลัง poll หนักๆ หลายรอบ)
                    # → เปิดเบราว์เซอร์ใหม่ + login ใหม่ แล้วทำเลขคำขอเดิมซ้ำ (ไม่ข้าม)
                    if "closed" in msg.lower() or "Target page" in msg or "context or browser" in msg:
                        print(f"  [!] เบราว์เซอร์ถูกปิดกลางทาง ({msg}) — เปิดใหม่ + login ใหม่ แล้วลองเลขนี้อีกครั้ง...")
                        try:
                            ctx.close()
                        except Exception:
                            pass
                        try:
                            b.close()
                        except Exception:
                            pass
                        b, ctx, page = _new_browser_ctx(pw)
                        try:
                            login(page, cfg)
                            print(f"  [+] login ใหม่สำเร็จ ({page.url[:70]}) — ทำเลขคำขอ {req_no} ต่อ")
                        except Exception as e2:
                            print(f"  [!] login ใหม่ไม่สำเร็จ: {e2} — ข้ามเลขนี้ไปก่อน")
                            results.append({"req_no": req_no, "name_eng": "", "status": "ERROR",
                                           "docs": {dk: {"status": "", "file": "", "error": "login ใหม่ไม่สำเร็จ"} for dk in DOC_KEYS}})
                            i += 1
                        # ไม่ i += 1 ที่นี่ — วนลูปซ้ำเลขคำขอเดิมด้วย browser ใหม่
                        continue
                    print(f"  [!] error: {msg}")
                    results.append({"req_no": req_no, "name_eng": "", "status": "ERROR",
                                   "docs": {dk: {"status": "", "file": "", "error": msg} for dk in DOC_KEYS}})
                    i += 1

            print(f"\n{'=' * 70}\n[สรุป]\n{'=' * 70}")
            for r in results:
                doc_summary = " | ".join(
                    f"{RESULT_DOC_TYPES[dk]['label']}={r['docs'][dk]['status'] or '-'}" for dk in DOC_KEYS
                )
                print(f"  {r['req_no']} ({r['name_eng']}): {r['status']} — {doc_summary}")
            print(f"\nไฟล์ที่ได้อยู่ใน: {OUT_DIR}")
        finally:
            try:
                ctx.close()
            except Exception:
                pass
            try:
                b.close()
            except Exception:
                pass
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
