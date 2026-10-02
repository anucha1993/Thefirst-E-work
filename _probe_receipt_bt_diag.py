"""Diagnostic: หาสาเหตุที่ดาวน์โหลด บต.55 / บต.52 / บต.56 ไม่ได้ (โหมดดาวน์โหลดใบเสร็จ)
ใช้ flow เดียวกับ _process_one_receipt (ไม่ apply_wa_filter, ค้นหาตรงด้วยเลขคำขอ)
แล้ว dump สถานะแท็บ 'เอกสารตอบรับ' + หาลิงก์ GetDocumentConfirm ที่ label ตรง บต.55/52/56

ใช้: python _probe_receipt_bt_diag.py 6912530007462
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
    login, goto_tracking, _is_logged_out, ensure_session,
    _search_request, _open_first_detail, _is_detail_loaded,
)

DEFAULT_REQS = ["6912530007462"]


def dump_one(page, req_no: str) -> None:
    print(f"\n{'=' * 70}\n[เลขคำขอ] {req_no}\n{'=' * 70}")
    goto_tracking(page)
    page.wait_for_timeout(1500)
    if _is_logged_out(page):
        print("  [!] session หลุด — ข้าม")
        return
    _search_request(page, req_no)

    has_result = page.evaluate(r"""() => !!document.querySelector('a[onclick*="openDetail"]')""")
    print(f"  มีผลค้นหา (openDetail link) = {has_result}")
    # dump ตารางผลค้นหาทั้งแถว (เผื่อหลายแถว/สถานะ)
    rows = page.evaluate(
        r"""() => [...document.querySelectorAll('table tbody tr')].map(tr => (tr.innerText||'').replace(/\s+/g,' ').trim()).filter(Boolean)"""
    )
    print(f"  [ตารางผลค้นหา] -> {len(rows)} แถว")
    for r in rows[:10]:
        print(f"    {r[:150]}")

    if not _open_first_detail(page):
        print("  [!] เปิด detail ไม่สำเร็จ (ไม่พบผลค้นหา/openDetail) -> STOP")
        return
    print(f"  detail URL: {page.url}")
    print(f"  _is_detail_loaded = {_is_detail_loaded(page)}")

    resp_href = page.evaluate(
        r"""() => { const a=[...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a')]
            .find(a=>/เอกสารตอบรับ/.test(a.innerText||'')); return a?(a.getAttribute('href')||''):''; }"""
    )
    print(f"  [เอกสารตอบรับ] tab href = {resp_href!r}  (ว่าง = ไม่พบแท็บเลย)")
    if not resp_href:
        # dump รายชื่อแท็บทั้งหมดที่เจอ เผื่อชื่อแท็บเปลี่ยน
        all_tabs = page.evaluate(
            r"""() => [...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a')]
                .map(a => (a.innerText||'').trim()).filter(Boolean)"""
        )
        print(f"  [แท็บทั้งหมดที่เจอในหน้า] -> {all_tabs}")
        return

    clicked = False
    try:
        page.locator(f'a[href="{resp_href}"]').first.click(timeout=5000)
        clicked = True
    except Exception as e:
        print(f"    native click ล้มเหลว: {str(e).splitlines()[0][:80]} -> ลอง JS click")
        clicked = page.evaluate(
            r"""() => { const f=[...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a')]
                .find(a=>/เอกสารตอบรับ/.test(a.innerText||'')); if(f){f.click();return true;} return false;}"""
        )
    print(f"    clicked = {clicked}")
    pane_sel = resp_href if resp_href.startswith("#") else ".tab-pane.active"

    last_txt = ""
    for i in range(20):
        page.wait_for_timeout(500)
        last_txt = page.evaluate(
            "(sel)=>{const p=document.querySelector(sel)||document.querySelector('.tab-pane.active');return p?((p.innerText||'').trim()):'';}",
            pane_sel,
        )
        if last_txt and "ไม่มีเอกสารการตอบรับ" not in last_txt and len(last_txt) > 15:
            print(f"    pane โหลดเสร็จหลัง ~{(i + 1) * 0.5:.1f}s")
            break
    else:
        print(f"    pane ไม่เปลี่ยน/ว่าง (ข้อความ: {last_txt[:200]!r})")

    print("\n  ---- pane text เต็ม (เอกสารตอบรับ) ----")
    print("  " + last_txt.replace("\n", "\n  ")[:3000])

    # ---- หา label ที่ตรง บต.55 / บต.52(ชื่อเต็ม) / บต.56 พร้อมจำนวนลิงก์ GetDocumentConfirm ใกล้ๆ ----
    patterns = {
        "บต.55": r"บต\.?\s*55",
        "บต.52 (แบบแจ้งการจ้างคนต่างด้าวทำงาน)": r"แบบแจ้งการจ้างคนต่างด้าวทำงาน",
        "บต.56": r"บต\.?\s*56",
    }
    for name, pat in patterns.items():
        info = page.evaluate(
            r"""(pat) => {
                const re = new RegExp(pat);
                function findLabel(a) {
                    let node = a.parentElement;
                    for (let i = 0; i < 6 && node; i++) {
                        const prev = node.previousElementSibling;
                        if (prev) return (prev.innerText || '').trim();
                        node = node.parentElement;
                    }
                    return '';
                }
                const all = [...document.querySelectorAll('[onclick*="GetDocumentConfirm"]')];
                const matched = all.map(a => ({label: findLabel(a), onclick:(a.getAttribute('onclick')||'').slice(0,140)}))
                    .filter(x => re.test(x.label));
                return {totalGetDocumentConfirmLinks: all.length, matched};
            }""",
            pat,
        )
        print(f"\n  [{name}] pattern={pat!r}")
        print(f"    ทั้งหน้ามีลิงก์ GetDocumentConfirm ทั้งหมด = {info['totalGetDocumentConfirmLinks']}")
        print(f"    ที่ label ตรง pattern นี้ = {len(info['matched'])}")
        for m in info["matched"]:
            print(f"      label={m['label']!r}  onclick={m['onclick']}")

    # ---- outerHTML ของ #DetailDocumentList (โครงสร้างจริงทั้งหมด) ----
    rows_html = page.evaluate(
        r"""() => {
            const cont = document.querySelector('#DetailDocumentList');
            if (!cont) return null;
            return [...cont.children].map(el => (el.outerHTML || '').slice(0, 1500));
        }"""
    )
    if rows_html is None:
        print("\n  [!] ไม่พบ #DetailDocumentList element เลยในหน้านี้")
    else:
        print(f"\n  [#DetailDocumentList children] -> {len(rows_html)} แถว")
        for i, h in enumerate(rows_html, start=1):
            print(f"\n  --- แถว {i} ---")
            print("  " + h.replace("\n", "\n  "))


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
    login_cfg = {"username": cfg["username"], "password": cfg["password"],
                 "user_type": cfg["user_type"], "method": cfg["method"],
                 "login_timeout_ms": 120_000}
    print(f"[i] บัญชี: {cfg['username']} | เลขคำขอ: {reqs}")

    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, args=["--start-maximized"])
        ctx = b.new_context(locale="th-TH", viewport={"width": 1500, "height": 900},
                            ignore_https_errors=True, accept_downloads=True)
        page = ctx.new_page()
        try:
            print("[1] login... (แก้ reCAPTCHA)")
            login(page, login_cfg)
            print(f"[+] login OK ({page.url[:70]})")
            goto_tracking(page)
            if _is_logged_out(page):
                ensure_session(page, cfg); goto_tracking(page)
            for req_no in reqs:
                try:
                    dump_one(page, req_no)
                except Exception as e:
                    print(f"  [!] error: {e}")
            print("\n[เสร็จ] ปิด browser ได้เลย")
            page.wait_for_timeout(5000)
        finally:
            ctx.close()
            b.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
