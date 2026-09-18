"""Diagnostic: หาสาเหตุที่ 'ใบแจ้งผลใบอนุญาตทำงาน' ไม่ถูกดาวน์โหลด (แต่เอกสารอื่นได้หมด)
ค้นหาตรงด้วยเลขคำขอ (ไม่ filter สถานะ) → เปิด detail → เข้าแท็บ 'เอกสารตอบรับ' → dump
ทุก element ที่คลิกได้ + จำนวน onclick=GetDocumentConfirm ที่อยู่ใกล้ข้อความ 'ใบแจ้งผล'
ใช้: python _probe_result_notice_diag.py 69125300001634 69125300002204 69125300002219
พิมพ์เฉพาะ username — ไม่พิมพ์รหัสผ่าน
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
    login, goto_tracking, apply_wa_filter, _is_logged_out, ensure_session,
    _search_request, _open_first_detail, _is_detail_loaded,
)

DEFAULT_REQS = ["69125300001634", "69125300002204", "69125300002219"]


def dump_one(page, req_no: str) -> None:
    print(f"\n{'=' * 70}\n[เลขคำขอ] {req_no}\n{'=' * 70}")
    goto_tracking(page)
    if _is_logged_out(page):
        print("  [!] session หลุด — ข้าม")
        return
    apply_wa_filter(page, "", status_ids=["WP", "WCOSNA", "WA", "AP", "SS"])
    _search_request(page, req_no)
    if not _open_first_detail(page):
        print("  [!] เปิด detail ไม่สำเร็จ (ไม่พบผลค้นหา/openDetail)")
        return
    print(f"  detail URL: {page.url}")
    print(f"  _is_detail_loaded = {_is_detail_loaded(page)}")

    resp_href = page.evaluate(
        r"""() => { const a=[...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a')]
            .find(a=>/เอกสารตอบรับ/.test(a.innerText||'')); return a?(a.getAttribute('href')||''):''; }"""
    )
    print(f"  [เอกสารตอบรับ] tab href = {resp_href!r}")
    clicked = False
    try:
        if resp_href.startswith("#"):
            page.locator(f'a[href="{resp_href}"]').first.click(timeout=5000)
        else:
            page.get_by_text("เอกสารตอบรับ", exact=False).first.click(timeout=5000)
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
        print(f"    pane ไม่เปลี่ยน (ข้อความ: {last_txt[:80]!r})")

    # ---- dump ทุก element ที่มีข้อความตรง /ใบแจ้งผล/ พร้อมจำนวนลิงก์ GetDocumentConfirm ในตัวมันเอง ----
    matches = page.evaluate(
        r"""(sel) => {
            const re = /ใบแจ้งผล/;
            const pane = document.querySelector(sel) || document.querySelector('.tab-pane.active') || document;
            const out = [];
            pane.querySelectorAll('*').forEach(el => {
                const txt = (el.innerText || '').trim();
                if (!txt || !re.test(txt)) return;
                const links = [...el.querySelectorAll('a, button, [onclick]')]
                    .filter(b => (b.getAttribute('onclick') || '').includes('GetDocumentConfirm'));
                out.push({
                    tag: el.tagName.toLowerCase(),
                    txt: txt.replace(/\s+/g, ' ').slice(0, 100),
                    nLinks: links.length,
                    onclicks: links.map(l => (l.getAttribute('onclick') || '').slice(0, 140)),
                });
            });
            return out;
        }""",
        pane_sel,
    )
    print(f"\n  [Elements matching /ใบแจ้งผล/] -> {len(matches)} elements")
    for m in matches:
        print(f"    <{m['tag']}> nLinks(GetDocumentConfirm)={m['nLinks']}  txt={m['txt']!r}")
        for oc in m["onclicks"]:
            print(f"        onclick= {oc}")

    print("\n  ---- pane text เต็ม (เอกสารตอบรับ) ----")
    print("  " + last_txt.replace("\n", "\n  ")[:2000])

    # ---- outerHTML แต่ละ 'แถวเอกสาร' จาก #DetailDocumentList โดยตรง (ลูกโดยตรงของ container) ----
    rows_html = page.evaluate(
        r"""() => {
            const cont = document.querySelector('#DetailDocumentList');
            if (!cont) return [];
            return [...cont.children].map(el => (el.outerHTML || '').slice(0, 2500));
        }"""
    )
    print(f"\n  [#DetailDocumentList children outerHTML] -> {len(rows_html)} แถว")
    for i, h in enumerate(rows_html, start=1):
        print(f"\n  --- แถว {i} ---")
        print("  " + h.replace("\n", "\n  "))

    # ---- ทุก element ที่คลิกได้ทั้งแท็บ (ดูภาพรวม) ----
    clickables = page.evaluate(
        r"""(sel) => {
            const pane = document.querySelector(sel) || document.querySelector('.tab-pane.active') || document;
            return [...pane.querySelectorAll('a,button,[onclick],[href]')]
                .map(el => ({tag:el.tagName.toLowerCase(),
                             txt:(el.innerText||el.textContent||'').replace(/\s+/g,' ').trim().slice(0,70),
                             onclick:(el.getAttribute('onclick')||'').slice(0,140)}))
                .filter(o => o.onclick);
        }""",
        pane_sel,
    )
    print(f"\n  [ทุก element ที่มี onclick ในแท็บ] -> {len(clickables)}")
    for o in clickables:
        print(f"    <{o['tag']}> '{o['txt']}'  onclick={o['onclick']}")


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
                 "user_type": cfg["user_type"], "method": cfg["method"]}
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
