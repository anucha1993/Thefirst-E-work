"""Scan account records เพื่อหา 1 record ที่ tab 'การนัดหมาย' มีเอกสารจริง
แล้ว dump DOM ของ pane นั้น
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

from scrape_wa import (  # noqa: E402
    login, goto_tracking, apply_wa_filter, collect_all_wa_rows,
    build_detail_url, ROOT,
)

OUT_JSON = ROOT / "_probe_appt_scan_out.json"
SHOTS = ROOT / "screenshots" / "appt_scan"
SHOTS.mkdir(parents=True, exist_ok=True)

# JS: กด tab การนัดหมาย + อ่านเนื้อ pane
APPT_DUMP_JS = r"""() => {
    const vis = el => !!(el && el.offsetParent !== null);
    const txt = e => (e.textContent || '').replace(/\s+/g,' ').trim();

    // คลิก tab การนัดหมาย
    const nav = Array.from(document.querySelectorAll('a, .nav-link, .nav-tabs a, .nav a, [role="tab"]'))
        .filter(vis);
    const t = nav.find(a => /^\s*การนัดหมาย\s*$/.test((a.textContent||'').trim()));
    if (t) t.click();

    // รอ pane โผล่ (async) — return แล้วให้ Python รอต่อ
    return { clicked: !!t, tabTxt: t ? (t.textContent||'').trim() : '' };
}"""

APPT_READ_JS = r"""() => {
    const vis = el => !!(el && el.offsetParent !== null);
    const txt = e => (e.textContent || '').replace(/\s+/g,' ').trim();

    // หา pane #tab_default_5 หรือที่ active
    const pane = document.querySelector('#tab_default_5')
                 || Array.from(document.querySelectorAll('.tab-pane'))
                        .filter(p => p.classList.contains('active') || vis(p))
                        .find(p => /นัดหมาย/.test(p.innerText || ''));
    if (!pane) return { paneFound: false };

    const paneText = txt(pane).slice(0, 4000);
    const hasEmpty = /ไม่มีการนัดหมาย|no appointment/i.test(paneText);

    // downloads/links ใน pane
    const items = Array.from(pane.querySelectorAll('a, button, [onclick]'))
        .filter(vis)
        .map(e => ({
            tag: e.tagName,
            text: txt(e).slice(0, 100),
            href: e.getAttribute('href') || '',
            onclick: (e.getAttribute('onclick') || '').slice(0, 250),
            id: e.id || '',
            cls: (e.className || '').toString().slice(0, 100),
        }))
        .filter(x => x.onclick || x.href || x.text);

    // headings
    const headings = Array.from(pane.querySelectorAll('h1,h2,h3,h4,h5,h6,legend,.card-header,strong,b,label'))
        .filter(vis).map(txt).filter(t => t.length > 2 && t.length < 200).slice(0, 30);

    // tables
    const tables = Array.from(pane.querySelectorAll('table'))
        .filter(vis)
        .map(t => ({
            rows: Array.from(t.querySelectorAll('tr')).slice(0, 10).map(r =>
                Array.from(r.querySelectorAll('th, td')).map(c => txt(c).slice(0, 100))
            ),
        }));

    return {
        paneFound: true,
        isEmpty: hasEmpty,
        paneText,
        items,
        headings,
        tables,
    };
}"""


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }
    if not cfg["username"] or not cfg["password"]:
        print("[!] ตั้ง EWP_USERNAME / EWP_PASSWORD ใน .env ก่อน")
        return 1

    print(f"[i] login as: {cfg['username']}")

    result: dict = {"login_user": cfg["username"], "scanned": []}

    with sync_playwright() as pw:
        b = pw.chromium.launch(
            headless=False,
            args=["--ignore-certificate-errors", "--start-maximized"],
        )
        ctx = b.new_context(
            locale="th-TH", ignore_https_errors=True,
            viewport={"width": 1600, "height": 1000},
            accept_downloads=True,
        )
        page = ctx.new_page()
        try:
            login(page, cfg)
            page.wait_for_timeout(1500)

            # หา records สถานะ AP + SS (สถานะที่น่ามีใบนัดหมาย)
            goto_tracking(page)
            page.wait_for_timeout(2000)
            apply_wa_filter(page, "", status_ids=["AP", "SS"])
            rows = collect_all_wa_rows(page, log=print)
            print(f"\n[+] collected {len(rows)} rows (AP + SS)")

            # ลอง scan max 30 record แรก — หาที่ pane 'การนัดหมาย' ไม่ empty
            MAX_SCAN = 30
            found_url = None
            for i, r in enumerate(rows[:MAX_SCAN], 1):
                req_no = r.get("reqNo") or ""
                status = r.get("status") or ""
                url = build_detail_url(r)
                print(f"\n[{i:2d}/{min(MAX_SCAN, len(rows))}] {req_no}  status={status}")
                if not url:
                    print("    (no url)"); continue

                page.goto(url, wait_until="domcontentloaded", timeout=25_000)
                page.wait_for_timeout(2200)

                # คลิก tab การนัดหมาย
                click_res = page.evaluate(APPT_DUMP_JS)
                page.wait_for_timeout(1800)  # รอ pane โหลด
                info = page.evaluate(APPT_READ_JS)
                is_empty = info.get("isEmpty") if info.get("paneFound") else True
                items_ct = len(info.get("items") or [])
                headings_ct = len(info.get("headings") or [])
                pane_len = len(info.get("paneText") or "")
                print(f"    paneFound={info.get('paneFound')}  isEmpty={is_empty}  "
                      f"items={items_ct}  headings={headings_ct}  paneLen={pane_len}")

                summary = {
                    "req_no": req_no, "status": status, "url": url,
                    "isEmpty": is_empty, "items_ct": items_ct,
                    "headings_ct": headings_ct, "paneLen": pane_len,
                }
                result["scanned"].append(summary)

                if info.get("paneFound") and not is_empty and (items_ct > 0 or pane_len > 100):
                    found_url = url
                    result["FIRST_NONEMPTY"] = {
                        "req_no": req_no, "status": status, "url": url,
                        "appt_pane": info,
                    }
                    page.screenshot(path=str(SHOTS / f"FOUND_{req_no}.png"), full_page=True)
                    print(f"    ✓ FOUND — ใบนัดหมายจริง ที่ {req_no}")
                    print(f"\n=== pane content ===")
                    print(f"paneText ({pane_len} chars):")
                    print("  " + info["paneText"][:1200].replace("\n", "\n  "))
                    print(f"\nheadings ({headings_ct}):")
                    for h in info["headings"]:
                        print(f"  · {h[:150]}")
                    print(f"\nitems ({items_ct}):")
                    for it in info["items"]:
                        print(f"  · <{it['tag']}> text={it['text'][:60]!r}  "
                              f"onclick={it['onclick'][:120]!r}")
                        if it.get("href"):
                            print(f"      href={it['href'][:120]}")
                    if info.get("tables"):
                        print(f"\ntables ({len(info['tables'])}):")
                        for ti, t in enumerate(info["tables"]):
                            for ri, row in enumerate(t["rows"]):
                                print(f"  [t{ti}.r{ri}] {row}")
                    break

            OUT_JSON.write_text(
                json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            print(f"\n[done] dump → {OUT_JSON}")
            if not found_url:
                print(f"[!] scanned {MAX_SCAN} records — ไม่เจอ pane ที่มีใบนัดหมาย")

            input("\nกด Enter เพื่อปิดเบราว์เซอร์...")
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
