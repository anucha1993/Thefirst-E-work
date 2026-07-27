"""Probe DOM ของแท็บ 'การนัดหมาย' บนหน้า detail คำขอ MoU renewal
Target URL: https://eworkpermit.doe.go.th/RenewMOU/DetailFormRenewMOU?user_id=303816&status=APSS&group_id=69123000038421&form_type=MT_59_MOU_RENEWAL
ไม่ hardcode URL — user แก้ path/params ผ่าน env ได้
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

from scrape_wa import login, _read_login_accounts, ROOT  # noqa: E402

TARGET_URL = os.getenv(
    "PROBE_APPT_URL",
    "https://eworkpermit.doe.go.th/RenewMOU/DetailFormRenewMOU?user_id=303816&status=APSS&group_id=69123000038421&form_type=MT_59_MOU_RENEWAL",
)

OUT_JSON = ROOT / "_probe_appt_tab_out.json"
SHOTS = ROOT / "screenshots" / "appt_tab"
SHOTS.mkdir(parents=True, exist_ok=True)


DUMP_TABS_JS = r"""() => {
    // เก็บทุก tab บนหน้า detail (พอที่จะเห็นชื่อ + href/onclick + visible)
    const vis = el => !!(el && el.offsetParent !== null);
    const txt = e => (e.textContent || '').replace(/\s+/g,' ').trim();
    const tabs = Array.from(document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a, [role="tab"]'))
        .filter(vis)
        .map(e => ({
            text: txt(e).slice(0, 60),
            href: e.getAttribute('href') || '',
            id: e.id || '',
            cls: (e.className || '').toString().slice(0, 100),
            onclick: (e.getAttribute('onclick') || '').slice(0, 120),
        }))
        .filter(t => t.text);
    return { url: location.href, tabs };
}"""


DUMP_APPT_PANE_JS = r"""() => {
    // dump: pane ที่ active หลังคลิกแท็บ 'การนัดหมาย'
    // - หา element ที่มี label/heading เกี่ยวกับ 'ใบนัดหมาย' / 'นัดหมาย' / 'appointment'
    // - list ลิงก์/ปุ่ม onclick ที่เกี่ยวกับ download PDF (GetDocumentConfirm / GetAppointment* / window.open ฯลฯ)
    const vis = el => !!(el && el.offsetParent !== null);
    const norm = s => (s || '').replace(/\s+/g, ' ').trim();
    const txt = e => norm(e.textContent || '');

    // หา tab-pane ที่ active หรือ visible
    const panes = Array.from(document.querySelectorAll('.tab-pane, [role="tabpanel"], .card, section'))
        .filter(vis);
    // ใช้ pane ที่มี "นัดหมาย" ในข้อความมากที่สุด
    let target = null;
    let bestScore = 0;
    for (const p of panes) {
        const t = txt(p);
        const score = (t.match(/นัดหมาย/g) || []).length;
        if (score > bestScore && t.length < 8000) {
            bestScore = score;
            target = p;
        }
    }
    const scope = target || document.body;

    // เก็บ headings
    const headings = Array.from(scope.querySelectorAll('h1,h2,h3,h4,h5,h6,legend,.card-header,.panel-heading,strong,b,label'))
        .filter(vis).map(txt).filter(t => t.length > 2 && t.length < 200).slice(0, 40);

    // ลิงก์/ปุ่ม download — เก็บทุกอันที่มี onclick หรือ href
    const downloadables = Array.from(scope.querySelectorAll('a, button, [onclick]'))
        .filter(vis)
        .map(e => ({
            tag: e.tagName,
            text: txt(e).slice(0, 100),
            href: e.getAttribute('href') || '',
            onclick: (e.getAttribute('onclick') || '').slice(0, 200),
            id: e.id || '',
            cls: (e.className || '').toString().slice(0, 100),
        }))
        .filter(x => x.onclick || x.href);

    // สาระของ pane ทั้งหมด (ให้ user อ่านตาได้ง่าย)
    const paneText = txt(scope).slice(0, 3000);

    return {
        paneFound: !!target,
        headings,
        downloadables,
        paneText,
    };
}"""


def main() -> int:
    # อ่านบัญชีจาก .env
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }
    if not cfg["username"] or not cfg["password"]:
        print("[!] ตั้ง EWP_USERNAME / EWP_PASSWORD ใน .env ก่อน")
        return 1

    print(f"[i] login as: {cfg['username']}  ({cfg['user_type']})")
    print(f"[i] target: {TARGET_URL}")

    result: dict = {"target_url": TARGET_URL, "login_user": cfg["username"]}

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
            print(f"[+] login ok  |  cur url: {page.url}")

            page.goto(TARGET_URL, wait_until="domcontentloaded", timeout=45_000)
            page.wait_for_timeout(3500)
            print(f"[+] arrived: {page.url}")
            page.screenshot(path=str(SHOTS / "01_arrived.png"), full_page=True)

            # dump ทุก tab บนหน้า detail
            tabs_info = page.evaluate(DUMP_TABS_JS)
            result["tabs_on_page"] = tabs_info
            print(f"\n=== tabs on page ({len(tabs_info['tabs'])}) ===")
            for t in tabs_info["tabs"]:
                print(f"  · text={t['text']!r:40}  href={t['href']!r:35}  onclick={t['onclick'][:60]!r}")

            # คลิก tab 'การนัดหมาย'
            clicked = page.evaluate(
                r"""() => {
                    const vis = el => !!(el && el.offsetParent !== null);
                    const els = Array.from(document.querySelectorAll('a, .nav-link, .nav-tabs a, .nav a, [role="tab"], button'))
                        .filter(vis);
                    const f = els.find(a => /การนัดหมาย/.test((a.textContent||'').trim()));
                    if (f) { f.click(); return (f.textContent||'').trim(); }
                    return '';
                }"""
            )
            print(f"\n[i] clicked tab: {clicked!r}")
            page.wait_for_timeout(2500)
            page.screenshot(path=str(SHOTS / "02_after_click_appt_tab.png"), full_page=True)

            # dump content ของ pane การนัดหมาย
            appt_info = page.evaluate(DUMP_APPT_PANE_JS)
            result["appt_tab"] = appt_info
            print("\n=== 'การนัดหมาย' pane ===")
            print(f"  paneFound: {appt_info['paneFound']}")
            print(f"\n  headings ({len(appt_info['headings'])}):")
            for h in appt_info["headings"][:20]:
                print(f"    · {h[:150]}")
            print(f"\n  downloadables ({len(appt_info['downloadables'])}):")
            for d in appt_info["downloadables"][:30]:
                print(
                    f"    · <{d['tag']}> text={d['text'][:60]!r:65}"
                    f"  onclick={d['onclick'][:100]!r}"
                )
                if d.get("href"):
                    print(f"        href={d['href'][:120]}")
            print("\n  paneText (first 800 chars):")
            print("    " + appt_info["paneText"][:800].replace("\n", "\n    "))

            OUT_JSON.write_text(
                json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            print(f"\n[done] dump → {OUT_JSON}")

            input("\nกด Enter เพื่อปิดเบราว์เซอร์...")
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
