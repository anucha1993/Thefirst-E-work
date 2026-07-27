"""เจาะจง 1 record ที่ user บอกว่ามีใบนัดหมายจริง — dump ทุก tab-pane content แบบละเอียด
ให้เวลารอ AJAX เยอะขึ้น + trigger tab ด้วย jQuery หลายวิธี
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

TARGET_URL = os.getenv(
    "PROBE_APPT_URL",
    "https://eworkpermit.doe.go.th/RenewMOU/DetailFormRenewMOU?user_id=303816&status=APSS&group_id=69123000059294&form_type=MT_59_MOU_RENEWAL",
)

OUT_JSON = ROOT / "_probe_appt_detail_out.json"
SHOTS = ROOT / "screenshots" / "appt_detail"
SHOTS.mkdir(parents=True, exist_ok=True)


DUMP_ALL_PANES_JS = r"""() => {
    // dump content ของ tab-pane ทุกอัน (id=tab_default_1..N)
    const txt = e => (e && e.textContent || '').replace(/\s+/g,' ').trim();
    const vis = el => !!(el && el.offsetParent !== null);
    const panes = Array.from(document.querySelectorAll('.tab-pane'));
    return panes.map(p => {
        const active = p.classList.contains('active') || p.classList.contains('in') || vis(p);
        const inner = (p.innerHTML || '').length;
        const t = txt(p);
        const items = Array.from(p.querySelectorAll('a, button, [onclick]'))
            .map(e => ({
                tag: e.tagName,
                text: txt(e).slice(0, 100),
                href: e.getAttribute('href') || '',
                onclick: (e.getAttribute('onclick') || '').slice(0, 300),
                id: e.id || '',
                cls: (e.className || '').toString().slice(0, 100),
                vis: vis(e),
            }))
            .filter(x => x.onclick || x.href || x.text);
        return {
            id: p.id, cls: (p.className || '').toString().slice(0, 100),
            active, htmlLen: inner, textLen: t.length,
            textFirst200: t.slice(0, 400),
            items,
        };
    });
}"""


TRIGGER_APPT_TAB_JS = r"""() => {
    // ลอง trigger tab การนัดหมาย ทุกวิธี
    const results = [];

    // 1. คลิก anchor ที่ href="#tab_default_5"
    const a1 = document.querySelector('a[href="#tab_default_5"]');
    if (a1) { a1.click(); results.push('clicked #tab_default_5 anchor'); }

    // 2. jQuery tab('show')
    if (window.jQuery) {
        try {
            jQuery('a[href="#tab_default_5"]').tab('show');
            results.push('jQuery tab show');
        } catch(e) { results.push('jQuery tab err: ' + e.message.slice(0,60)); }
    }

    // 3. หา element ที่มีข้อความ 'การนัดหมาย' แล้วคลิก
    const els = Array.from(document.querySelectorAll('a, .nav-link, .nav-tabs a, .nav a, [role="tab"]'));
    const t = els.find(a => /^\s*การนัดหมาย\s*$/.test((a.textContent||'').trim()));
    if (t) { t.click(); results.push('clicked การนัดหมาย text'); }

    return results;
}"""


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }
    print(f"[i] target: {TARGET_URL}")

    result = {"url": TARGET_URL}
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
            # capture ทุก network request เกี่ยวกับ tab นัดหมาย
            captured = []

            def on_req(req):
                u = req.url or ""
                if any(k in u for k in ("Appointment", "Appt", "appointment", "GetAppoint",
                                         "GetDocumentConfirm", "GetSummary", "GetForm")):
                    captured.append({"method": req.method, "url": u[:220]})

            page.on("request", on_req)

            login(page, cfg)
            page.wait_for_timeout(1500)

            page.goto(TARGET_URL, wait_until="domcontentloaded", timeout=45_000)
            page.wait_for_timeout(4000)
            print(f"[+] arrived: {page.url}")

            # Trigger tab การนัดหมาย
            trig = page.evaluate(TRIGGER_APPT_TAB_JS)
            print(f"\n[trigger results]: {trig}")

            # รอ AJAX นานขึ้น
            for wait_s in (3, 3, 3):
                page.wait_for_timeout(wait_s * 1000)
                page.screenshot(path=str(SHOTS / f"after_{wait_s}s.png"), full_page=True)

            # dump ทุก pane
            panes = page.evaluate(DUMP_ALL_PANES_JS)
            result["panes"] = panes
            result["network_captured"] = captured

            print(f"\n=== all tab-panes ({len(panes)}) ===")
            for p in panes:
                marker = "★" if p["active"] else " "
                print(f"  {marker} #{p['id']:20}  active={p['active']}  "
                      f"htmlLen={p['htmlLen']:6}  textLen={p['textLen']:5}  "
                      f"items={len(p['items'])}")
                if p["textLen"] > 5:
                    print(f"       first: {p['textFirst200'][:180]}")

            # โฟกัส pane 5 (การนัดหมาย)
            pane5 = next((p for p in panes if p["id"] == "tab_default_5"), None)
            if pane5:
                print("\n\n=== #tab_default_5 detail ===")
                print(f"active={pane5['active']}  htmlLen={pane5['htmlLen']}  "
                      f"textLen={pane5['textLen']}")
                print(f"\nfull text (first 2000 chars):")
                print("  " + pane5["textFirst200"][:2000].replace("\n", "\n  "))
                print(f"\nitems ({len(pane5['items'])}):")
                for it in pane5["items"]:
                    print(f"  · <{it['tag']}> vis={it['vis']}  text={it['text'][:60]!r}")
                    if it["onclick"]:
                        print(f"      onclick={it['onclick'][:200]}")
                    if it["href"]:
                        print(f"      href={it['href'][:120]}")

            print(f"\n=== network requests to Appointment* / GetDocument* ===")
            for c in captured[:30]:
                print(f"  {c['method']} {c['url']}")

            OUT_JSON.write_text(
                json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            print(f"\n[done] → {OUT_JSON}")

            input("\nกด Enter เพื่อปิด...")
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
