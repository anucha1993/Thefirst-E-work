"""Probe เข้าไปใน iframe #link_appointment เพื่อหา ใบนัดหมาย / ปุ่มพิมพ์
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
OUT_JSON = ROOT / "_probe_appt_iframe_out.json"
SHOTS = ROOT / "screenshots" / "appt_iframe"
SHOTS.mkdir(parents=True, exist_ok=True)


IFRAME_DUMP_JS = r"""() => {
    const vis = el => !!(el && el.offsetParent !== null);
    const txt = e => (e && e.textContent || '').replace(/\s+/g,' ').trim();
    return {
        url: location.href,
        title: document.title,
        bodyTextFirst: txt(document.body).slice(0, 2500),
        buttons: Array.from(document.querySelectorAll('button, a, [onclick], [role="button"]'))
            .filter(vis)
            .map(e => ({
                tag: e.tagName,
                text: txt(e).slice(0, 100),
                href: e.getAttribute('href') || '',
                onclick: (e.getAttribute('onclick') || '').slice(0, 250),
                id: e.id || '',
                cls: (e.className || '').toString().slice(0, 100),
            }))
            .filter(x => x.text || x.onclick || x.href),
        headings: Array.from(document.querySelectorAll('h1,h2,h3,h4,h5,h6,legend,label,strong'))
            .filter(vis).map(txt).filter(t => t.length > 2 && t.length < 200).slice(0, 30),
        images: Array.from(document.querySelectorAll('img'))
            .filter(vis).map(i => ({src: (i.src||'').slice(-100), alt: i.alt || ''}))
            .slice(0, 10),
    };
}"""


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }
    result = {"target_url": TARGET_URL}

    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, args=["--start-maximized"])
        ctx = b.new_context(locale="th-TH", viewport={"width": 1600, "height": 1000},
                            accept_downloads=True)
        page = ctx.new_page()
        try:
            login(page, cfg)
            page.wait_for_timeout(1500)
            page.goto(TARGET_URL, wait_until="domcontentloaded", timeout=45_000)
            page.wait_for_timeout(4000)

            # Trigger tab การนัดหมาย
            page.evaluate("""() => {
                if (window.jQuery) jQuery('a[href="#tab_default_5"]').tab('show');
                const a = document.querySelector('a[href="#tab_default_5"]');
                if (a) a.click();
            }""")
            page.wait_for_timeout(3000)

            # หา iframe #link_appointment
            iframe_info = page.evaluate(r"""() => {
                const f = document.querySelector('#link_appointment') 
                       || document.querySelector('iframe.iframe_show')
                       || document.querySelector('iframe[src*="queue"]');
                if (!f) return null;
                return {
                    id: f.id, src: f.src || '', name: f.name || '',
                    width: f.getBoundingClientRect().width,
                    height: f.getBoundingClientRect().height,
                };
            }""")
            print(f"[iframe] {iframe_info}")
            result["iframe_info"] = iframe_info

            if not iframe_info:
                print("[!] ไม่พบ iframe — record นี้อาจไม่มีใบนัดหมาย")
                return 1

            # รอ iframe โหลด + dump content
            page.wait_for_timeout(6000)
            frames = page.frames
            print(f"[frames] {len(frames)} frames on page:")
            for i, fr in enumerate(frames):
                print(f"  [{i}] name={fr.name!r} url={fr.url[:100]}")

            # หา frame ที่ src มี "queue" / "bookingdate"
            appt_frame = None
            for fr in frames:
                if "queue" in fr.url or "bookingdate" in fr.url:
                    appt_frame = fr
                    break
            if not appt_frame:
                print("[!] ไม่พบ frame ที่ src มี queue/bookingdate")
                return 1

            print(f"\n[+] appointment frame: {appt_frame.url[:150]}")
            # รอ frame load เพิ่ม
            try:
                appt_frame.wait_for_load_state("domcontentloaded", timeout=15_000)
            except Exception:
                pass
            page.wait_for_timeout(3000)

            page.screenshot(path=str(SHOTS / "01_iframe_page.png"), full_page=True)

            # dump content ใน frame
            try:
                data = appt_frame.evaluate(IFRAME_DUMP_JS)
            except Exception as e:
                print(f"[!] eval error: {e}")
                data = {"error": str(e)}
            result["iframe_content"] = data

            print("\n=== IFRAME CONTENT ===")
            print(f"URL: {data.get('url','')[:150]}")
            print(f"Title: {data.get('title','')}")
            print(f"\nbodyText (first 1500):")
            print("  " + (data.get("bodyTextFirst") or "")[:1500].replace("\n", "\n  "))

            hs = data.get("headings") or []
            print(f"\nheadings ({len(hs)}):")
            for h in hs[:20]:
                print(f"  · {h[:150]}")

            bts = data.get("buttons") or []
            print(f"\nbuttons/links ({len(bts)}):")
            for it in bts[:40]:
                print(f"  · <{it['tag']}> text={it['text'][:70]!r:75}  ")
                if it.get("onclick"):
                    print(f"      onclick={it['onclick'][:150]}")
                if it.get("href"):
                    print(f"      href={it['href'][:120]}")

            imgs = data.get("images") or []
            if imgs:
                print(f"\nimages ({len(imgs)}):")
                for im in imgs:
                    print(f"  · alt={im['alt']!r}  src=...{im['src']}")

            OUT_JSON.write_text(
                json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            print(f"\n[done] → {OUT_JSON}")

            input("\n Enter to close ...")
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
