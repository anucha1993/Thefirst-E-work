"""Test SS record — click tab_default_5 explicitly + dump all iframes."""
from __future__ import annotations
import json, os, sys
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import login, ROOT, goto_tracking, apply_wa_filter, collect_all_wa_rows, build_detail_url


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }
    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, args=["--start-maximized"])
        ctx = b.new_context(locale="th-TH", viewport={"width": 1600, "height": 1000})
        page = ctx.new_page()
        try:
            login(page, cfg); page.wait_for_timeout(1200)
            goto_tracking(page); page.wait_for_timeout(1500)
            apply_wa_filter(page, "", status_ids=["SS"]); page.wait_for_timeout(1500)
            rows = collect_all_wa_rows(page, log=lambda *a: None)
            ss = [r for r in rows if str(r.get("status") or "").upper() == "SS"]
            tgt = ss[0]
            print(f"SS target: {tgt.get('reqNo')}")
            page.goto(build_detail_url(tgt), wait_until="domcontentloaded", timeout=25_000)
            page.wait_for_timeout(3000)

            # Click tab 5 explicitly
            print("\n=== click tab_default_5 ===")
            tab_link = page.locator('a[href="#tab_default_5"]').first
            tab_link.click(timeout=8000)
            page.wait_for_timeout(5000)

            # Dump ALL iframes
            print("\n=== all iframes on page ===")
            frames = page.evaluate(r"""
                () => [...document.querySelectorAll('iframe')].map(f => ({
                    id: f.id, name: f.name, src: (f.src||'').slice(0, 250),
                    parent_tab: (f.closest('.tab-pane')||{}).id || '',
                    visible: !!f.offsetParent,
                    w: f.offsetWidth, h: f.offsetHeight,
                }))
            """)
            for fr in frames:
                print(f"  · id={fr['id']} vis={fr['visible']} tab={fr['parent_tab']} {fr['w']}x{fr['h']}")
                print(f"    src={fr['src']}")

            # Look inside tab_default_5 content
            print("\n=== tab_default_5 inner HTML sample ===")
            content = page.evaluate(r"""
                () => {
                    const tab = document.getElementById('tab_default_5');
                    if (!tab) return null;
                    return {
                        html: tab.innerHTML.slice(0, 2000),
                        buttons: [...tab.querySelectorAll('button, a.btn')].map(b => ({
                            text: (b.textContent||'').replace(/\s+/g,' ').trim().slice(0,60),
                            id: b.id, onclick: (b.getAttribute('onclick')||'').slice(0,120),
                        })),
                    };
                }
            """)
            if content:
                print(f"buttons in tab: {len(content['buttons'])}")
                for b_ in content["buttons"][:10]:
                    print(f"  · '{b_['text']}' id={b_['id']} onclick={b_['onclick']}")
                print(f"\nHTML sample:\n{content['html'][:1500]}")

            try: input("\nEnter to close ...")
            except Exception: pass
        finally:
            ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
