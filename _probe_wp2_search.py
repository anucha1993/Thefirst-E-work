"""Probe: ค้นหา 'WP2' ในหน้าเว็บ (DOM, scripts, ajax) เพื่อหาว่ามันอยู่ไหน"""
from __future__ import annotations
import os
import sys
import json
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()

from scrape_wa import login, goto_tracking  # noqa: E402


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
        "headless": False,
    }
    ajax_hits = []
    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False,
                               args=["--ignore-certificate-errors", "--start-maximized"])
        ctx = b.new_context(locale="th-TH", ignore_https_errors=True,
                            viewport={"width": 1600, "height": 1000})
        page = ctx.new_page()

        # capture ajax requests to /Etracking or /Tracking APIs
        def on_req(req):
            u = req.url or ""
            if "tracking" in u.lower() or "etracking" in u.lower():
                try:
                    post = req.post_data or ""
                except Exception:
                    post = ""
                ajax_hits.append({"url": u, "method": req.method, "post_data": post[:500]})
        page.on("request", on_req)

        login(page, cfg)
        goto_tracking(page)
        page.wait_for_timeout(5000)

        # ---- 1) grep 'WP2' / 'wp2' ในทั้ง page (DOM + scripts) ----
        matches = page.evaluate(
            r"""() => {
              const out = {domHits: [], scriptSnippets: [], statusIdsInJs: []};
              // 1a) DOM: element ที่มี WP2 ใน text/attribute
              const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_ELEMENT, null);
              let node;
              while ((node = walker.nextNode())) {
                const oc = (node.getAttribute && node.getAttribute('onclick')) || '';
                const id = node.id || '';
                const nm = node.getAttribute && node.getAttribute('name') || '';
                const val = node.getAttribute && node.getAttribute('value') || '';
                const cls = node.className || '';
                const txt = (node.innerText || '').slice(0, 50);
                if (/WP2/i.test(oc) || /WP2/i.test(id) || /WP2/i.test(nm) || /WP2/i.test(val) || /WP2/i.test(cls)) {
                  out.domHits.push({
                    tag: node.tagName, id, name: nm, value: val, className: cls.slice(0, 100),
                    onclick: oc.slice(0, 200), text: txt,
                  });
                  if (out.domHits.length >= 30) return out;
                }
              }
              // 1b) grep in inline script bodies
              document.querySelectorAll('script:not([src])').forEach(s => {
                const t = s.textContent || '';
                if (/WP2/i.test(t)) {
                  // เก็บ snippet 200 chars รอบ 'WP2'
                  const idx = t.search(/WP2/i);
                  out.scriptSnippets.push(t.slice(Math.max(0, idx-100), idx+200));
                  if (out.scriptSnippets.length >= 10) return;
                }
              });
              // 1c) หา GetDataRequestFormEtrackingForAlien source เต็ม + ฟังก์ชันเกี่ยวข้อง
              try {
                if (typeof GetDataRequestFormEtrackingForAlien === 'function') {
                  out.statusIdsInJs.push({
                    fn: 'GetDataRequestFormEtrackingForAlien',
                    src: GetDataRequestFormEtrackingForAlien.toString().slice(0, 6000),
                  });
                }
              } catch(e) {}
              return out;
            }"""
        )
        Path("_probe_wp2_search.json").write_text(
            json.dumps({"domHits": matches.get("domHits", []),
                        "scriptSnippets": matches.get("scriptSnippets", []),
                        "ajaxRequests": ajax_hits,
                        "GetDataFn": (matches.get("statusIdsInJs") or [{}])[0].get("src", "")},
                       ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(f"[✓] dumped → _probe_wp2_search.json")
        print(f"    DOM hits: {len(matches.get('domHits', []))}")
        print(f"    script snippets: {len(matches.get('scriptSnippets', []))}")
        print(f"    ajax requests captured: {len(ajax_hits)}")
        # print DOM hits
        for h in matches.get("domHits", [])[:10]:
            print(f"      · tag={h['tag']} id={h['id']!r} value={h['value']!r} text={h['text']!r}")
            if h.get("onclick"):
                print(f"        onclick: {h['onclick']}")
        # print ajax URLs
        print("\n  Ajax URLs captured (first 5):")
        for a in ajax_hits[:5]:
            print(f"    {a['method']} {a['url'][:120]}")
            if a['post_data']:
                print(f"      post: {a['post_data'][:200]}")

        page.wait_for_timeout(15000)
        ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
