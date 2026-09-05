"""วินิจฉัย DOM ของหน้า detail CHANGE_44 — ทำไม pane อ่าน innerText ว่าง
ไม่พิมพ์รหัสผ่าน/ข้อมูลลับ
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import scrape_wa as s  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

URL = "https://eworkpermit.doe.go.th/Forms/DetailFormRequest44/69144400629422?id=RA17662449023455621"
OUT = s.ROOT / "_probe_change44_dom_out.json"

DIAG_JS = r"""() => {
  const panes = ['tab-status', 'tab-request', 'tab-alien'];
  const norm = s => (s || '').replace(/\s+/g, ' ').trim();
  const res = {};
  // แท็บลิงก์ทั้งหมด
  res.tabLinks = Array.from(document.querySelectorAll('a[data-toggle],a[data-bs-toggle],[role=tab]'))
    .slice(0, 20)
    .map(a => ({
      text: norm(a.innerText),
      href: a.getAttribute('href') || '',
      dataToggle: a.getAttribute('data-toggle') || a.getAttribute('data-bs-toggle') || '',
      target: a.getAttribute('data-target') || a.getAttribute('data-bs-target') || '',
      tag: a.tagName,
    }));
  res.iframeCountDoc = document.querySelectorAll('iframe').length;
  res.panes = {};
  for (const pid of panes) {
    const p = document.getElementById(pid);
    if (!p) { res.panes[pid] = {exists: false}; continue; }
    const cs = getComputedStyle(p);
    res.panes[pid] = {
      exists: true,
      className: p.className,
      display: cs.display,
      visibility: cs.visibility,
      offsetHeight: p.offsetHeight,
      innerTextLen: (p.innerText || '').length,
      innerHTMLLen: (p.innerHTML || '').length,
      textContentLen: (p.textContent || '').length,
      iframeCount: p.querySelectorAll('iframe').length,
      headStep: p.querySelectorAll('.head-step').length,
      labelFormInfo: p.querySelectorAll('.label-form-info').length,
      formInfo: p.querySelectorAll('.form-info').length,
      innerHTMLHead: (p.innerHTML || '').slice(0, 1200),
    };
  }
  return res;
}"""


def dump_state(page, tag: str) -> dict:
    try:
        return page.evaluate(DIAG_JS)
    except Exception as e:
        return {"error": str(e)[:200]}


def main() -> int:
    cfg = s.load_config(require_login=True)
    cfg["headless"] = False
    out: dict = {"url": URL}

    with sync_playwright() as pw:
        browser = s._launch_chromium(pw, cfg, ["--disable-blink-features=AutomationControlled"])
        ctx = browser.new_context(locale="th-TH", timezone_id="Asia/Bangkok",
                                  viewport={"width": 1500, "height": 900})
        page = ctx.new_page()
        try:
            print("[1] login...")
            s.login(page, cfg)
            print("[2] goto detail URL ตรง ๆ ...")
            page.goto(URL, wait_until="domcontentloaded", timeout=25000)
            page.wait_for_timeout(1500)
            print("    landed:", page.url)
            out["landedUrl"] = page.url

            out["before_click"] = dump_state(page, "before")
            print("[3] before click — panes:")
            for pid, st in out["before_click"].get("panes", {}).items():
                print(f"    {pid}: {json.dumps(st, ensure_ascii=False)[:220]}")

            # คลิกแท็บ 'คำขออนุญาต' (#tab-request) แล้วดูใหม่
            print("[4] คลิก #tab-request ...")
            page.evaluate("() => { const a = document.querySelector('a[href=\"#tab-request\"]'); if (a) a.click(); }")
            page.wait_for_timeout(1500)
            out["after_click_request"] = dump_state(page, "after")
            st = out["after_click_request"].get("panes", {}).get("tab-request", {})
            print(f"    tab-request หลังคลิก: {json.dumps(st, ensure_ascii=False)[:400]}")

            # แสดง tab links
            print("[5] tab links:")
            for tl in out["before_click"].get("tabLinks", [])[:12]:
                print(f"    {json.dumps(tl, ensure_ascii=False)}")
        finally:
            OUT.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"[เสร็จ] เขียน {OUT.name}")
            ctx.close()
            browser.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
