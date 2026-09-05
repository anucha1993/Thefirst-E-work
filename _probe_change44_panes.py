"""ดัมพ์ .tab-pane + nav links จริงของหน้า CHANGE_44 (main frame)
เพื่อดูว่า pane ใช้ id/class อะไร และแท็บไหนคือ สถานะคำขอ/คำขออนุญาต
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
OUT = s.ROOT / "_probe_change44_panes_out.json"

DUMP_JS = r"""() => {
  const norm = s => (s || '').replace(/\s+/g, ' ').trim();
  const navLinks = Array.from(document.querySelectorAll('.nav-tabs a, .nav a, a[data-toggle="tab"], a[role="tab"]'))
    .map(a => ({
      text: norm(a.innerText),
      href: a.getAttribute('href') || '',
      ariaControls: a.getAttribute('aria-controls') || '',
      dataTarget: a.getAttribute('data-target') || a.getAttribute('data-bs-target') || '',
      cls: a.className,
    }));
  const panes = Array.from(document.querySelectorAll('.tab-pane')).map((p, i) => ({
    index: i,
    id: p.id || '',
    cls: p.className,
    ariaLabelledby: p.getAttribute('aria-labelledby') || '',
    innerTextLen: (p.innerText || '').length,
    innerTextHead: norm(p.innerText).slice(0, 220),
    headStep: Array.from(p.querySelectorAll('.head-step')).map(h => norm(h.innerText)).slice(0, 8),
  }));
  return { navLinks, panes };
}"""


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
            print("[2] goto detail URL ...")
            page.goto(URL, wait_until="domcontentloaded", timeout=25000)
            page.wait_for_timeout(2000)
            data = page.evaluate(DUMP_JS)
            out.update(data)

            print("\n[nav links]")
            for nl in data["navLinks"]:
                print(f"   text={nl['text']!r} href={nl['href']} aria-controls={nl['ariaControls']} data-target={nl['dataTarget']}")
            print("\n[.tab-pane]")
            for p in data["panes"]:
                print(f"   #{p['index']} id={p['id']!r} cls={p['cls']!r} labelledby={p['ariaLabelledby']!r} len={p['innerTextLen']}")
                print(f"       headSteps={p['headStep']}")
                print(f"       head={p['innerTextHead']!r}")
        finally:
            OUT.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"\n[เสร็จ] เขียน {OUT.name}")
            ctx.close()
            browser.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
