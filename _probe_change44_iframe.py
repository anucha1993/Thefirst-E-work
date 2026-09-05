"""วินิจฉัย iframe ของหน้า detail CHANGE_44 — เนื้อหาแท็บอยู่ใน iframe หรือไม่
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
OUT = s.ROOT / "_probe_change44_iframe_out.json"

# ตรวจ pane ภายใน frame ใด ๆ
FRAME_JS = r"""() => {
  const norm = s => (s || '').replace(/\s+/g, ' ').trim();
  const panes = ['tab-status', 'tab-request', 'tab-alien'];
  const res = {
    url: location.href,
    idsWithTab: Array.from(document.querySelectorAll('[id]'))
      .map(e => e.id).filter(id => /tab/i.test(id)).slice(0, 40),
    headStep: document.querySelectorAll('.head-step').length,
    labelFormInfo: document.querySelectorAll('.label-form-info').length,
    formInfo: document.querySelectorAll('.form-info').length,
    tabPanes: document.querySelectorAll('.tab-pane').length,
    panes: {},
  };
  for (const pid of panes) {
    const p = document.getElementById(pid);
    res.panes[pid] = p ? {
      exists: true,
      innerTextLen: (p.innerText || '').length,
      innerTextHead: norm(p.innerText).slice(0, 400),
      headStep: p.querySelectorAll('.head-step').length,
      labelFormInfo: p.querySelectorAll('.label-form-info').length,
    } : { exists: false };
  }
  return res;
}"""


def main() -> int:
    cfg = s.load_config(require_login=True)
    cfg["headless"] = False
    out: dict = {"url": URL, "frames": []}

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
            print("    landed:", page.url)

            # iframe element attrs จาก parent
            out["iframeAttrs"] = page.evaluate(r"""() =>
              Array.from(document.querySelectorAll('iframe')).map(f => ({
                id: f.id, name: f.name, src: f.src,
                cls: f.className, w: f.offsetWidth, h: f.offsetHeight,
              }))
            """)
            print("[3] iframe attrs:", json.dumps(out["iframeAttrs"], ensure_ascii=False))

            # วนทุก frame (รวม main) แล้ว eval FRAME_JS
            print("[4] scan ทุก frame:")
            for i, fr in enumerate(page.frames):
                info = {"index": i, "isMain": fr is page.main_frame, "url": fr.url}
                try:
                    # คลิกแท็บ request ภายใน frame นี้ก่อน (เผื่อ lazy)
                    fr.evaluate("() => { const a = document.querySelector('a[href=\"#tab-request\"]'); if (a) a.click(); }")
                    fr.wait_for_timeout(600)
                    fr.evaluate("() => { const a = document.querySelector('a[href=\"#tab-status\"]'); if (a) a.click(); }")
                    fr.wait_for_timeout(600)
                    info["diag"] = fr.evaluate(FRAME_JS)
                except Exception as e:
                    info["error"] = str(e)[:200]
                out["frames"].append(info)
                d = info.get("diag", {})
                print(f"    frame#{i} main={info['isMain']} url={fr.url[:70]}")
                print(f"       headStep={d.get('headStep')} tabPanes={d.get('tabPanes')} "
                      f"idsWithTab={d.get('idsWithTab')}")
                for pid, st in (d.get("panes") or {}).items():
                    print(f"       {pid}: {json.dumps(st, ensure_ascii=False)[:200]}")
        finally:
            OUT.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"[เสร็จ] เขียน {OUT.name}")
            ctx.close()
            browser.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
