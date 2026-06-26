"""ดึง source ของ cropImage / selectFileupload / handleFileUpload (3.5 รูปถ่าย crop).
login → เปิดฟอร์ม บต.30 → toString ฟังก์ชัน global. ไม่ต้องเดิน step2.
รัน: .venv\\Scripts\\python.exe _probe_bt30_crop.py   (ปิดเองอัตโนมัติ)
"""
from __future__ import annotations

import json
from playwright.sync_api import sync_playwright

from scrape_wa import login, _read_login_accounts, _open_bt30_form, ROOT

OUT = ROOT / "_probe_bt30_crop_out.json"
log = print


def main():
    accounts = _read_login_accounts(ROOT / "UsernameLogin.xlsx")
    acct = next(iter(accounts.values()))
    cfg = {"username": acct["username"], "password": acct["password"],
           "user_type": acct["type"], "method": acct.get("method") or "E-Workpermit"}

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False, args=["--ignore-certificate-errors", "--start-maximized"])
        ctx = browser.new_context(locale="th-TH", ignore_https_errors=True, viewport={"width": 1500, "height": 900})
        page = ctx.new_page()
        out = {}
        try:
            login(page, cfg); page.wait_for_timeout(1200)
            _open_bt30_form(page, log=log)
            page.wait_for_timeout(1500)
            out = page.evaluate(r"""() => {
              const names = ['initCropper','creatfileCrop','clearCropData','closeModalCrop','resetFileInput','checkTypeNameSizeFile'];
              const fns = {};
              for (const n of names) {
                try { fns[n] = (typeof window[n] === 'function') ? window[n].toString() : ('typeof=' + typeof window[n]); }
                catch (e) { fns[n] = 'ERR ' + e; }
              }
              // หา global ที่ชื่อมี crop
              const cropGlobals = [];
              for (const k in window) {
                try { if (/crop/i.test(k)) cropGlobals.push(k + ':' + typeof window[k]); } catch (e) {}
              }
              return { url: location.href, fns, cropGlobals };
            }""")
            OUT.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
            log("\n[done] →", OUT)
            log("URL:", out.get("url"))
            log("cropGlobals:", out.get("cropGlobals"))
            for n, src in out.get("fns", {}).items():
                head = src if src.startswith("typeof") else (src[:600] + " ...")
                log(f"\n  ===== {n} =====\n{head}")
        except Exception as e:
            log("[ERR]", e)
        finally:
            ctx.close(); browser.close()


if __name__ == "__main__":
    main()
