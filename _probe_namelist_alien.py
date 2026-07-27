"""Probe: หา DOM ของตาราง 'รายชื่อคนต่างด้าวที่ทำการยื่นคำขอต่ออายุใบอนุญาตทำงานแล้ว'
URL: https://eworkpermit.doe.go.th/Requtst63_2/NameListAlien?form_type=MT_63_2_3103_RENEWAL
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

from scrape_wa import login  # noqa: E402


def main() -> int:
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "บริษัทนำเข้า (บนจ.)"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
        "headless": False,
    }
    URL = "https://eworkpermit.doe.go.th/Requtst63_2/NameListAlien?form_type=MT_63_2_3103_RENEWAL"

    ajax_hits = []
    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False,
                               args=["--ignore-certificate-errors", "--start-maximized"])
        ctx = b.new_context(locale="th-TH", ignore_https_errors=True,
                            viewport={"width": 1600, "height": 1000})
        page = ctx.new_page()

        def on_req(req):
            u = req.url or ""
            if "NameListAlien" in u or "Requtst63_2" in u:
                try: post = req.post_data or ""
                except Exception: post = ""
                ajax_hits.append({"url": u, "method": req.method, "post": post[:800]})
        page.on("request", on_req)

        print(f"[i] login as {cfg['username']!r} ({cfg['user_type']})")
        login(page, cfg)
        print(f"[i] logged in — url: {page.url}")

        print(f"[i] goto {URL}")
        page.goto(URL, wait_until="domcontentloaded", timeout=45_000)
        page.wait_for_timeout(4000)
        print(f"[i] arrived: {page.url}")

        # ---- Dump DOM structure ----
        info = page.evaluate(r"""() => {
            const out = {tables: [], pageLenSelects: [], paginators: [], headings: []};
            document.querySelectorAll('table').forEach(t => {
                const vis = t.offsetParent !== null;
                if (!vis) return;
                const rect = t.getBoundingClientRect();
                const rows = t.querySelectorAll('tbody tr').length;
                const heads = Array.from(t.querySelectorAll('thead th, thead td'))
                    .map(h => (h.textContent||'').replace(/\s+/g,' ').trim());
                let isDT = false, dtInfo = null;
                if (window.jQuery && window.jQuery.fn.dataTable) {
                    try {
                        isDT = window.jQuery.fn.dataTable.isDataTable(t);
                        if (isDT) {
                            const dt = window.jQuery(t).DataTable();
                            const pi = dt.page.info();
                            dtInfo = {recordsTotal: pi.recordsTotal, recordsDisplay: pi.recordsDisplay,
                                      page: pi.page, pages: pi.pages, length: pi.length,
                                      serverSide: dt.settings()[0].oFeatures.bServerSide,
                                      ajaxUrl: (dt.settings()[0].ajax && (typeof dt.settings()[0].ajax === 'string' ? dt.settings()[0].ajax : dt.settings()[0].ajax.url)) || ''};
                        }
                    } catch(e) { dtInfo = {err: String(e)}; }
                }
                // sample first 2 rows
                const sample = Array.from(t.querySelectorAll('tbody tr')).slice(0, 2).map(tr =>
                    Array.from(tr.querySelectorAll('td')).map(td => (td.innerText||'').replace(/\s+/g,' ').trim().slice(0, 150))
                );
                out.tables.push({
                    id: t.id, cls: t.className.slice(0, 100),
                    w: rect.width, h: rect.height, rows,
                    heads, isDT, dtInfo, sample,
                });
            });
            // page length selectors
            document.querySelectorAll('select[name$="_length"], select.dataTables_length').forEach(s => {
                out.pageLenSelects.push({
                    id: s.id, name: s.name,
                    opts: Array.from(s.options).map(o => o.value),
                    current: s.value,
                });
            });
            // paginator info
            document.querySelectorAll('.dataTables_paginate, .pagination').forEach(p => {
                out.paginators.push({
                    cls: p.className.slice(0, 60),
                    text: (p.textContent||'').replace(/\s+/g,' ').trim().slice(0, 200),
                });
            });
            // headings
            out.headings = Array.from(document.querySelectorAll('h1,h2,h3,h4,h5'))
                .map(h => (h.textContent||'').trim()).filter(Boolean).slice(0, 15);
            return out;
        }""")
        Path("_probe_namelist_alien_out.json").write_text(
            json.dumps({"info": info, "ajax": ajax_hits}, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )

        print("\n=== Tables ===")
        for i, t in enumerate(info.get("tables", [])):
            print(f"  [{i}] id={t['id']!r} rows={t['rows']} isDT={t['isDT']}")
            if t.get("dtInfo"):
                d = t["dtInfo"]
                print(f"      recordsTotal={d.get('recordsTotal')} pages={d.get('pages')} length={d.get('length')} serverSide={d.get('serverSide')} ajax={d.get('ajaxUrl')!r}")
            if t.get("heads"):
                print(f"      heads: {t['heads']}")
            if t.get("sample"):
                for r_i, r in enumerate(t["sample"]):
                    print(f"      row{r_i}: {r}")
        print("\n=== Page Length Selects ===")
        for s in info["pageLenSelects"]:
            print(f"  id={s['id']!r} name={s['name']!r} opts={s['opts']} current={s['current']!r}")
        print("\n=== Ajax hits ===")
        for a in ajax_hits[:6]:
            print(f"  {a['method']} {a['url'][:120]}")
            if a.get("post"):
                print(f"    post: {a['post'][:250]}")

        print("\n[i] จะเปิดค้างไว้ 30 วิ ดูโครงสร้าง")
        page.wait_for_timeout(30_000)
        ctx.close(); b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
