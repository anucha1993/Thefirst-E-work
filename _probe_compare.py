"""Compare server responses for start=0 and start=10"""
import sys, json
sys.path.insert(0, ".")
from playwright.sync_api import sync_playwright
from scrape_wa import load_config, login, goto_account_aliens, _click_alien_subtab

cfg = load_config()
with sync_playwright() as pw:
    b = pw.chromium.launch(headless=False, args=["--ignore-certificate-errors"])
    ctx = b.new_context(locale="th-TH", ignore_https_errors=True)
    page = ctx.new_page()
    login(page, cfg)
    page.wait_for_timeout(1500)
    goto_account_aliens(page)
    page.wait_for_timeout(2000)
    _click_alien_subtab(page, "ลูกจ้างที่มีใบอนุญาตทำงาน")
    page.wait_for_timeout(3000)

    # ดึง 3 หน้าผ่าน api-direct เปรียบเทียบ row identity
    meta = page.evaluate("""() => {
      const dt = $('#tb_aliens').DataTable();
      return {url: dt.ajax.url(), params: dt.ajax.params()};
    }""")
    url = meta["url"]
    if url.startswith("/"):
        from urllib.parse import urlparse
        u = urlparse(page.url); url = f"{u.scheme}://{u.netloc}{url}"
    base = meta["params"]
    def post(start, draw):
        form = {}
        def flat(prefix, val):
            if isinstance(val, dict):
                for k, v in val.items(): flat(f"{prefix}[{k}]" if prefix else k, v)
            elif isinstance(val, list):
                for i, v in enumerate(val): flat(f"{prefix}[{i}]", v)
            else:
                form[prefix] = "" if val is None else str(val)
        flat("", base)
        form["start"] = str(start); form["length"] = "10"; form["draw"] = str(draw)
        r = page.context.request.post(url, form=form, headers={"X-Requested-With":"XMLHttpRequest","Referer":page.url})
        return r.json()

    p1 = post(0, 1)
    p2 = post(10, 2)
    p3 = post(20, 3)
    def keys(j):
        data = j.get("data") or []
        return [(d.get("alien_id"), d.get("alien_id_card"), d.get("alien_emp_rel_id"), d.get("alien_permit_no"), d.get("alien_th_name") or d.get("alien_name")) for d in data]
    print(f"recordsTotal={p1.get('recordsTotal')} recordsFiltered={p1.get('recordsFiltered')}")
    print("=== page1 keys ==="); [print(k) for k in keys(p1)]
    print("=== page2 keys ==="); [print(k) for k in keys(p2)]
    print("=== page3 keys ==="); [print(k) for k in keys(p3)]
    s1 = set(map(str, keys(p1))); s2 = set(map(str, keys(p2))); s3 = set(map(str, keys(p3)))
    print(f"page1 ∩ page2 = {len(s1 & s2)} (of {len(s1)})")
    print(f"page2 ∩ page3 = {len(s2 & s3)} (of {len(s2)})")
    print("=== first row of p1 (raw) ===")
    if p1.get("data"): print(json.dumps(p1["data"][0], ensure_ascii=False)[:600])
    print("=== first row of p2 (raw) ===")
    if p2.get("data"): print(json.dumps(p2["data"][0], ensure_ascii=False)[:600])
    ctx.close(); b.close()
