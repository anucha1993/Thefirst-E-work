import sys, json
sys.path.insert(0, ".")
from playwright.sync_api import sync_playwright
from scrape_wa import load_config, login, goto_account_aliens, _click_alien_subtab

cfg = load_config()
captured = []
with sync_playwright() as pw:
    b = pw.chromium.launch(headless=False, args=["--ignore-certificate-errors"])
    ctx = b.new_context(locale="th-TH", ignore_https_errors=True)
    page = ctx.new_page()

    def on_response(resp):
        if "GetAlienList" in resp.url:
            try:
                j = resp.json()
                first3 = [{"id": d.get("alien_id"), "name_th": (d.get("firstname_th") or "")+" "+(d.get("lastname_th") or "")} for d in (j.get("data") or [])[:3]]
                captured.append({"post": resp.request.post_data, "first3": first3, "total": j.get("recordsTotal")})
            except Exception as e:
                captured.append({"err": str(e)})
    page.on("response", on_response)

    login(page, cfg); page.wait_for_timeout(1500)
    goto_account_aliens(page); page.wait_for_timeout(2000)
    _click_alien_subtab(page, "ลูกจ้างที่มีใบอนุญาตทำงาน")
    page.wait_for_timeout(4000)

    # ค้นหา custom pagination button รอบๆ tb_aliens
    info = page.evaluate("""() => {
      const tbl = document.getElementById('tb_aliens');
      const parent = tbl ? tbl.closest('.card, .dataTables_wrapper, .tab-pane, div') : null;
      // หา button/a class paginate-right ทั้งหน้า
      const allBtns = Array.from(document.querySelectorAll('.paginate-right, .paginate-left, [class*=paginate]')).filter(e => e.offsetParent !== null).map(e => ({tag: e.tagName, id: e.id, cls: e.className, txt: (e.textContent||'').trim().slice(0,30), html: e.outerHTML.slice(0,250)}));
      // หา container ของ tb_aliens
      const wrap = tbl ? tbl.parentElement.parentElement.outerHTML.slice(0,2000) : null;
      return {visiblePaginationBtns: allBtns};
    }""")
    print("=== visible paginate buttons ===")
    print(json.dumps(info, indent=2, ensure_ascii=False))

    print("\n=== คลิก paginate-right ครั้งที่ 1 ===")
    captured.clear()
    page.locator(".paginate-right").first.click(timeout=5000)
    page.wait_for_timeout(2500)
    for c in captured: print(json.dumps(c, ensure_ascii=False)[:400])

    print("\n=== คลิก paginate-right ครั้งที่ 2 ===")
    captured.clear()
    page.locator(".paginate-right").first.click(timeout=5000)
    page.wait_for_timeout(2500)
    for c in captured: print(json.dumps(c, ensure_ascii=False)[:400])

    print("\n=== คลิก paginate-right ครั้งที่ 3 ===")
    captured.clear()
    page.locator(".paginate-right").first.click(timeout=5000)
    page.wait_for_timeout(2500)
    for c in captured: print(json.dumps(c, ensure_ascii=False)[:400])

    ctx.close(); b.close()
