"""เปรียบเทียบ: JS click vs Playwright real click ที่ปุ่ม 'ถัดไป'"""
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
                names = []
                for d in (j.get("data") or [])[:3]:
                    nm = (d.get("firstname_th") or "") + " " + (d.get("lastname_th") or "")
                    names.append({"alien_id": d.get("alien_id"), "name": nm.strip()})
                captured.append({"post": resp.request.post_data, "first_names": names})
            except Exception as e:
                captured.append({"err": str(e)})
    page.on("response", on_response)

    login(page, cfg); page.wait_for_timeout(1500)
    goto_account_aliens(page); page.wait_for_timeout(2000)
    _click_alien_subtab(page, "ลูกจ้างที่มีใบอนุญาตทำงาน")
    page.wait_for_timeout(3000)

    print("\n=== TEST 1: JS click a.click() ===")
    captured.clear()
    page.evaluate("""() => {
      const li = document.querySelector('#tb_aliens_next');
      const a = li.querySelector('a') || li;
      a.click();
    }""")
    page.wait_for_timeout(2500)
    for c in captured: print(json.dumps(c, ensure_ascii=False)[:300])

    print("\n=== TEST 2: Playwright real click ===")
    captured.clear()
    try:
        page.locator("#tb_aliens_next").click(timeout=5000)
    except Exception as e:
        print("click err:", e)
    page.wait_for_timeout(2500)
    for c in captured: print(json.dumps(c, ensure_ascii=False)[:300])

    print("\n=== TEST 3: Playwright click on a tag ===")
    captured.clear()
    try:
        page.locator("#tb_aliens_next a").click(timeout=5000)
    except Exception as e:
        print("click err:", e)
    page.wait_for_timeout(2500)
    for c in captured: print(json.dumps(c, ensure_ascii=False)[:300])

    print("\n=== TEST 4: jQuery trigger click on a ===")
    captured.clear()
    page.evaluate("""() => { $('#tb_aliens_next a').trigger('click'); }""")
    page.wait_for_timeout(2500)
    for c in captured: print(json.dumps(c, ensure_ascii=False)[:300])

    ctx.close(); b.close()
