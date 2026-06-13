"""Probe: หลัง login แล้ว สังเกต XHR ที่ DataTable ยิงไปเมื่อกด page 2 และ 3"""
import sys, json, time
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
        url = resp.url
        if "GetAlienList" in url or "AlienList" in url:
            try:
                body = resp.text()[:500]
            except Exception:
                body = "<no body>"
            captured.append({"status": resp.status, "url": url, "method": resp.request.method, "post": resp.request.post_data, "body_snippet": body})
            print(f"[XHR] {resp.status} {resp.request.method} {url}\n  POST: {resp.request.post_data[:300] if resp.request.post_data else None}\n  BODY: {body[:300]}")
    page.on("response", on_response)

    login(page, cfg)
    page.wait_for_timeout(2000)
    goto_account_aliens(page)
    page.wait_for_timeout(2000)
    _click_alien_subtab(page, "ลูกจ้างที่มีใบอนุญาตทำงาน")
    page.wait_for_timeout(3000)

    print("=== Now navigating page 2 ===")
    page.evaluate(r"""async () => {
      const dt = $('#tb_aliens').DataTable();
      console.log('PRE page', dt.page(), 'len', dt.page.info().length, 'total', dt.page.info().recordsTotal);
      return new Promise(res => {
        dt.one('xhr.dt', (e, settings, json, xhr) => {
          window._dbg = {status: xhr.status, dataLen: json && json.data ? json.data.length : null, total: json && json.recordsTotal};
        });
        dt.one('draw.dt', res);
        dt.page(1).draw('page');
        setTimeout(res, 20000);
      });
    }""")
    page.wait_for_timeout(2000)
    dbg = page.evaluate("window._dbg")
    print("PAGE2 result:", dbg)
    print("=== Now navigating page 3 ===")
    page.evaluate(r"""async () => {
      const dt = $('#tb_aliens').DataTable();
      return new Promise(res => {
        dt.one('xhr.dt', (e, settings, json, xhr) => {
          window._dbg = {status: xhr.status, dataLen: json && json.data ? json.data.length : null, total: json && json.recordsTotal};
        });
        dt.one('draw.dt', res);
        dt.page(2).draw('page');
        setTimeout(res, 20000);
      });
    }""")
    page.wait_for_timeout(2000)
    print("PAGE3 result:", page.evaluate("window._dbg"))
    print("\n=== Try real UI click 'next' ===")
    try:
        page.locator("#tb_aliens_next a, #tb_aliens_next").first.click(timeout=5000)
        page.wait_for_timeout(3000)
    except Exception as e:
        print("click err", e)
    print("captured count =", len(captured))
    for c in captured[-5:]:
        print(json.dumps(c, ensure_ascii=False)[:400])
    ctx.close(); b.close()
