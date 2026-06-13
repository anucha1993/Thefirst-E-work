import sys, json
sys.path.insert(0, ".")
from playwright.sync_api import sync_playwright
from scrape_wa import load_config, login, goto_account_aliens, _click_alien_subtab

cfg = load_config()
with sync_playwright() as pw:
    b = pw.chromium.launch(headless=False, args=["--ignore-certificate-errors"])
    ctx = b.new_context(locale="th-TH", ignore_https_errors=True)
    page = ctx.new_page()
    login(page, cfg); page.wait_for_timeout(1500)
    goto_account_aliens(page); page.wait_for_timeout(2000)
    _click_alien_subtab(page, "ลูกจ้างที่มีใบอนุญาตทำงาน")
    page.wait_for_timeout(4000)
    info = page.evaluate("""() => {
      const out = [];
      document.querySelectorAll('table').forEach(t => {
        const visible = t.offsetParent !== null;
        const rect = t.getBoundingClientRect();
        const isDT = (typeof $ !== 'undefined') && $.fn.dataTable.isDataTable(t);
        let dtInfo = null;
        if (isDT) {
          try {
            const dt = $(t).DataTable();
            const pi = dt.page.info();
            dtInfo = {recordsTotal: pi.recordsTotal, recordsDisplay: pi.recordsDisplay, page: pi.page, pages: pi.pages, length: pi.length, serverSide: dt.settings()[0].oFeatures.bServerSide};
          } catch(e) { dtInfo = {err: String(e)}; }
        }
        out.push({id: t.id, cls: t.className, visible, w: rect.width, h: rect.height, tbody_rows: t.querySelectorAll('tbody tr').length, isDT, dtInfo});
      });
      return out;
    }""")
    print(json.dumps(info, indent=2, ensure_ascii=False))
    ctx.close(); b.close()
