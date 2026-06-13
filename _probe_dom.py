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
      const allTables = Array.from(document.querySelectorAll('table')).map(t => ({id: t.id, cls: t.className, rows: t.querySelectorAll('tbody tr').length}));
      // หา pagination element ทุกแบบ
      const paginations = Array.from(document.querySelectorAll('.dataTables_paginate, .pagination, [class*=paginate]')).map(p => ({id: p.id, cls: p.className, html: p.outerHTML.slice(0, 500)}));
      const nextBtns = Array.from(document.querySelectorAll('[id$=_next], .next, .paginate_button.next, a:not(:has(*))')).filter(a => /ถัดไป|next/i.test(a.textContent||'')).map(a => ({id: a.id, cls: a.className, txt: (a.textContent||'').trim().slice(0,30), tag: a.tagName, html: a.outerHTML.slice(0,200)}));
      return {tables: allTables, paginations, nextBtns};
    }""")
    print(json.dumps(info, indent=2, ensure_ascii=False))
    ctx.close(); b.close()
