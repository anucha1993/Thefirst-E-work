"""Diagnostic: เปิด detail ของ 1 เลขคำขอ แล้ว dump โครงสร้างจริง
- ชื่อ tab ทั้งหมด
- แท็บ 'เอกสารตอบรับ': element ที่มี onclick + ข้อความ (ดูว่าเอกสารใช้ฟังก์ชันอะไร / มีไหม)
- แท็บ 'ข้อมูลคนต่างด้าว' + 'คำขออนุญาต': snippet ข้อความ (เช็คว่าโครงสร้างต่างไหม)
ใช้: python _probe_result_diag.py [เลขคำขอ]   (ดีฟอลต์ 69125300001797)
พิมพ์เฉพาะ username — ไม่พิมพ์รหัสผ่าน
"""
from __future__ import annotations
import os
import sys
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).parent))
load_dotenv()
from scrape_wa import (  # noqa: E402
    login, goto_tracking, apply_wa_filter, collect_all_wa_rows,
    _open_detail_direct, _is_detail_loaded, _is_logged_out, ensure_session,
)

DEFAULT_REQ = "69125300001797"


def main() -> int:
    req_no = (sys.argv[1].strip() if len(sys.argv) > 1 else DEFAULT_REQ)
    cfg = {
        "username": os.getenv("EWP_USERNAME", ""),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน"),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit"),
    }
    if not cfg["username"] or not cfg["password"]:
        print("ERR: EWP_USERNAME/EWP_PASSWORD ไม่ถูกตั้งใน .env")
        return 1
    login_cfg = {"username": cfg["username"], "password": cfg["password"],
                 "user_type": cfg["user_type"], "method": cfg["method"]}
    print(f"[i] บัญชี: {cfg['username']} | เลขคำขอ: {req_no}")

    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, args=["--start-maximized"])
        ctx = b.new_context(locale="th-TH", viewport={"width": 1500, "height": 900},
                            ignore_https_errors=True, accept_downloads=True)
        page = ctx.new_page()
        try:
            print("[1] login... (แก้ reCAPTCHA)")
            login(page, login_cfg)
            print(f"[+] login OK ({page.url[:70]})")
            goto_tracking(page)
            if _is_logged_out(page):
                ensure_session(page, cfg); goto_tracking(page)
            apply_wa_filter(page, "", status_ids=["WP", "WCOSNA", "WA", "AP", "SS"])
            rows = collect_all_wa_rows(page)
            by_req = {str(r.get("reqNo", "")).strip(): r for r in rows if r.get("reqNo")}
            row = by_req.get(req_no)
            if not row:
                print(f"[!] ไม่พบเลข {req_no} ในบัญชีนี้ ({len(by_req)} คำขอ)")
                return 2
            print(f"[2] row: form_type={row.get('form_type')} status={row.get('status')} "
                  f"statusText={row.get('statusText')} desc={row.get('desc')}")
            if not _open_detail_direct(page, row):
                print("[!] เปิด detail ไม่สำเร็จ")
                return 3
            print(f"[3] detail URL: {page.url}")
            print(f"    _is_detail_loaded = {_is_detail_loaded(page)}")

            # ---- dump tab names ----
            tabs = page.evaluate(
                r"""() => [...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a')]
                    .map(a => ({txt:(a.innerText||'').replace(/\s+/g,' ').trim(),
                                href:a.getAttribute('href')||''}))
                    .filter(t => t.txt)"""
            )
            print("\n[TABS]")
            for t in tabs:
                print(f"    - '{t['txt']}'  href={t['href']}")

            # ---- คลิกแท็บเอกสารตอบรับ (native click) + รอ AJAX + dump ----
            resp_href = page.evaluate(
                r"""() => { const a=[...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a')]
                    .find(a=>/เอกสารตอบรับ/.test(a.innerText||'')); return a?(a.getAttribute('href')||''):''; }"""
            )
            print(f"\n[เอกสารตอบรับ] tab href = {resp_href!r}")
            # native click (trigger event handler จริง) → fallback JS click
            clicked = False
            try:
                if resp_href.startswith("#"):
                    page.locator(f'a[href="{resp_href}"]').first.click(timeout=5000)
                else:
                    page.get_by_text("เอกสารตอบรับ", exact=False).first.click(timeout=5000)
                clicked = True
            except Exception as e:
                print(f"    native click ล้มเหลว: {str(e).splitlines()[0][:80]} → ลอง JS click")
                clicked = page.evaluate(
                    r"""() => { const f=[...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a')]
                        .find(a=>/เอกสารตอบรับ/.test(a.innerText||'')); if(f){f.click();return true;} return false;}"""
                )
            print(f"    clicked = {clicked}")
            pane_sel = resp_href if resp_href.startswith("#") else ".tab-pane.active"
            # poll รอ pane โหลด (ไม่ว่างและไม่ใช่ 'ไม่มีเอกสารการตอบรับ') สูงสุด ~10s
            last_txt = ""
            for i in range(20):
                page.wait_for_timeout(500)
                last_txt = page.evaluate(
                    "(sel)=>{const p=document.querySelector(sel)||document.querySelector('.tab-pane.active');return p?((p.innerText||'').trim()):'';}",
                    pane_sel,
                )
                if last_txt and "ไม่มีเอกสารการตอบรับ" not in last_txt and len(last_txt) > 15:
                    print(f"    pane โหลดเสร็จหลัง ~{(i + 1) * 0.5:.1f}s")
                    break
            else:
                print(f"    pane ไม่เปลี่ยน (ข้อความ: {last_txt[:60]!r})")

            # ทุก element ที่คลิกได้ในแท็บ — ดูว่าไอคอนดู/ดาวน์โหลดใช้อะไร
            clickables = page.evaluate(
                r"""(sel) => {
                    const pane = document.querySelector(sel) || document.querySelector('.tab-pane.active') || document;
                    return [...pane.querySelectorAll('a,button,[onclick],[href]')]
                        .map(el => ({tag:el.tagName.toLowerCase(),
                                     txt:(el.innerText||el.textContent||'').replace(/\s+/g,' ').trim().slice(0,60),
                                     onclick:(el.getAttribute('onclick')||'').slice(0,140),
                                     href:(el.getAttribute('href')||'').slice(0,140),
                                     title:(el.getAttribute('title')||el.getAttribute('data-original-title')||'').slice(0,40),
                                     cls:(el.className||'').toString().slice(0,50)}))
                        .filter(o => o.onclick || (o.href && o.href!=='#') || /btn|icon|fa/.test(o.cls));
                }""",
                pane_sel,
            )
            print(f"    clickable elements: {len(clickables)}")
            for o in clickables:
                print(f"      • <{o['tag']}> '{o['txt']}' title={o['title']!r} cls={o['cls']!r}")
                if o["onclick"]:
                    print(f"          onclick= {o['onclick']}")
                if o["href"] and o["href"] != "#":
                    print(f"          href= {o['href']}")

            pane_html = page.evaluate(
                "(sel)=>{const p=document.querySelector(sel)||document.querySelector('.tab-pane.active');return p?(p.innerHTML||'').slice(0,3500):'(no pane)';}",
                pane_sel,
            )
            print("\n    ---- pane innerHTML (เอกสารตอบรับ, 3500 ตัวแรก) ----")
            print("    " + pane_html.replace("\n", "\n    "))
            print("\n    ---- pane text (เอกสารตอบรับ) ----")
            print("    " + last_txt.replace("\n", "\n    ")[:1500])

            # ---- ค้นหา JS ที่โหลด #DetailDocumentList (function + endpoint) ----
            print("\n[JS] window functions ที่เกี่ยวกับเอกสาร:")
            fns = page.evaluate(
                r"""() => Object.getOwnPropertyNames(window)
                    .filter(k => { try { return typeof window[k]==='function'
                        && /doc|confirm|acknowledg|เอกสาร/i.test(k); } catch(e){ return false; } })"""
            )
            print(f"    {fns}")
            print("\n[JS] inline <script> ที่อ้างถึง DetailDocumentList/divHasDoc/GetDocument:")
            scripts = page.evaluate(
                r"""() => {
                    const hits=[];
                    document.querySelectorAll('script').forEach(s=>{
                        const t=s.textContent||'';
                        if(/DetailDocumentList|divHasDoc|divNoneDoc|GetDocumentConfirm|GetDocument|Acknowledg/i.test(t)){
                            hits.push(t);
                        }
                    });
                    return hits.join('\n/*=====SCRIPT BREAK=====*/\n');
                }"""
            )
            print("    " + (scripts or "(ไม่พบ inline script — อาจอยู่ในไฟล์ .js ภายนอก)")[:6000].replace("\n", "\n    "))

            # ลองหา endpoint ในไฟล์ js ภายนอก
            ext_srcs = page.evaluate(
                r"""() => [...document.querySelectorAll('script[src]')].map(s=>s.src)
                    .filter(u=>/detail|request|document|tracking|form41/i.test(u))"""
            )
            print(f"\n[JS] external script src ที่น่าสน: {ext_srcs}")

            # ---- ทดลอง trigger event ที่อาจโหลดเอกสาร แล้วเช็ค #DetailDocumentList ซ้ำ ----
            page.evaluate(
                r"""() => {
                    const a=[...document.querySelectorAll('a')].find(x=>/เอกสารตอบรับ/.test(x.innerText||''));
                    if(!a) return;
                    try{ if(window.jQuery){ jQuery(a).trigger('click'); jQuery(a).trigger('shown.bs.tab'); } }catch(e){}
                    a.dispatchEvent(new Event('click',{bubbles:true}));
                    a.dispatchEvent(new MouseEvent('click',{bubbles:true}));
                }"""
            )
            page.wait_for_timeout(3500)
            after = page.evaluate(
                r"""() => { const d=document.querySelector('#DetailDocumentList');
                    const h=document.querySelector('#divHasDoc'); const n=document.querySelector('#divNoneDoc');
                    return {listLen:(d?(d.innerHTML||'').length:-1),
                            hasDocDisplay:(h?getComputedStyle(h).display:'?'),
                            noneDisplay:(n?getComputedStyle(n).display:'?')}; }"""
            )
            print(f"[trigger] หลัง trigger events: DetailDocumentList.len={after['listLen']} "
                  f"divHasDoc.display={after['hasDocDisplay']!r} divNoneDoc.display={after['noneDisplay']!r}")

            # ---- (A) อ่าน hidden inputs + บัญชีที่ login + ลองยิง endpoint ด้วยหลายค่า ----
            print("\n[row] ค่าจาก collect_all_wa_rows:")
            for k in ("reqNo", "user_id", "group_id", "status", "form_type", "institution_id", "id"):
                print(f"      {k} = {row.get(k)!r}")
            import json as _json
            page_ids = page.evaluate(
                r"""() => {
                    const gv = id => { const e=document.getElementById(id); return e? (e.value ?? e.textContent ?? '') : null; };
                    const hiddens = {};
                    document.querySelectorAll('input').forEach(e=>{
                        const key = e.id || e.getAttribute('name'); if(!key) return;
                        if(/group|form|doc|request|^id$|user/i.test(key)) hiddens[key] = e.value;
                    });
                    let imp=''; try{ imp=(JSON.stringify(window.important_data)||'').slice(0,600);}catch(e){}
                    return {group_id: gv('group_id'), form_id: gv('form_id'), id: gv('id'), hiddens, important_data: imp};
                }"""
            )
            print("\n[ids] hidden inputs + important_data:")
            print("    " + _json.dumps(page_ids, ensure_ascii=False, indent=2)[:2200].replace("\n", "\n    "))

            candidates = {
                "reqNo": req_no,
                "#group_id.val": str(page_ids.get("group_id") or ""),
                "#form_id.val": str(page_ids.get("form_id") or ""),
                "#id.val": str(page_ids.get("id") or ""),
            }
            for hk, hv in (page_ids.get("hiddens") or {}).items():
                if hv and str(hv) not in candidates.values():
                    candidates[f"#{hk}"] = str(hv)
            print("\n[A] ลองยิง POST /ConfirmationDocument/GetDetailSectionMouList ด้วยค่าต่าง ๆ:")
            for name, val in candidates.items():
                if not val:
                    print(f"    - {name}: (ว่าง)")
                    continue
                res = page.evaluate(
                    r"""async (gid) => { try { const r = await $.ajax({type:"POST", url:"/ConfirmationDocument/GetDetailSectionMouList",
                        data:{group_Id: gid}, dataType:"json"}); return {n:(r&&r.data?r.data.length:-1)}; }
                        catch(e){ return {err:String(e&&e.status)}; } }""",
                    val,
                )
                res2 = page.evaluate(
                    r"""async (gid) => { try { const r = await $.ajax({type:"POST", url:"/RequestForm41/GetDetailDocumentList",
                        data:{group_Id: gid}, dataType:"json"}); return {n:(r&&r.data?(Array.isArray(r.data)?r.data.length:'obj'):-1)}; }
                        catch(e){ return {err:String(e&&e.status)}; } }""",
                    val,
                )
                print(f"    - {name} = {val!r} → MouList:{res}  Form41List:{res2}")

            # ---- แท็บข้อมูลคนต่างด้าว snippet ----
            page.evaluate(
                r"""() => { const f=[...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a')]
                    .find(a=>/ข้อมูลคนต่างด้าว/.test(a.innerText||'')); if(f)f.click(); }"""
            )
            page.wait_for_timeout(1200)
            alien_text = page.evaluate(
                r"""() => { const p=document.querySelector('.tab-pane.active')||document.body;
                    return (p.innerText||'').replace(/\n{2,}/g,'\n').trim().slice(0,1000); }"""
            )
            print("\n    ---- ข้อมูลคนต่างด้าว snippet ----")
            print("    " + alien_text.replace("\n", "\n    "))
            return 0
        finally:
            page.wait_for_timeout(1200)
            ctx.close(); b.close()


if __name__ == "__main__":
    raise SystemExit(main())
