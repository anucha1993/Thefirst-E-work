"""
สคริปต์ดึงข้อมูลคำขอที่มีสถานะ "รอยื่นเอกสารเพิ่มเติม" (WA) จาก e-WorkPermit
แล้วบันทึกเป็นไฟล์ Excel สรุปว่าแต่ละคำขอ "ติดปัญหาอะไร"

วิธีใช้:
    python scrape_wa.py                     # ทำทั้งหมด
    python scrape_wa.py --limit 5           # ทำแค่ 5 รายการแรก (สำหรับทดสอบ)
    python scrape_wa.py --out report.xlsx   # ระบุไฟล์ output
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from datetime import date, datetime
from pathlib import Path
from typing import Any

from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from playwright.sync_api import Page, TimeoutError as PWTimeoutError, sync_playwright

# บังคับ stdout/stderr เป็น UTF-8 เพื่อกันปัญหา UnicodeEncodeError บน Windows
# (เช่น console code page เป็น cp874/Thai หรือเมื่อ redirect output ลงไฟล์)
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[union-attr]
    except Exception:
        pass

LOGIN_URL = "https://eworkpermit.doe.go.th/Login"
TRACKING_URL = "https://eworkpermit.doe.go.th/Permit/Tracking"
DETAIL_BASE = "https://eworkpermit.doe.go.th"

# Map form_type → endpoint path (ถอดจาก openDetail() ฝั่งเว็บ)
FORM_TYPE_ENDPOINTS: dict[str, str] = {
    "WP_32": "/Permit/DetailTracking",
    "MT_13_EXIT": "/Forms/DetailFormRequestExit13",
    "MT_13_1_INFORM": "/Permit/DetailTracking13Begin",
    "MT_41_4_59": "/RequestForm41/DetailRequest41",
    "MT_43": "/Forms/DetailFormRequest43",
    "MT_46_59": "/RequestForm46_59/DetailFormRequest46_59",
    "MT_50_1": "/Forms/DetailFormRequest50",
    "MT_59": "/RequestForm59/DetailTracking59",
    "MT_60_1": "/Permit/DetailTracking60_1",
    "MT_60_2": "/Permit/DetailTracking60_2",
    "MT_61": "/Permit/DetailTracking61",
    "MT_62_BOI": "/RequestFormBOI62/DetailFormRequestBOI62",
    "MT_62": "/RequestForm62/DetailFormRequest62",
    "MT_63": "/RequestForm63/DetailFormRequest63",
    "MT_63_1": "/Permit/DetailTracking63_1",
    "MT_64": "/Permit/DetailTracking64",
    "MT_63_2": "/Permit/DetailTracking63_2",
    "NAMELIST_SURVEY": "/Permit/DetailTracking63_2_Namelist",
    "NAMELIST_SURVEY_19": "/Permit/DetailTracking63_2_Namelist",
    "MT_63_2_AGN_CHANGE": "/Permit/DetailTracking63_2_Change",
    "CHANGE_22": "/Forms/DetailFormRequest44",
    "CHANGE_44": "/Forms/DetailFormRequest44",
    "CHANGE_44_22": "/Forms/DetailFormRequest44",
    "CHANGE_45": "/Forms/DetailFormRequest44",
    "CHANGE_44_45": "/Forms/DetailFormRequest44",
    "MT_59_RENEWAL": "/Renew59/DetailFormRenew59",
    "MT_59_MOU_RENEWAL": "/RenewMOU/DetailFormRenewMOU",
    "MT_62_RENEWAL": "/Renew62/DetailRenew62",
    "MT_62_BOI_RENEWAL": "/RenewBOI/DetailRenewBOI",
    "MT_63_RENEWAL": "/Renew63/DetailFormRenew63",
    "MT_63_1_RENEWAL": "/Renew63_1/DetailFormRenew63_1",
    "MT_63_2_RENEWAL": "/Renew/DetailFormRenew63_2",
    "MT_61_EXTEND": "/Renew/DetailFormRenew61",
    "MT_63_2_AGN_RENEWAL": "/Renew63_2/DetailRequestRenew63_2",
    "REPLACE_CARD_25": "/Forms/DetailFormRequest25",
    "REPLACE_CARD_MOU": "/Forms/DetailFormRequest25",
    "MT_63_2_19": "/Request63_2_19/DetailRequest63_2_19",
    "MT_63_V_2": "/RequestForm63/DetailFormRequest63_V_2",
    "MT_63_2_3103_RENEWAL": "/Requtst63_2/DetailRequest63_2_13",
    "MT_63_2_1302_RENEWAL": "/Requtst63_2/DetailRequest63_2_13",
}

# form_type ที่ openDetail() สร้าง URL แบบ path (group_id ใน path, id เป็น query เดียว)
# แทนที่จะเป็น query string ปกติ — ถ้าใช้แบบ query กับพวกนี้จะได้ 404
PATH_STYLE_FORM_TYPES: frozenset[str] = frozenset({
    "MT_13_EXIT", "MT_50_1", "MT_43",
    "CHANGE_22", "CHANGE_44", "CHANGE_44_22", "CHANGE_45", "CHANGE_44_45",
    "REPLACE_CARD_25", "REPLACE_CARD_MOU",
})

# Map form_type → edit endpoint (จาก openPageEdit()) — ใช้ดึงบันทึกเพิ่มเติม
EDIT_ENDPOINTS: dict[str, str] = {
    "MT_59": "/RequestForm59/EditFormRequest59",
    "MT_63": "/RequestForm63/EditFormRequest63",
    "MT_61": "/Permit/EditFormRequest61",
    "MT_62": "/RequestForm62/EditFormRequest62",
    "MT_63_2": "/Permit/EditFormRequest63_2",
    "MT_62_BOI": "/RequestFormBOI62/EditFormRequestBOI62",
    "MT_59_RENEWAL": "/Renew59/EditFormRenew59",
    "MT_63_RENEWAL": "/Renew63/EditFormRenew63",
    "MT_62_RENEWAL": "/Renew62/EditlRenew62",
    "MT_62_BOI_RENEWAL": "/RenewBOI/EditlRenewBOI",
    "MT_63_1_RENEWAL": "/Renew63_1/EditFormRenew63_1",
    "MT_63_2_RENEWAL": "/Renew/EditFormRenew63_2",
    "MT_63_1": "/Permit/EditFormRequest63_1",
    "MT_64": "/Permit/EditFormRequest64",
    "MT_61_EXTEND": "/Renew/EditFormRenew61",
    "MT_59_MOU_RENEWAL": "/RenewMOU/EditFormRenewMOU",
    "MT_63_2_AGN_RENEWAL": "/Renew63_2/UpdateRequestRenew63_2",
    "MT_63_2_19": "/Request63_2_19/UpdateRequest63_2_19",
    "MT_63_V_2": "/RequestForm63/EditFormRequest63_V_2",
    "MT_63_2_3103_RENEWAL": "/Requtst63_2/UpdateRequest63_2_13_31",
    "MT_63_2_1302_RENEWAL": "/Requtst63_2/UpdateRequest63_2_13_31",
}


def build_detail_url(row: dict) -> str:
    """สร้าง URL หน้า detail ตาม form_type — ถอดตรรกะจาก openDetail() ฝั่งเว็บเป๊ะ ๆ

    มี 2 รูปแบบ (ตาม openDetail):
      • path-style — group_id อยู่ใน path, id เป็น query เดียว: {endpoint}/{group_id}?id={id}
        ใช้กับ: MT_13_EXIT, MT_50_1, MT_43, CHANGE_22/44/44_22/45/44_45, REPLACE_CARD_25/MOU
      • query-style (ที่เหลือ): {endpoint}?user_id=..&status=..&group_id=..&type=..&form_type=..&id=..
    """
    ft = row.get("form_type", "")
    endpoint = FORM_TYPE_ENDPOINTS.get(ft)
    if not endpoint:
        # fallback — ลองใช้ endpoint ของ MT_41_4_59 ไปก่อน
        endpoint = "/RequestForm41/DetailRequest41"

    group_id = row.get("group_id")
    rid = row.get("id")

    # path-style: group_id ต่อท้าย path, ใส่ ?id=... เฉพาะเมื่อมี id
    if ft in PATH_STYLE_FORM_TYPES and group_id not in (None, "", "null", "undefined"):
        from urllib.parse import quote
        url = f"{DETAIL_BASE}{endpoint}/{quote(str(group_id))}"
        if rid not in (None, "", "null", "undefined"):
            url += f"?id={quote(str(rid))}"
        return url

    # query-style
    params = {
        "user_id": row.get("user_id"),
        "status": row.get("status"),
        "group_id": group_id,
        "type": row.get("institution_id"),
        "form_type": ft,
        "id": rid,
    }
    from urllib.parse import urlencode
    clean = {k: v for k, v in params.items() if v not in (None, "", "null", "undefined")}
    return f"{DETAIL_BASE}{endpoint}?{urlencode(clean)}"


def build_edit_url(row: dict) -> str | None:
    """สร้าง URL หน้า edit ตาม form_type — ถ้าไม่มี endpoint คืน None"""
    ft = row.get("form_type", "")
    endpoint = EDIT_ENDPOINTS.get(ft)
    if not endpoint:
        return None
    params = {
        "user_id": row.get("user_id"),
        "status": row.get("status"),
        "group_id": row.get("group_id"),
        "type": row.get("institution_id"),
        "form_type": ft,
        "id": row.get("id"),
    }
    from urllib.parse import urlencode
    clean = {k: v for k, v in params.items() if v not in (None, "", "null", "undefined")}
    return f"{DETAIL_BASE}{endpoint}?{urlencode(clean)}"

ROOT = Path(__file__).parent
SCREENSHOTS = ROOT / "screenshots"
SCREENSHOTS.mkdir(exist_ok=True)
# โฟลเดอร์เก็บไฟล์รายงาน (Report) ทุกโหมด ให้เป็นระเบียบ
REPORTS_DIR = ROOT / "reports"

# ─────────────────────────────────────────────────────────────────
# Profile per request_type — กำหนดพฤติกรรมของแต่ละประเภทคำขอ
# ─────────────────────────────────────────────────────────────────
# Status checkbox ids (จาก probe): WP, WCOSNA, WA, AP, SS
#   WP=รอชำระเงิน, WCOSNA=รอตรวจสอบเอกสารของบริษัท OS,
#   WA=รอยื่นเอกสารเพิ่มเติม, AP=รอนัดหมาย, SS=ดำเนินการเสร็จสิ้น
REQUEST_PROFILES: dict[str, dict] = {
    "_default": {
        "filter_status_ids": ["WA"],
        "capture_extra_notes": True,  # มีปุ่ม "ดำเนินการแก้ไข" → เก็บ บันทึกเพิ่มเติม
    },
    # การแจ้งคนต่างด้าวออกจากงานของนายจ้าง
    "MT_13_EXIT": {
        "filter_status_ids": ["WCOSNA", "WA"],
        "capture_extra_notes": False,  # ไม่มีปุ่มแก้ไขในขั้นนี้
    },
}

# Global whitelist — scrape detail เฉพาะแถวที่ "สถานะ" มีคำเหล่านี้เป็น substring
# ถ้าเป็น None / [] = ไม่กรอง (scrape ทุกแถว)
# ตั้งผ่าน .env: EWP_STATUS_WHITELIST="status1|status2|..."  (คั่นด้วย |)
DEFAULT_STATUS_WHITELIST: list[str] = [
    "รออนุมัติคำขอของนายทะเบียน (สำนักงานจัดหางาน)",
    "รอยื่นเอกสารเพิ่มเติม",
    "รอพิจารณาคำขอของผู้ช่วยนายทะเบียน (สำนักงานจัดหางาน)",
    "รอชำระเงิน",
    "รอนัดหมาย",
    "ดำเนินการเสร็จสิ้น",
]


def get_profile(request_type: str) -> dict:
    """Merge default + override profile สำหรับ request_type ที่กำหนด"""
    base = dict(REQUEST_PROFILES["_default"])
    base.update(REQUEST_PROFILES.get(request_type or "", {}))
    return base


def row_matches_whitelist(row: dict, whitelist: list[str] | None) -> bool:
    """ตรวจว่า status ของ row มีคำใน whitelist เป็น substring หรือไม่"""
    if not whitelist:
        return True
    txt = (row.get("statusText") or "")
    return any(w and w in txt for w in whitelist)


# ลำดับฟิลด์ที่ต้องการดึงจากแท็บ "ข้อมูลคนต่างด้าว" (ใช้ตรงๆ จาก label ของหน้าเว็บ)
ALIEN_FIELDS = [
    "ประเภทผู้ใช้งาน",
    "เลขทะเบียนนิติบุคคล",
    "ชื่อสถานประกอบการ(ไทย)",
    "ชื่อสถานประกอบการ(Eng)",
    "ประเภทนิติบุคคล",
    "วันที่จดทะเบียน",
    "สถานะนิติบุคคล",
    "ทุนจดทะเบียน",
    "รหัสธุรกิจ 5 หลัก(tsic)",
    "วัตถุประสงค์",
    "ที่อยู่สถานที่ทำงาน/สาขา",
    "ชื่อผู้จัดการผู้มีอำนาจกระทำการเเทนนิติบุคคล",
    "ใบอนุญาตเลขที่",
    "ออกให้วันที่",
    "ใช้ได้ถึงวันที่",
    "อีเมล",
    "โทรศัพท์",
    "โทรสาร",
    "เลขบัตรประจำตัวประชาชนผู้กระทำการแทน",
    "ชื่อ(ไทย)",
    "ชื่อ(Eng)",
    "สัญชาติคนต่างด้าว",
    "ชื่อบริษัท",
    "ที่อยู่",
    "ตามหนังสือเเต่งตั้งลงวันที่",
    "ประเภทนายจ้าง",
    "รหัสนายจ้าง",
    "เลขบัตรประจำตัวประชาชน/เลขที่นิติบุคคล",
    "ที่อยู่นายจ้าง",
]

# หัวข้อ section (ไม่มี value ต่อท้าย) — ใช้เป็น "ตัวหยุด" เวลา parse
SECTION_HEADERS = {
    "ข้อมูลผู้รับอนุญาตให้นำคนต่างด้าวมาทำงาน",
    "ข้อมูลสำหรับติดต่อ",
    "ข้อมูลผู้ดำเนินการแทน",
    "ข้อมูลนำเข้าคนต่างด้าว",
    "ข้อมูลบริษัทนำเข้าคนต่างด้าวในประเทศต้นทาง",
    "ข้อมูลนายจ้าง",
}

KNOWN_LABELS = set(ALIEN_FIELDS)
STOP_TOKENS = KNOWN_LABELS | SECTION_HEADERS


def env_bool(name: str, default: bool = False) -> bool:
    v = os.getenv(name)
    return default if v is None else v.strip().lower() in {"1", "true", "yes", "y", "on"}


def _launch_chromium(pw, cfg: dict, extra_args: list[str] | None = None):
    """เปิด Chromium ตาม cfg — รวม logic การเปิดเบราว์เซอร์ไว้ที่เดียว
    โหมดการแสดงผล (เรียงตามลำดับความสำคัญ):
      1. hide_window=True (และไม่ headless) → เปิดเบราว์เซอร์ 'จริง' (ไม่ headless)
         แต่ย้ายหน้าต่างออกไปนอกจอ → มองไม่เห็น แต่เว็บตรวจไม่พบว่าเป็น headless
         (เหมาะกับเว็บที่ block headless / ต้องการพฤติกรรมเหมือน browser จริง)
      2. headless=True → headless จริง (ไม่มีหน้าต่างเลย เบาสุด)
      3. ปกติ → แสดงหน้าต่างตามปกติ
    """
    args = list(extra_args or [])
    hide = bool(cfg.get("hide_window", False))
    headless = bool(cfg.get("headless", False))
    if hide and not headless:
        # --start-maximized จะ override ตำแหน่งหน้าต่าง → ตัดออกเมื่อจะซ่อนนอกจอ
        args = [a for a in args if a != "--start-maximized"]
        args += ["--window-position=-32000,-32000", "--window-size=1920,1080"]
        return pw.chromium.launch(headless=False, args=args)
    return pw.chromium.launch(headless=headless, args=args)


def load_config(require_login: bool = True) -> dict:
    load_dotenv(ROOT / ".env")
    cfg = {
        "username": os.getenv("EWP_USERNAME", "").strip(),
        "password": os.getenv("EWP_PASSWORD", ""),
        "user_type": os.getenv("EWP_USER_TYPE", "ผู้กระทำการแทน").strip(),
        "method": os.getenv("EWP_LOGIN_METHOD", "E-Workpermit").strip(),
        "headless": env_bool("EWP_HEADLESS", False),
        "request_type": os.getenv("EWP_REQUEST_TYPE", "MT_59_MOU_RENEWAL").strip(),
    }
    # whitelist สถานะ — override ผ่าน .env (คั่นด้วย |), ค่าว่าง = ใช้ default
    wl = os.getenv("EWP_STATUS_WHITELIST", "").strip()
    if wl:
        cfg["status_whitelist"] = [s.strip() for s in wl.split("|") if s.strip()]
    # โหมดที่ login จาก Excel (receipts/results/bt30) ไม่ต้องใช้ EWP_USERNAME/PASSWORD ใน .env
    if require_login and (not cfg["username"] or not cfg["password"]):
        raise SystemExit("[ERROR] ตั้งค่า EWP_USERNAME / EWP_PASSWORD ใน .env ก่อน")
    return cfg


def login(page: Page, cfg: dict) -> None:
    page.goto(LOGIN_URL, wait_until="domcontentloaded")
    # รอจน dropdown user_type โหลด options เสร็จ (เคยพบ race ในระบบช้า)
    user_type_label = cfg["user_type"]
    try:
        page.wait_for_function(
            r"""(label) => {
              const sels = document.querySelectorAll('select');
              if (sels.length < 2) return false;
              return Array.from(sels[1].options).some(o => (o.textContent||'').trim() === label);
            }""",
            arg=user_type_label,
            timeout=15_000,
        )
    except PWTimeoutError:
        pass
    page.locator("select").nth(1).select_option(label=user_type_label)
    page.wait_for_timeout(500)
    method_value = "2" if cfg["method"].lower().replace("-", "") == "eworkpermit" else "1"
    page.evaluate(
        """(v) => {
            const r = document.querySelector(`input[type=radio][name=radio][value="${v}"]`);
            if (r) { r.checked = true; r.dispatchEvent(new Event('change',{bubbles:true})); r.click(); }
        }""",
        method_value,
    )
    try:
        page.get_by_role("heading", name=cfg["method"], exact=True).first.click(timeout=3000)
    except PWTimeoutError:
        pass
    page.wait_for_timeout(500)
    # ระบบมี input หลายช่อง (#username_login / #employer_login / #agency_login)
    # แสดงเฉพาะตัวที่ตรงกับ user_type → กรอกตัวที่ visible เท่านั้น
    user_box = page.locator(
        "#username_login:visible, #employer_login:visible, #agency_login:visible"
    ).first
    user_box.wait_for(state="visible", timeout=10_000)
    user_box.fill(cfg["username"])
    page.locator("#password_login").fill(cfg["password"])
    # ใช้ JS click แทน .click() — เพื่อไม่ให้ Playwright ค้างรอ navigation จนถึงสถานะ
    # 'load' (บางหน้า e-Service มี resource ที่โหลดไม่จบ → timeout 30s แม้ login สำเร็จ)
    # จากนั้นรอแค่ URL ออกจากหน้า /Login ก็พอ
    page.evaluate(
        "() => { const b = document.querySelector('#validate_login'); if (b) b.click(); }"
    )
    page.wait_for_url(lambda u: "/Login" not in u, timeout=int(cfg.get("login_timeout_ms", 30_000)))
    try:
        page.wait_for_load_state("domcontentloaded", timeout=15_000)
    except PWTimeoutError:
        pass
    print(f"[+] Login OK → {page.url}")


def goto_tracking(page: Page) -> None:
    """เปิดหน้า e-Tracking — ลองเรียก openpageTracking() ก่อน, fallback เป็น goto URL ตรง"""
    cur = page.url or ""
    if "/Permit/Tracking" in cur:
        page.wait_for_timeout(500)
        return
    try:
        ok = page.evaluate(
            "() => { try { if (typeof openpageTracking === 'function') { openpageTracking(); return true; } } catch(e) {} return false; }"
        )
    except Exception:
        ok = False
    if not ok:
        page.goto("https://eworkpermit.doe.go.th/Permit/Tracking",
                  wait_until="domcontentloaded", timeout=30_000)
    page.wait_for_url("**/Permit/Tracking", timeout=20_000)
    page.wait_for_timeout(2500)


def ensure_tracking_ready(page: Page, cfg: dict, request_type: str = "", log=print) -> bool:
    """ยืนยันว่าอยู่ที่หน้า e-Tracking และ filter ถูกใส่แล้ว
    ถ้า session หมด → login ใหม่ + เปิด tracking + ใส่ filter ใหม่
    คืน True ถ้าพร้อมใช้งาน, False ถ้าฟื้นไม่ได้
    """
    try:
        url = page.url or ""
    except Exception:
        url = ""
    on_tracking = "/Permit/Tracking" in url
    if on_tracking and not _is_logged_out(page):
        return True
    try:
        if _is_logged_out(page):
            log("      ⚠ ตรวจพบ session หมดอายุ — กำลัง login ใหม่...")
            login(page, cfg)
        log("      ↻ กลับไปหน้า e-Tracking และตั้งค่า filter ใหม่...")
        goto_tracking(page)
        # ใช้ filter ที่ผู้ใช้เลือกใน cfg ก่อน — fallback ไป profile default ถ้าไม่มี
        status_ids = (
            cfg.get("filter_status_ids")
            or get_profile(request_type).get("filter_status_ids")
            or ["WA"]
        )
        apply_wa_filter(
            page, request_type or "", status_ids=status_ids,
            date_from=(cfg.get("date_from") or ""),
            date_to=(cfg.get("date_to") or ""),
        )
        page.wait_for_timeout(1500)
        return "/Permit/Tracking" in (page.url or "")
    except Exception as e:
        try:
            log(f"      ✗ ฟื้น session ไม่สำเร็จ: {e}")
        except Exception:
            pass
        return False


def _is_logged_out(page: Page) -> bool:
    """ตรวจว่าหน้าปัจจุบันถูกเด้งไปหน้า login หรือยัง
    ตรวจทั้ง URL และ content (กรณี session หมดแบบ silent ที่ URL ไม่เปลี่ยน)
    """
    try:
        url = page.url or ""
    except Exception:
        return False
    # เจอ /Login ใน URL = ถูก redirect ไปหน้า login แน่นอน
    if "/Login" in url:
        return True
    # NOTE: อย่าใช้เงื่อนไข "URL ลงท้าย doe.go.th" — root URL (https://eworkpermit.doe.go.th/)
    # คือหน้า home ของผู้ใช้ที่ login แล้ว ไม่ใช่ logout — เคย false-positive ทุก record
    # ตรวจ content: เซิร์ฟเวอร์บางครั้ง silent-expire โดยไม่ redirect
    # - มี input password → หน้า login แทรกมา
    # - มี modal/ข้อความ "Session Timeout / หมดอายุ / เข้าสู่ระบบใหม่" (popup ค้างหน้า detail)
    try:
        return bool(page.evaluate(
            r"""() => {
                if (document.querySelector('input[type="password"]')) return true;
                const body = (document.body && document.body.innerText) || '';
                if (/Session\s*Timeout/i.test(body)) return true;
                if (/Session.*หมดอายุ|กรุณาเข้าสู่ระบบใหม่|เข้าสู่ระบบใหม่อีกครั้ง/.test(body)) return true;
                return false;
            }"""
        ))
    except Exception:
        return False


def ensure_session(page: Page, cfg: dict, target_url: str | None = None, log=print) -> bool:
    """ถ้า session หมด/ถูก logout → พยายาม login ใหม่แล้ว goto target_url
    คืน True ถ้า session ใช้งานได้ (หรือ login ใหม่สำเร็จ), False ถ้าฟื้นไม่ได้
    """
    if not _is_logged_out(page):
        return True
    try:
        log("      ⚠ ตรวจพบ session หมดอายุ — กำลัง login ใหม่...")
        login(page, cfg)
        if target_url:
            page.goto(target_url, wait_until="domcontentloaded", timeout=25_000)
            page.wait_for_timeout(2500)
        return not _is_logged_out(page)
    except Exception as e:
        try:
            log(f"      ✗ Login ใหม่ไม่สำเร็จ: {e}")
        except Exception:
            pass
        return False


def apply_wa_filter(
    page: Page,
    request_type: str = "",
    status_ids: list[str] | None = None,
    date_from: str = "",
    date_to: str = "",
) -> None:
    """ติ๊ก checkbox สถานะ (อาจหลายตัว) + ตั้งค่ารายการคำขอ + วันที่ยื่นคำขอ แล้ว trigger รีเฟรช

    Args:
        request_type: รหัสรายการคำขอ (เช่น MT_13_EXIT, MT_59_MOU_RENEWAL)
                      "" = ไม่กรอง (ทั้งหมด)
        status_ids: list ของ checkbox id ที่ต้องติ๊ก (เช่น ["WA"], ["WCOSNA","WA"])
                    ถ้า None → ใช้ profile ของ request_type
        date_from: วันที่เริ่ม (รูปแบบ วว/ดด/ปปปป) — "" = ไม่กรอง
        date_to:   วันที่สิ้นสุด (รูปแบบ วว/ดด/ปปปป) — "" = ไม่กรอง
    """
    if status_ids is None:
        status_ids = list(get_profile(request_type).get("filter_status_ids") or ["WA"])

    # ขั้นที่ 1: ติ๊ก status + เซ็ต dropdown + วันที่ แล้วยิง GetData ครั้งแรก
    page.evaluate(
        """({rt, ids, df, dt}) => {
            // เคลียร์ checkbox ทุกตัวก่อน เพื่อไม่ให้ค้างจากครั้งก่อน
            ['WP','WCOSNA','WA','AP','SS'].forEach(id => {
                const cb = document.getElementById(id);
                if (cb) cb.checked = false;
            });
            // ติ๊กเฉพาะที่ต้องการ
            ids.forEach(id => {
                const cb = document.getElementById(id);
                if (cb) cb.checked = true;
            });
            const sel = document.getElementById('Filter_request_list');
            if (sel) {
                sel.value = rt || '';
                if (window.jQuery) window.jQuery(sel).val(rt || '').trigger('change');
            }
            // ---- ตั้งค่าวันที่ยื่นคำขอ (จาก → ถึง) ----
            // ID จริงบนเว็บ: search_start_date, search_end_date
            // ใช้ bootstrap-datepicker → เรียก 'update' ให้ picker sync internal state
            function setDateById(id, val) {
                const inp = document.getElementById(id);
                if (!inp) return;
                const $ = window.jQuery;
                if ($ && $(inp).data('datepicker')) {
                    try {
                        if (val) {
                            $(inp).datepicker('update', val);
                        } else {
                            // เคลียร์ค่า: setDate(null) + clear input value
                            $(inp).datepicker('setDate', null);
                            inp.value = '';
                            $(inp).datepicker('update', '');
                        }
                        // trigger changeDate เพื่อให้ event handler ที่ผูกไว้ทำงาน
                        $(inp).trigger('changeDate');
                    } catch(e) { /* fallback ด้านล่าง */ }
                }
                // fallback: set .value ตรงๆ (เผื่อไม่มี datepicker plugin)
                if (inp.value !== (val || '')) inp.value = val || '';
                inp.dispatchEvent(new Event('input',  {bubbles: true}));
                inp.dispatchEvent(new Event('change', {bubbles: true}));
                inp.dispatchEvent(new Event('blur',   {bubbles: true}));
            }
            setDateById('search_start_date', df || '');
            setDateById('search_end_date',   dt || '');
            if (typeof GetDataRequestFormEtrackingForAlien === 'function') {
                GetDataRequestFormEtrackingForAlien();
            }
        }""",
        {"rt": request_type or "", "ids": status_ids,
         "df": date_from or "", "dt": date_to or ""},
    )
    page.wait_for_timeout(3500)
    # ขั้นที่ 2: ยิงอีกรอบเพื่อกัน race กับ select2/datepicker
    page.evaluate(
        """({rt, ids, df, dt}) => {
            ids.forEach(id => {
                const cb = document.getElementById(id);
                if (cb) cb.checked = true;
            });
            const sel = document.getElementById('Filter_request_list');
            if (sel && window.jQuery) window.jQuery(sel).val(rt || '').trigger('change');
            // reset ค่าวันที่อีกรอบ (เผื่อ datepicker เขียนทับตอน init)
            function _setDate(id, v) {
                const inp = document.getElementById(id);
                if (!inp) return;
                const $ = window.jQuery;
                if ($ && $(inp).data('datepicker')) {
                    try {
                        if (v) $(inp).datepicker('update', v);
                        else { $(inp).datepicker('setDate', null); inp.value = ''; }
                    } catch(e) {}
                }
                if (inp.value !== (v || '')) inp.value = v || '';
                inp.dispatchEvent(new Event('change', {bubbles: true}));
            }
            _setDate('search_start_date', df || '');
            _setDate('search_end_date',   dt || '');
            if (typeof GetDataRequestFormEtrackingForAlien === 'function') {
                GetDataRequestFormEtrackingForAlien();
            }
        }""",
        {"rt": request_type or "", "ids": status_ids,
         "df": date_from or "", "dt": date_to or ""},
    )
    page.wait_for_timeout(3500)

    # ขั้นที่ 3: รอให้ ajax GetDataRequestFormEtrackingForAlien ตอบกลับจริงๆ
    # (สำคัญเมื่อบัญชีมีข้อมูลเยอะ — server response อาจใช้เวลามากกว่า 7s ที่ wait_for_timeout รวม)
    # trigger GetData อีกครั้ง แล้วดัก response — กันกรณีที่ 2 ครั้งแรก request ค้างอยู่/ยังไม่ตอบ
    try:
        with page.expect_response(
            lambda r: "GetDataRequestFormEtrackingForAlien" in (r.url or "")
                      and r.request.method == "POST",
            timeout=60_000,
        ):
            page.evaluate(
                r"""() => {
                  if (typeof GetDataRequestFormEtrackingForAlien === 'function') {
                    GetDataRequestFormEtrackingForAlien();
                  }
                }"""
            )
        # ให้ DataTable วาดผลลัพธ์เสร็จ
        page.wait_for_timeout(1500)
    except Exception:
        # ถ้าไม่มี ajax เกิดขึ้น (เช่น function ไม่มี) → ปล่อยผ่าน
        pass


def apply_request_type_filter(page: Page, request_type: str) -> None:
    """ตั้งค่า dropdown 'รายการคำขอ' (#Filter_request_list) เป็น select2
    เช่น 'MT_59_MOU_RENEWAL' = MoU renewal
    ใส่ '' หรือ '0' = ทั้งหมด (ไม่ filter)
    """
    if not request_type or request_type in {"0", "ALL", "all"}:
        return
    page.evaluate(
        """(val) => {
            const sel = document.getElementById('Filter_request_list');
            if (!sel) return;
            sel.value = val;
            if (window.jQuery) {
                window.jQuery(sel).val(val).trigger('change');
            } else {
                sel.dispatchEvent(new Event('change', {bubbles:true}));
            }
        }""",
        request_type,
    )
    page.wait_for_timeout(800)
    # คลิกปุ่ม "ค้นหา" เพื่อ refilter ตาราง
    clicked = page.evaluate(
        r"""() => {
          const btns = Array.from(document.querySelectorAll('button, a, input[type=button], input[type=submit]'));
          const b = btns.find(x => /ค้นหา/.test((x.innerText||x.value||'').trim()));
          if (b) { b.click(); return true; }
          // fallback: เรียกฟังก์ชันโดยตรง
          if (typeof GetDataRequestFormEtrackingForAlien === 'function') {
              GetDataRequestFormEtrackingForAlien();
              return 'js';
          }
          return false;
        }"""
    )
    page.wait_for_timeout(4000)


def set_all_page_length(page: Page) -> int:
    """ขยาย page length ของ DataTables เป็น CHUNK_SIZE (server cap ~1000).
    คืน: pageLen ปัจจุบัน หรือ -1 ถ้าไม่พบ DataTable
    """
    CHUNK_SIZE = 1000
    return page.evaluate(
        """(N) => {
            const $ = window.jQuery;
            if ($ && $.fn && $.fn.dataTable) {
                let ok = -1;
                $('table.dataTable').each(function() {
                    try {
                        $(this).DataTable().page.len(N).draw();
                        ok = N;
                    } catch(e) {}
                });
                return ok;
            }
            const sel = document.querySelector('select[name$="_length"]');
            if (sel) {
                if (!Array.from(sel.options).some(o => o.value === String(N))) {
                    sel.add(new Option(String(N), String(N)));
                }
                sel.value = String(N);
                sel.dispatchEvent(new Event('change', {bubbles:true}));
                return N;
            }
            return -1;
        }""", CHUNK_SIZE
    )


def _wait_datatable_idle(page: Page, table_id: str = "datatableE_Tracking", max_wait_ms: int = 90_000) -> dict:
    """รอจน DataTables วาดเสร็จ (DOM rows คงที่ และไม่มี processing indicator)"""
    last = -1; stable = 0; waited = 0
    info = {}
    while waited < max_wait_ms:
        page.wait_for_timeout(500); waited += 500
        info = page.evaluate(
            """(tid) => {
                const $ = window.jQuery;
                try {
                    const dt = $('#' + tid).DataTable();
                    const i = dt.page.info();
                    const proc = $('#' + tid + '_processing').is(':visible');
                    return {
                        page: i.page, pages: i.pages,
                        recordsTotal: i.recordsTotal, recordsDisplay: i.recordsDisplay,
                        length: i.length,
                        dom: document.querySelectorAll('#' + tid + ' tbody tr').length,
                        processing: proc,
                    };
                } catch(e) { return {err: String(e)}; }
            }""", table_id
        )
        if info.get("err") or info.get("processing"):
            stable = 0; continue
        cur = info.get("dom") or 0
        if cur == last and cur > 0:
            stable += 1
            if stable >= 3:
                break
        else:
            stable = 0; last = cur
    return info


def collect_all_wa_rows(page: Page, log=print, table_id: str = "datatableE_Tracking") -> list[dict]:
    """เก็บ rows ทุกหน้าโดยตั้ง page.len=1000 แล้ววนทุก page ของ DataTables.
    Server-side DataTables มี cap ~1000 rows/request — ต้อง paginate ดึงทีละหน้า.
    """
    # ตั้ง page length = 1000 และไปหน้าแรก
    page.evaluate(
        """(tid) => {
            const $ = window.jQuery;
            try {
                const dt = $('#' + tid).DataTable();
                dt.page.len(1000).page(0).draw('page');
            } catch(e) {}
        }""", table_id
    )
    info = _wait_datatable_idle(page, table_id)
    pages = int(info.get("pages") or 1)
    recordsDisplay = int(info.get("recordsDisplay") or 0)
    log(f"      [paginate] ทั้งหมด {recordsDisplay} แถว / {pages} หน้า (1000 แถว/หน้า)")

    seen: set = set()
    all_rows: list[dict] = []
    cur_page_idx = 0
    while True:
        # เก็บ rows หน้านี้
        rows_now = page.evaluate(
            r"""(tid) => {
              const out = [];
              document.querySelectorAll('#' + tid + ' tbody tr').forEach(tr => {
                const a = tr.querySelector('a[onclick*="openDetail"]');
                if (!a) return;
                const onclickRaw = a.getAttribute('onclick') || '';
                const hrefRaw = a.getAttribute('href') || '';
                const m = onclickRaw.match(/openDetail\(([^)]+)\)/);
                if (!m) return;
                const args = m[1].split(',').map(s => s.trim().replace(/^"|"$/g,''));
                const [user_id, group_id, status, form_type, institution_id, id] = args;
                const tds = tr.querySelectorAll('td');
                const reqNo = (tds[1] && tds[1].querySelector('span')) ? tds[1].querySelector('span').innerText.trim() : '';
                const requester = (tds[1] && tds[1].querySelectorAll('span')[1]) ? tds[1].querySelectorAll('span')[1].innerText.trim() : '';
                const desc = (tds[2] && tds[2].querySelector('span')) ? tds[2].querySelector('span').innerText.trim().replace(/\s+/g,' ').slice(0,300) : '';
                const dateSubmit = tds[3] ? tds[3].innerText.trim() : '';
                const dateUpdate = tds[4] ? tds[4].innerText.trim() : '';
                const statusText = tds[5] ? tds[5].innerText.trim().replace(/\s+/g,' ') : '';
                out.push({user_id, group_id, status, form_type, institution_id, id, reqNo, requester, desc, dateSubmit, dateUpdate, statusText, onclickRaw, hrefRaw});
              });
              return out;
            }""", table_id
        )
        added = 0
        for r in rows_now:
            key = (r.get("group_id"), r.get("user_id"), r.get("id"))
            if key in seen:
                continue
            seen.add(key)
            all_rows.append(r); added += 1
        log(f"      [paginate] หน้า {cur_page_idx + 1}/{pages} → +{added} (สะสม {len(all_rows)})")

        # ถ้าครบทุกหน้าแล้ว stop
        if cur_page_idx >= pages - 1 or len(all_rows) >= recordsDisplay > 0:
            break
        # ไปหน้าถัดไป
        page.evaluate(
            """(tid) => {
                const $ = window.jQuery;
                try { $('#' + tid).DataTable().page('next').draw('page'); } catch(e) {}
            }""", table_id
        )
        info = _wait_datatable_idle(page, table_id)
        new_page = int(info.get("page") or -1)
        if new_page == cur_page_idx:
            log(f"      [paginate] ไม่ขยับหน้า — หยุด")
            break
        cur_page_idx = new_page
        if cur_page_idx > pages + 5:
            log(f"      [paginate] เกินหน้าสุดท้าย — หยุด")
            break
    return all_rows


def collect_wa_rows(page: Page) -> list[dict]:
    """อ่านทุกแถวที่มี openDetail(...) ในตาราง และเก็บเฉพาะที่ status = WA"""
    rows = page.evaluate(
        r"""() => {
          const out = [];
          document.querySelectorAll('table tbody tr').forEach(tr => {
            const a = tr.querySelector('a[onclick*="openDetail"]');
            if (!a) return;
            const onclickRaw = a.getAttribute('onclick') || '';
            const hrefRaw = a.getAttribute('href') || '';
            const m = onclickRaw.match(/openDetail\(([^)]+)\)/);
            if (!m) return;
            // parse arguments
            const args = m[1].split(',').map(s => s.trim().replace(/^"|"$/g,''));
            const [user_id, group_id, status, form_type, institution_id, id] = args;
            const tds = tr.querySelectorAll('td');
            const reqNo = (tds[1] && tds[1].querySelector('span'))
                ? tds[1].querySelector('span').innerText.trim() : '';
            const requester = (tds[1] && tds[1].querySelectorAll('span')[1])
                ? tds[1].querySelectorAll('span')[1].innerText.trim() : '';
            const desc = (tds[2] && tds[2].querySelector('span'))
                ? tds[2].querySelector('span').innerText.trim().replace(/\s+/g,' ').slice(0,300) : '';
            const dateSubmit = tds[3] ? tds[3].innerText.trim() : '';
            const dateUpdate = tds[4] ? tds[4].innerText.trim() : '';
            const statusText = tds[5] ? tds[5].innerText.trim().replace(/\s+/g,' ') : '';
            out.push({user_id, group_id, status, form_type, institution_id, id, reqNo, requester, desc, dateSubmit, dateUpdate, statusText, onclickRaw, hrefRaw});
          });
          return out;
        }"""
    )
    return rows


def scrape_detail(page: Page, row: dict, cfg: dict | None = None, log=print, capture_extra_notes: bool = True, capture_appointment: bool = False) -> dict:
    """เข้าหน้า detail แล้วดึง หมายเหตุ + ข้อมูลคนต่างด้าว + คำขออนุญาต + บันทึกเพิ่มเติม

    แต่ละ step มี try/except แยกกัน — ถ้าหน้านี้ไม่มี tab/ปุ่ม/section ที่คาดหวัง
    จะ skip เฉพาะส่วนนั้นและคืนค่าที่อ่านได้ทั้งหมด ไม่หยุดทั้งระบบ
    ถ้า session หมดอายุ → พยายาม login ใหม่อัตโนมัติ
    """
    result: dict = {
        "alert_title": "",
        "note": "",
        "note_raw": "",
        "fields": {},
        "structured_fields": {},
        "pane_raw": "",
        "kham_kor_text": "",
        "extra_notes": "",
        "scrape_errors": "",
    }
    errors: list[str] = []

    url = build_detail_url(row)
    nav_ok = False
    # 1) ลองคลิก link ของแถวในตาราง (ถ้ายังอยู่หน้า tracking) — เปิดผ่าน
    #    handler ของเว็บเอง (openDetail) เพื่อให้รองรับทุก form_type/status
    if "/Etracking" in page.url:
        group_id = str(row.get("group_id", "") or "")
        link_loc = None
        if group_id:
            cand = page.locator(f'a[onclick*="openDetail"][onclick*="{group_id}"]').first
            try:
                if cand.count() > 0:
                    link_loc = cand
            except Exception:
                link_loc = None
        if link_loc is not None:
            try:
                with page.expect_navigation(timeout=20_000, wait_until="domcontentloaded"):
                    link_loc.click()
                page.wait_for_timeout(3000)
                nav_ok = True
            except Exception as e:
                errors.append(f"clickRow:{str(e).splitlines()[0][:120]}")

    # 2) ถ้าไม่สำเร็จ — เรียก openDetail() ผ่าน JS แล้วรอ navigation
    if not nav_ok:
        try:
            args = [
                row.get("user_id", ""), row.get("group_id", ""), row.get("status", ""),
                row.get("form_type", ""), row.get("institution_id", ""), row.get("id", ""),
            ]
            try:
                with page.expect_navigation(timeout=15_000, wait_until="domcontentloaded"):
                    page.evaluate(
                        """(a) => {
                            if (typeof openDetail === 'function') {
                                openDetail(a[0], a[1], a[2], a[3], a[4], a[5]);
                            }
                        }""",
                        args,
                    )
                page.wait_for_timeout(3000)
                nav_ok = True
            except Exception:
                pass
        except Exception as e:
            errors.append(f"openDetail:{str(e).splitlines()[0][:120]}")

    # 3) Fallback สุดท้าย — สร้าง URL เอง (อาจล้มเหลวถ้า endpoint mapping ผิด)
    if not nav_ok:
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=25_000)
            page.wait_for_timeout(3500)
            nav_ok = True
        except Exception as e:
            msg = str(e).split("Call log:")[0].strip().splitlines()[0][:200]
            if "ERR_HTTP_RESPONSE_CODE_FAILURE" in msg:
                short = f"เปิดหน้า detail ไม่ได้ (HTTP error) status={row.get('status','?')}"
            elif "Timeout" in msg or "timeout" in msg:
                short = "เปิดหน้า detail timeout"
            else:
                short = f"เปิดหน้า detail ล้มเหลว: {msg}"
            errors.append(f"goto:{short}")
            result["scrape_errors"] = " | ".join(errors)
            # พยายามกลับไปหน้า tracking เพื่อให้ row ถัดไปทำงานต่อได้
            if cfg is not None:
                try:
                    rt = (cfg.get("request_type") or row.get("form_type") or "").strip()
                    ensure_tracking_ready(page, cfg, rt, log=log)
                except Exception:
                    pass
            return result

    # ตรวจ session — ถ้าถูกเด้งไปหน้า login ก็ login ใหม่แล้วเข้าหน้าเดิม
    if cfg is not None and _is_logged_out(page):
        if not ensure_session(page, cfg, target_url=url, log=log):
            errors.append("session:relogin_failed")
            result["scrape_errors"] = " | ".join(errors)
            return result

    # 1) เก็บแท็บ คำขออนุญาต (#tab_default_2)
    try:
        page.evaluate(
            """() => {
                const a = document.querySelector('a[href="#tab_default_2"]');
                if (a) a.click();
            }"""
        )
        page.wait_for_timeout(1200)
        result["kham_kor_text"] = page.evaluate(
            r"""() => {
              const pane = document.querySelector('#tab_default_2');
              if (!pane) return '';
              return (pane.innerText || '').replace(/[ \t]+/g,' ').replace(/\n{2,}/g,'\n').trim().slice(0,8000);
            }"""
        ) or ""
    except Exception as e:
        errors.append(f"tab_default_2:{e}")

    # 2) คลิกแท็บ "ข้อมูลคนต่างด้าว" และดึง note/alert/pane
    try:
        page.evaluate(
            """() => {
                const a = document.querySelector('a[href="#tab_default_2_1"]');
                if (a) a.click();
            }"""
        )
        page.wait_for_timeout(1500)

        payload = page.evaluate(
            r"""() => {
              let noteText = '';
              const info = document.querySelector('.div_info');
              if (info) noteText = info.innerText.replace(/\s+/g,' ').trim();
              let alertTitle = '';
              const at = document.querySelector('.div_info .text-header, .div_info strong, .div_info h4, .div_info h5');
              if (at) alertTitle = at.innerText.trim();
              const pane = document.querySelector('#tab_default_2_1');
              const paneText = pane ? pane.innerText : '';

              // Fallback — สแกนทั้งหน้าหา element ที่มี "หมายเหตุ" (รองรับ MT_13_EXIT
              // ที่ note อยู่ใน info box สีชมพู ไม่ใช่ใน .div_info)
              let fallbackNotes = [];
              if (!/หมายเหตุ\s*[:：]/.test(noteText)) {
                const candidates = Array.from(document.querySelectorAll(
                  '.alert, .alert-info, .bg-info, .info-box, .note, [class*="info"], [class*="alert"], [class*="note"], p, div'
                ));
                const seen = new Set();
                for (const el of candidates) {
                  // ข้าม container ใหญ่เกิน (เอาเฉพาะ element ที่เนื้อหาเป็นข้อความ note โดยตรง)
                  if (el.children.length > 6) continue;
                  const t = (el.innerText || '').replace(/\s+/g,' ').trim();
                  if (!t || t.length > 1200) continue;
                  if (!/หมายเหตุ\s*[:：]/.test(t)) continue;
                  if (seen.has(t)) continue;
                  seen.add(t);
                  fallbackNotes.push(t);
                }
              }
              return {noteText, alertTitle, paneText, fallbackNotes};
            }"""
        ) or {}
        note = payload.get("noteText", "") or ""
        m = re.search(r"(หมายเหตุ\s*[:：].*)", note)
        result["alert_title"] = payload.get("alertTitle", "") or ""
        if m:
            result["note"] = m.group(1).strip()
        else:
            # ใช้ fallback (เลือกอันสั้นสุด = element เฉพาะเจาะจงสุด)
            fb = payload.get("fallbackNotes") or []
            if fb:
                fb.sort(key=len)
                fb_text = fb[0]
                m2 = re.search(r"(หมายเหตุ\s*[:：].*)", fb_text)
                result["note"] = (m2.group(1) if m2 else fb_text).strip()
            else:
                result["note"] = note
        result["note_raw"] = note or (payload.get("fallbackNotes") or [""])[0]
        result["pane_raw"] = payload.get("paneText", "") or ""
        try:
            result["fields"] = parse_alien_pane(result["pane_raw"]) or {}
        except Exception as e:
            errors.append(f"parse_alien:{e}")
    except Exception as e:
        errors.append(f"tab_default_2_1:{e}")

    # 2.5) ดึง label/value ทุกคู่ — ใช้ได้กับทุก form_type ที่มี .label-form-info
    try:
        result["structured_fields"] = extract_label_value_pairs(page) or {}
    except Exception as e:
        errors.append(f"structured:{e}")

    # 2.6) การนัดหมาย (เฉพาะ Template นัดหมาย) — เข้าแท็บ #tab_default_5 อ่าน iframe
    #      ดึง วันที่/เวลา/สถานที่ + ชื่อแรงงาน + ชื่อบริษัท (ทำก่อน extra_notes ที่อาจ navigate ออก)
    if capture_appointment:
        try:
            appt = _extract_appointment_details(page, log=log)
            result["appt_place"] = appt.get("place", "")
            result["appt_date"] = appt.get("date", "")
            result["appt_time"] = appt.get("time", "")
            result["appt_worker_name"] = appt.get("worker_name", "")
            result["appt_passport"] = appt.get("passport", "")
            result["appt_error"] = appt.get("error", "")
            # ชื่อบริษัท: reuse _pick_establishment/_pick_employer จาก structured_fields
            pairs = _pairs_from_structured(result.get("structured_fields") or {})
            estab_company, _prov, _sec = _pick_establishment(pairs)
            emp_company, _emp_prov = _pick_employer(pairs)
            result["appt_company"] = emp_company or estab_company
            # ชื่อแรงงาน: จาก iframe ก่อน, fallback structured_fields
            if not result["appt_worker_name"]:
                result["appt_worker_name"] = _worker_name_from_structured(pairs)
            # account username (สำหรับคอลัมน์ Username — ใช้ได้ทั้งโหมดเดี่ยว/หลายบัญชี)
            if cfg and not result.get("account_username"):
                result["account_username"] = cfg.get("username", "")
        except Exception as e:
            errors.append(f"appointment:{str(e).splitlines()[0][:120]}")

    # 3) บันทึกเพิ่มเติม — ทำหลังสุดเพราะอาจ navigate ไปหน้า edit
    if capture_extra_notes:
        try:
            result["extra_notes"] = _capture_extra_notes(page, row, cfg=cfg, log=log) or ""
        except Exception as e:
            errors.append(f"extra_notes:{e}")

    if errors:
        result["scrape_errors"] = " | ".join(errors)
    return result


def _capture_extra_notes(page: Page, row: dict | None = None, cfg: dict | None = None, log=print) -> str:
    """เก็บข้อความใน section 'บันทึกเพิ่มเติม' (รายการ <li>)
    1) ลองจากหน้า detail ปัจจุบัน
    2) ถ้าไม่เจอและมี row → navigate ไปหน้า edit แล้วเก็บ
    3) ถ้า session หมดระหว่างทาง → พยายาม login ใหม่ แล้วลองไปหน้า edit อีกครั้ง
    """
    js_extract = r"""() => {
      const heads = Array.from(document.querySelectorAll('*'))
        .filter(el => /บันทึกเพิ่มเติม/.test((el.innerText||'').slice(0,200)));
      for (const h of heads) {
        let cur = h;
        for (let i=0; i<8 && cur; i++) {
          const lis = cur.querySelectorAll('li');
          if (lis.length > 0) {
            const items = Array.from(lis).map(li => (li.innerText||'').replace(/\s+/g,' ').trim()).filter(Boolean);
            // dedupe
            const seen = new Set();
            return items.filter(x => { if (seen.has(x)) return false; seen.add(x); return true; });
          }
          cur = cur.parentElement;
        }
      }
      return [];
    }"""
    notes = page.evaluate(js_extract)
    if notes:
        return "\n".join(notes)

    # ลอง navigate ไปหน้า edit
    if row is not None:
        edit_url = build_edit_url(row)
        if edit_url:
            try:
                page.goto(edit_url, wait_until="domcontentloaded", timeout=20000)
                page.wait_for_timeout(4000)
                # ถ้า session หมดระหว่างขอ edit URL → ลอง login ใหม่แล้วเข้า edit URL อีกครั้ง
                if cfg is not None and _is_logged_out(page):
                    if ensure_session(page, cfg, target_url=edit_url, log=log):
                        page.wait_for_timeout(2000)
                    else:
                        return ""
                notes2 = page.evaluate(js_extract)
                if notes2:
                    return "\n".join(notes2)
            except Exception:
                pass

    return ""


def extract_label_value_pairs(page: Page) -> dict[str, str]:
    """ดึงคู่ label/value ทั้งหมดจากหน้า detail แบบทั่วไป (รองรับทุก form_type)
    Pattern จาก e-WorkPermit: <p class="label-form-info">label</p> ตามด้วย
    <p class="form-info ...">value</p> (อยู่ใกล้กันใน DOM order)
    คืน dict {label: value} — ถ้า label ซ้ำจะใส่ " | " คั่น

    ค่าเริ่มต้นจะรวม section header (จาก .head-step) เป็น prefix เพื่อไม่ให้
    label ที่ซ้ำชนกัน เช่น "เลขที่" ใน "ใบอนุญาตทำงาน" vs "หนังสือเดินทาง"
    """
    return page.evaluate(
        r"""() => {
          // เก็บทุก label-form-info + .head-step ใน document order
          // เลือกเฉพาะใน main content (ข้าม sidebar/nav)
          const root = document.querySelector('.tab-content, .container, main, body');
          if (!root) return {};
          const all = Array.from(root.querySelectorAll('.label-form-info, .form-info, .head-step'));
          const result = {};
          let curSection = '';
          for (let i = 0; i < all.length; i++) {
            const el = all[i];
            const cls = el.className || '';
            const text = (el.innerText || '').replace(/\s+/g, ' ').trim();
            if (!text) continue;
            if (cls.includes('head-step')) {
              curSection = text;
              continue;
            }
            if (cls.includes('label-form-info')) {
              // หา .form-info ตัวถัดไป
              let val = '';
              for (let j = i + 1; j < all.length; j++) {
                const next = all[j];
                const ncls = next.className || '';
                if (ncls.includes('head-step') || ncls.includes('label-form-info')) break;
                if (ncls.includes('form-info')) {
                  val = (next.innerText || '').replace(/\s+/g, ' ').trim();
                  break;
                }
              }
              if (!val) continue;
              const key = curSection ? `${curSection} | ${text}` : text;
              if (result[key]) {
                if (!result[key].split(' || ').includes(val)) {
                  result[key] = result[key] + ' || ' + val;
                }
              } else {
                result[key] = val;
              }
            }
          }
          return result;
        }"""
    )


def parse_alien_pane(text: str) -> dict[str, str]:
    """label จะตามด้วย value 1 บรรทัด (บางทีหลายบรรทัดถ้ามี enter ในที่อยู่)
    หยุดสะสม value ทันทีเมื่อเจอ label/section header ตัวถัดไป"""
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    result: dict[str, str] = {}
    i = 0
    while i < len(lines):
        label = lines[i]
        if label in KNOWN_LABELS:
            value_parts: list[str] = []
            j = i + 1
            while j < len(lines) and lines[j] not in STOP_TOKENS:
                value_parts.append(lines[j])
                j += 1
                if len(value_parts) >= 4:
                    break
            if value_parts:
                val = " | ".join(value_parts).strip()
                if label in result and val and result[label] != val:
                    result[label] = result[label] + " || " + val
                elif label not in result:
                    result[label] = val
            i = j
        else:
            i += 1
    return result


def save_excel(
    rows: list[dict], out_path: Path, fast: bool = False, include_account: bool = False,
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "WA Summary"

    base_cols = [
        "ลำดับ",
        *(["บัญชี (Username)"] if include_account else []),
        "เลขที่คำขอ",
        "ผู้ยื่น",
        "วันที่ยื่นคำขอ",
        "อัปเดตล่าสุด",
        "สถานะ",
        "หัวข้อแจ้งเตือน",
        "หมายเหตุ",
        "บันทึกเพิ่มเติม",
        "รายการ",
        "URL คำขอ",
        "คำขออนุญาต (รายละเอียด)",
        "ข้อผิดพลาดที่ข้าม",
    ]

    # เก็บลำดับ key จาก structured_fields ตามที่พบครั้งแรก (รักษา order ของ DOM)
    dynamic_cols: list[str] = []
    seen: set[str] = set()
    for r in rows:
        for k in (r.get("structured_fields") or {}).keys():
            if k not in seen:
                seen.add(k)
                dynamic_cols.append(k)

    # legacy ALIEN_FIELDS — รวมเฉพาะที่ยังไม่อยู่ใน dynamic_cols (เผื่อ MT_41_4_59)
    extra_legacy = [k for k in ALIEN_FIELDS if k not in seen]

    cols = base_cols + dynamic_cols + extra_legacy
    ws.append(cols)

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="C2185B")
    for c in range(1, len(cols) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)

    for idx, r in enumerate(rows, start=1):
        f = r.get("fields", {}) or {}
        sf = r.get("structured_fields", {}) or {}
        row_values = [
            idx,
            *([r.get("account_username", "")] if include_account else []),
            r.get("reqNo", ""),
            r.get("requester", ""),
            r.get("dateSubmit", ""),
            r.get("dateUpdate", ""),
            r.get("statusText", ""),
            r.get("alert_title", ""),
            r.get("note", ""),
            r.get("extra_notes", ""),
            r.get("desc", ""),
            build_detail_url(r),
            r.get("kham_kor_text", ""),
            r.get("scrape_errors", ""),
        ]
        row_values += [sf.get(k, "") for k in dynamic_cols]
        row_values += [f.get(k, "") for k in extra_legacy]
        ws.append(row_values)

    # ปรับความกว้างคอลัมน์ + wrap
    widths = {
        "บัญชี (Username)": 26,
        "เลขที่คำขอ": 18, "ผู้ยื่น": 25, "วันที่ยื่นคำขอ": 22,
        "อัปเดตล่าสุด": 22, "สถานะ": 30, "หัวข้อแจ้งเตือน": 30,
        "หมายเหตุ": 60, "บันทึกเพิ่มเติม": 60, "รายการ": 50,
        "URL คำขอ": 50, "คำขออนุญาต (รายละเอียด)": 80,
    }
    for c_idx, name in enumerate(cols, start=1):
        col_letter = get_column_letter(c_idx)
        ws.column_dimensions[col_letter].width = widths.get(name, 22)
    # การจัดรูปแบบ wrap_text ทีละเซลล์ช้ามากเมื่อแถวเยอะ → ข้ามตอนเซฟระหว่างทาง (fast)
    if not fast:
        for r_idx in range(2, ws.max_row + 1):
            for c_idx in range(1, ws.max_column + 1):
                ws.cell(row=r_idx, column=c_idx).alignment = Alignment(
                    vertical="top", wrap_text=True
                )
    ws.freeze_panes = "B2"
    # เปิด AutoFilter เพื่อให้ User กรองตามคอลัมน์ได้ทันที
    ws.auto_filter.ref = ws.dimensions
    wb.save(out_path)


def save_excel_appointment(
    rows: list[dict], out_path: Path, fast: bool = False, include_account: bool = False,
) -> None:
    """บันทึกรายงาน 'Template นัดหมาย' — 1 แถว/คำขอ
    คอลัมน์: ลำดับ | Username | เลขคำขอ | ชื่อบริษัท | ชื่อแรงงาน | วันที่นัดหมาย | เวลานัด | สถานที่
    (ตารางธรรมดา ไม่ใส่สีแถว — ผู้ใช้กรองสถานะ 'นัดหมายแล้ว' เองใน UI)

    include_account: มีไว้ให้ signature ตรงกับ save_excel (สลับฟังก์ชันได้) — ไม่มีผล
    เพราะ Template นี้มีคอลัมน์ Username เสมออยู่แล้ว
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "นัดหมาย"

    cols = [
        "ลำดับ", "Username", "เลขคำขอ", "ชื่อบริษัท", "ชื่อแรงงาน",
        "วันที่นัดหมาย", "เวลานัด", "สถานที่",
    ]
    ws.append(cols)

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="C2185B")
    for c in range(1, len(cols) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)

    for idx, r in enumerate(rows, start=1):
        ws.append([
            idx,
            r.get("account_username", ""),
            r.get("reqNo", ""),
            r.get("appt_company", ""),
            r.get("appt_worker_name", ""),
            r.get("appt_date", ""),
            r.get("appt_time", ""),
            r.get("appt_place", ""),
        ])

    widths = {
        "ลำดับ": 8, "Username": 26, "เลขคำขอ": 18, "ชื่อบริษัท": 40,
        "ชื่อแรงงาน": 28, "วันที่นัดหมาย": 20, "เวลานัด": 14, "สถานที่": 55,
    }
    for c_idx, name in enumerate(cols, start=1):
        ws.column_dimensions[get_column_letter(c_idx)].width = widths.get(name, 20)
    if not fast:
        for r_idx in range(2, ws.max_row + 1):
            for c_idx in range(1, ws.max_column + 1):
                ws.cell(row=r_idx, column=c_idx).alignment = Alignment(
                    vertical="top", wrap_text=True
                )
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(out_path)


# ─────────────────────────────────────────────────────────────────
# Checkpoint / Resume — กัน session timeout/โปรแกรมพังกลางทาง (8000+ รายการ)
#   เก็บแต่ละรายการที่ดึงเสร็จลง .progress.jsonl ทันที (append ทีละบรรทัด)
#   รันใหม่ครั้งหน้า → โหลด checkpoint → ข้ามรายการที่ทำแล้ว ทำต่อจากเดิม
# ─────────────────────────────────────────────────────────────────
# regex จับ timestamp ท้ายชื่อไฟล์ เช่น "_20260609_153012"
_TS_RE = re.compile(r"_\d{8}_\d{6}$")


def _timestamped_path(path: Path) -> Path:
    """แทรกวันที่+เวลาต่อท้ายชื่อไฟล์ (ก่อนนามสกุล) กันชื่อซ้ำ
    เช่น WA_report.xlsx → WA_report_20260609_153012.xlsx
    ถ้ามี timestamp อยู่แล้วจะแทนที่ของเดิม (ไม่ซ้อน)
    """
    path = Path(path)
    stem = _TS_RE.sub("", path.stem)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    new_path = path.with_name(f"{stem}_{ts}{path.suffix}")
    new_path.parent.mkdir(parents=True, exist_ok=True)
    return new_path


def _progress_path(out_path: Path) -> Path:
    # ตัด timestamp ออกจากชื่อ เพื่อให้ checkpoint คงที่ข้ามการรันหลายครั้ง (resume ได้)
    out_path = Path(out_path)
    stem = _TS_RE.sub("", out_path.stem)
    return out_path.parent / (stem + out_path.suffix + ".progress.jsonl")


def _chunk_path(out_path: Path, index: int) -> Path:
    """ชื่อไฟล์ part สำหรับการแยกไฟล์ทุก N รายการ
    เช่น WA_report_20260609_153012.xlsx → WA_report_20260609_153012_part01.xlsx
    """
    out_path = Path(out_path)
    return out_path.with_name(f"{out_path.stem}_part{index:02d}{out_path.suffix}")


def _load_progress(path: Path, log=print) -> dict[str, dict]:
    """อ่าน checkpoint → {reqNo: row_dict} ของรายการที่ดึงเสร็จแล้ว"""
    done: dict[str, dict] = {}
    if not path.exists():
        return done
    try:
        with path.open("r", encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if not line:
                    continue
                try:
                    rec = json.loads(line)
                except Exception:
                    continue
                key = str(rec.get("reqNo") or "")
                if key:
                    done[key] = rec
        if done:
            log(f"      ↻ พบ checkpoint เดิม: ทำไปแล้ว {len(done)} รายการ — จะทำต่อจากเดิม (ข้ามที่ทำแล้ว)")
    except Exception as e:
        log(f"      อ่าน checkpoint ไม่ได้ (เริ่มใหม่): {e}")
    return done


def _append_progress(path: Path, row: dict) -> None:
    """เขียนรายการที่ดึงเสร็จลง checkpoint ทันที (crash-safe)"""
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("a", encoding="utf-8") as fh:
            fh.write(json.dumps(row, ensure_ascii=False, default=str) + "\n")
            fh.flush()
    except Exception:
        pass


def _load_keyed_progress(path: Path, key_field: str = "ckey", log=print) -> dict[str, dict]:
    """โหลด checkpoint แบบ generic → {key: rec} ใช้กับโหมด results/receipts ที่ key เป็น
    'username|reqNo' (เก็บใน field ckey). ต่างจาก _load_progress ที่ key ด้วย reqNo ล้วน
    """
    done: dict[str, dict] = {}
    if not path.exists():
        return done
    try:
        with path.open("r", encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if not line:
                    continue
                try:
                    rec = json.loads(line)
                except Exception:
                    continue
                k = str(rec.get(key_field) or "")
                if k:
                    done[k] = rec
        if done:
            log(f"      ↻ พบความคืบหน้าเดิม: ทำเสร็จแล้ว {len(done)} คำขอ — จะข้ามและทำต่อจากเดิม")
    except Exception as e:
        log(f"      อ่านความคืบหน้าไม่ได้ (เริ่มใหม่): {e}")
    return done



def run_scrape(
    cfg: dict,
    out_path: Path,
    limit: int = 0,
    log=print,
    progress=None,
    is_cancelled=None,
    sink: list | None = None,
    write_output: bool = True,
) -> tuple[int, Path]:
    """รัน scraping ทั้ง pipeline. ใช้ได้ทั้ง CLI/GUI

    sink: ถ้าส่ง list มา จะ append ทุกแถวที่ดึงได้ (ติดแท็ก account_username) ลง list นี้ด้วย
    write_output: False = ไม่เขียนไฟล์ Excel ของตัวเอง (ยังทำ checkpoint ปกติ) ใช้ตอนรวมไฟล์

    Args:
        cfg: dict ต้องมี username, password, user_type, method, headless
        out_path: path ไฟล์ Excel
        limit: 0 = ทั้งหมด, >0 = จำกัด N รายการ
        log: callable(str) สำหรับเขียน log
        progress: callable(done:int, total:int) สำหรับอัปเดต progress
        is_cancelled: callable() -> bool คืน True เพื่อหยุดกลางคัน

    Returns: (จำนวน rows ที่ scrape, path ไฟล์ที่บันทึก)
    """
    out_path = _timestamped_path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    log(f"      ไฟล์รายงาน: {out_path.name}")

    # Template รายงาน: 'main' (ปกติ) หรือ 'appointment' (Template นัดหมาย)
    report_template = (cfg.get("report_template") or "main").strip().lower()
    capture_appointment = (report_template == "appointment")
    _wa_save = save_excel_appointment if capture_appointment else save_excel
    if capture_appointment:
        log("      รูปแบบรายงาน: Template นัดหมาย (จะเข้าแท็บการนัดหมายเพื่อดึงวันที่/เวลา/สถานที่)")

    def _cancelled() -> bool:
        return bool(is_cancelled and is_cancelled())

    with sync_playwright() as pw:
        browser = _launch_chromium(pw, cfg, ["--disable-blink-features=AutomationControlled"])
        ctx = browser.new_context(
            locale="th-TH", timezone_id="Asia/Bangkok",
            viewport={"width": 1500, "height": 900},
        )
        page = ctx.new_page()
        try:
            log("[1/4] กำลังเข้าสู่ระบบ...")
            login(page, cfg)
            log(f"      ล็อกอินสำเร็จ ({page.url})")

            log("[2/4] กำลังเปิดหน้า e-Tracking และตั้งค่าฟิลเตอร์...")
            goto_tracking(page)
            # กัน session หลุดทันทีหลัง login
            if _is_logged_out(page):
                ensure_session(page, cfg, log=log)
                goto_tracking(page)
            # ---- รวบรวมรายการ request_type ที่จะกรอง ----
            # priority: cfg["request_types"] (list, จาก multi-select GUI)
            #          > cfg["request_type"] (single string, backward compat)
            req_types_list: list[str] = []
            for c in (cfg.get("request_types") or []):
                s = (c or "").strip()
                if s and s not in {"0", "ALL", "all"} and s not in req_types_list:
                    req_types_list.append(s)
            if not req_types_list:
                single = (cfg.get("request_type") or "").strip()
                if single and single not in {"0", "ALL", "all"}:
                    req_types_list = [single]
            # log ให้เห็นชัด
            if not req_types_list:
                log("      ฟิลเตอร์รายการคำขอ: ทั้งหมด (ไม่กรอง)")
            elif len(req_types_list) == 1:
                log(f"      ฟิลเตอร์รายการคำขอ: {req_types_list[0]}")
            else:
                log(f"      ฟิลเตอร์รายการคำขอ ({len(req_types_list)} รายการ): "
                    + ", ".join(req_types_list))
            # req_type = code แรก (ใช้ต่อสำหรับ profile default + ensure_tracking_ready)
            req_type = req_types_list[0] if req_types_list else ""
            profile = get_profile(req_type)
            # Override จาก cfg (GUI) ถ้ามี
            status_ids = cfg.get("filter_status_ids") or profile.get("filter_status_ids") or ["WA"]
            capture_extra = cfg.get("capture_extra_notes")
            if capture_extra is None:
                capture_extra = profile.get("capture_extra_notes", True)
            log(f"      ติ๊ก checkbox สถานะ: {', '.join(status_ids)}")
            # log ช่วงวันที่ (ถ้ามี)
            date_from = (cfg.get("date_from") or "").strip()
            date_to = (cfg.get("date_to") or "").strip()
            if date_from or date_to:
                log(f"      วันที่ยื่นคำขอ: {date_from or '(ต้นสุด)'} → {date_to or '(ล่าสุด)'}")
            # ---- เก็บ rows ----
            # ถ้าเลือกหลายรายการ → วน filter ทีละ code แล้ว merge (dedup โดย reqNo)
            def _fetch_rows_for(rt_code: str) -> list[dict]:
                apply_wa_filter(
                    page, rt_code, status_ids=status_ids,
                    date_from=date_from, date_to=date_to,
                )
                out_rows = collect_all_wa_rows(page, log=log)
                if not out_rows and _is_logged_out(page):
                    if ensure_tracking_ready(page, cfg, rt_code, log=log):
                        out_rows = collect_wa_rows(page)
                return out_rows

            if len(req_types_list) <= 1:
                rows = _fetch_rows_for(req_type)
            else:
                merged: list[dict] = []
                seen: set = set()
                for _i, _code in enumerate(req_types_list, 1):
                    if _cancelled():
                        break
                    log(f"      [{_i}/{len(req_types_list)}] กำลังดึงรายการคำขอ: {_code}")
                    chunk = _fetch_rows_for(_code)
                    added = 0
                    for r in chunk:
                        k = str(r.get("reqNo") or "")
                        if k and k in seen:
                            continue
                        if k:
                            seen.add(k)
                        merged.append(r)
                        added += 1
                    log(f"          ← ได้ {len(chunk)} แถว (ใหม่ {added}, รวม {len(merged)})")
                rows = merged
            total_collected = len(rows)

            # กรองตาม status whitelist (global) — scrape detail เฉพาะ status ที่ระบุ
            whitelist = cfg.get("status_whitelist") or DEFAULT_STATUS_WHITELIST
            if whitelist:
                before = len(rows)
                rows = [r for r in rows if row_matches_whitelist(r, whitelist)]
                if before != len(rows):
                    log(f"      กรองตาม whitelist สถานะ: {before} → {len(rows)} รายการ")

            log(f"[3/4] พบรายการที่ตรงเงื่อนไข: {len(rows)} / {total_collected} รายการ")
            if limit and limit > 0:
                rows = rows[:limit]
                log(f"      จำกัดเฉพาะ {len(rows)} รายการแรก")

            total = len(rows)
            if progress:
                progress(0, total)

            # โหลด checkpoint เดิม (ถ้ามี) เพื่อทำต่อจากที่ค้างไว้
            prog_path = _progress_path(out_path)
            done_map = _load_progress(prog_path, log=log)

            scraped: list[dict] = []
            newly_done = 0
            # เซฟ Excel ทุก N รายการ (ปรับได้จาก cfg) — เซฟเฉพาะไฟล์ part ปัจจุบัน (เบา)
            SAVE_EVERY = max(1, int(cfg.get("save_every") or 100))
            # login ใหม่เชิงรุกทุก N รายการ กัน session/timeout หลุดในงานยาว (0 = ปิด)
            RELOGIN_EVERY = int(cfg.get("relogin_every", 500) or 0)
            # แยกไฟล์ทุก N รายการ — ช่วยให้เซฟเร็วคงที่ (ไม่ต้องเขียนทับไฟล์ใหญ่ทั้งก้อนซ้ำๆ)
            # 0 = ไฟล์เดียว (ไม่แยก)
            CHUNK_SIZE = int(cfg.get("chunk_size", 1000) or 0)
            chunk_rows: list[dict] = []   # รายการของไฟล์ part ปัจจุบัน
            chunk_index = 1
            chunk_files: list[Path] = []

            def _save_current_chunk(final: bool = False):
                """เซฟ chunk ปัจจุบันลงไฟล์ part (หรือไฟล์เดียวถ้าไม่ได้แยก)
                final=True → จัดรูปแบบเต็ม (ตอนปิดไฟล์), final=False → โหมดเบา
                """
                if not write_output or not chunk_rows:
                    return None
                p = _chunk_path(out_path, chunk_index) if CHUNK_SIZE else out_path
                _wa_save(chunk_rows, p, fast=not final)
                if p not in chunk_files:
                    chunk_files.append(p)
                return p

            for i, row in enumerate(rows, 1):
                if _cancelled():
                    log("[!] ยกเลิกโดยผู้ใช้ — กำลังบันทึกสิ่งที่ดึงได้แล้ว")
                    break
                key = str(row.get("reqNo") or "")
                # resume: ถ้ารายการนี้เคยดึงเสร็จแล้ว → ใช้ข้อมูลเดิม ข้ามการเข้าหน้า detail
                resumed = bool(key and key in done_map)
                if resumed:
                    row.update(done_map[key])
                else:
                    log(f"  [{i}/{total}] {row['reqNo']} - {row['requester']}")
                    # ถ้าหลุดจากหน้า tracking (อยู่หน้า detail ของ row ก่อนหน้า) → กลับไปก่อน
                    if "/Etracking" not in page.url:
                        try:
                            page.go_back(wait_until="domcontentloaded", timeout=15_000)
                            page.wait_for_timeout(1500)
                        except Exception:
                            pass
                        if "/Etracking" not in page.url:
                            ensure_tracking_ready(page, cfg, req_type, log=log)
                    try:
                        detail = scrape_detail(
                            page, row, cfg=cfg, log=log,
                            capture_extra_notes=capture_extra,
                            capture_appointment=capture_appointment,
                        )
                        # ถ้า session relogin ไม่ผ่าน → ลองฟื้น tracking แล้วเรียกใหม่ 1 ครั้ง
                        if "session:relogin_failed" in (detail.get("scrape_errors") or ""):
                            if ensure_tracking_ready(page, cfg, req_type, log=log):
                                log("      ↻ ลองดึงรายการนี้อีกครั้งหลังฟื้น session...")
                                detail = scrape_detail(
                                    page, row, cfg=cfg, log=log,
                                    capture_extra_notes=capture_extra,
                                    capture_appointment=capture_appointment,
                                )
                        row.update(detail)
                        if detail.get("scrape_errors"):
                            log(f"      ! ดึงบางส่วนไม่ได้ (ข้าม): {detail['scrape_errors']}")
                    except Exception as e:
                        # safety net — scrape_detail ควรไม่ throw แต่กันไว้
                        log(f"      !! ผิดพลาดร้ายแรง (ข้ามรายการนี้): {e}")
                        row.setdefault("scrape_errors", str(e))
                    # เขียน checkpoint ทันทีต่อรายการ (กันพังกลางทาง)
                    if key:
                        _append_progress(prog_path, row)
                    newly_done += 1

                if sink is not None:
                    row["account_username"] = cfg.get("username", "")
                    sink.append(row)
                scraped.append(row)
                chunk_rows.append(row)

                # เซฟ chunk ปัจจุบันเป็นระยะ (เบา) — เขียนเฉพาะไฟล์ part เล็กๆ ไม่ใช่ทั้งก้อน
                if len(chunk_rows) % SAVE_EVERY == 0:
                    try:
                        p = _save_current_chunk()
                        if p and (len(scraped) % 100 == 0 or i == total):
                            log(f"      💾 บันทึกความคืบหน้า: {len(scraped)} รายการ (ไฟล์: {p.name})")
                    except Exception as e:
                        log(f"      (เซฟระหว่างทางไม่สำเร็จ: {e})")

                # ปิดไฟล์ part เมื่อครบ CHUNK_SIZE แล้วขึ้นไฟล์ part ใหม่
                if CHUNK_SIZE and len(chunk_rows) >= CHUNK_SIZE:
                    try:
                        p = _save_current_chunk(final=True)
                        log(f"      📑 แยกไฟล์ที่ {chunk_index}: {p.name} ({len(chunk_rows)} รายการ)")
                    except Exception as e:
                        log(f"      (เซฟไฟล์ part ไม่สำเร็จ: {e})")
                    chunk_index += 1
                    chunk_rows = []

                # login ใหม่เชิงรุกทุก RELOGIN_EVERY รายการ — กัน session/timeout หลุดในงานยาว
                # ข้อมูลถึงตรงนี้ถูกเซฟ checkpoint + ไฟล์ part แล้ว จึงปลอดภัยที่จะ relogin
                if RELOGIN_EVERY and newly_done and newly_done % RELOGIN_EVERY == 0 and i < total:
                    log(f"      🔄 ครบ {newly_done} รายการ — login ใหม่เชิงรุกกัน session timeout...")
                    try:
                        _save_current_chunk()
                    except Exception:
                        pass
                    try:
                        login(page, cfg)
                        if ensure_tracking_ready(page, cfg, req_type, log=log):
                            log(f"      ✓ login ใหม่สำเร็จ — ทำต่อจากรายการที่ {i + 1}")
                        else:
                            log("      (กลับเข้าหน้า e-Tracking ไม่สำเร็จ — จะใช้กลไกฟื้นอัตโนมัติตามปกติ)")
                    except Exception as e:
                        log(f"      (login ใหม่เชิงรุกไม่สำเร็จ: {e} — จะใช้กลไกฟื้นอัตโนมัติตามปกติ)")
                if progress:
                    progress(i, total)

            # ปิดไฟล์ part สุดท้าย (ที่ยังไม่ครบ CHUNK_SIZE) ด้วยการจัดรูปแบบเต็ม
            last_path = _save_current_chunk(final=True)
            if last_path:
                log(f"[4/4] บันทึกไฟล์ Excel: {last_path}")
            if CHUNK_SIZE and len(chunk_files) > 1:
                log(f"[เสร็จสิ้น] รวม {len(scraped)} แถว — แยกเป็น {len(chunk_files)} ไฟล์:")
                for f in chunk_files:
                    log(f"      • {f.name}")
            else:
                log(f"[เสร็จสิ้น] รวม {len(scraped)} แถว")
            # ลบ checkpoint เมื่อทำครบ (ไม่ถูกยกเลิก) — รันใหม่ครั้งหน้าจะเริ่มสด
            if not _cancelled():
                try:
                    prog_path.unlink(missing_ok=True)
                except Exception:
                    pass
            else:
                log(f"      (เก็บ checkpoint ไว้ที่ {prog_path.name} — รันใหม่จะทำต่อจากเดิม)")
            # คืน path: ถ้าแยกหลายไฟล์ → คืนโฟลเดอร์ (เปิดดูไฟล์ทั้งหมดได้), ไม่งั้นคืนไฟล์เดียว
            result_path = (
                out_path.parent if (CHUNK_SIZE and len(chunk_files) > 1)
                else (chunk_files[0] if chunk_files else out_path)
            )
            return len(scraped), result_path
        finally:
            ctx.close(); browser.close()


def run_scrape_multi(
    cfg: dict,
    login_excel: Path,
    out_path: Path,
    limit: int = 0,
    log=print,
    progress=None,
    is_cancelled=None,
    combine: bool = False,
) -> tuple[int, Path]:
    """โหมด e-Tracking หลายบัญชี — วน login ทุกบัญชีใน UsernameLogin.xlsx

    combine=False (ค่าเริ่มต้น): แยกไฟล์ต่อบัญชี ({stem}_{username}{suffix}) แต่ละบัญชีมี checkpoint/resume
    combine=True: รวมทุกบัญชีเป็นไฟล์เดียว (เพิ่มคอลัมน์ 'บัญชี (Username)' เพื่อแยกแถว)

    Args:
        cfg: ตัวเลือกรวม (headless, request_type, filter_status_ids, ฯลฯ) — ไม่ต้องมี username/password
        login_excel: ไฟล์ UsernameLogin.xlsx (คอลัมน์ Username, Password, Type)
        out_path: ไฟล์ฐาน

    Returns: (จำนวนแถวรวมทุกบัญชี, path ผลลัพธ์ — ไฟล์รวม ถ้า combine, ไม่งั้นเป็นโฟลเดอร์)
    """
    accounts = _read_login_accounts(login_excel)
    if not accounts:
        raise ValueError("ไม่พบบัญชีในไฟล์ UsernameLogin.xlsx")
    out_path = Path(out_path)
    out_dir = out_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = out_path.stem
    suffix = out_path.suffix or ".xlsx"
    n = len(accounts)
    report_template = (cfg.get("report_template") or "main").strip().lower()
    _wa_save = save_excel_appointment if report_template == "appointment" else save_excel

    if combine:
        # ── รวมทุกบัญชีเป็นไฟล์เดียว ──
        combined_path = _timestamped_path(out_path)
        all_rows: list[dict] = []
        ok_accounts = 0
        log(f"[Multi] เริ่มดึง e-Tracking จาก {n} บัญชี → รวมเป็นไฟล์เดียว: {combined_path.name}")
        for gi, (_ukey, acct) in enumerate(accounts.items(), start=1):
            if is_cancelled and is_cancelled():
                log("[!] ผู้ใช้ยกเลิก — หยุด"); break
            login_cfg = dict(cfg)
            login_cfg["username"] = acct["username"]
            login_cfg["password"] = acct["password"]
            login_cfg["user_type"] = acct["type"]
            if acct.get("method"):
                login_cfg["method"] = acct["method"]
            safe_user = (_receipt_safe_name(acct["username"]).strip() or f"user{gi}")
            # ใช้ per-account path เพื่อแยก checkpoint (ไม่สร้างไฟล์ xlsx เพราะ write_output=False)
            per_ckpt = out_dir / f"{stem}_{safe_user}{suffix}"
            before = len(all_rows)
            log(f"\n===== บัญชี {gi}/{n}: {acct['username']} ({acct['type']}) =====")
            try:
                run_scrape(
                    login_cfg, per_ckpt, limit=limit,
                    log=log, progress=progress, is_cancelled=is_cancelled,
                    sink=all_rows, write_output=False,
                )
                ok_accounts += 1
                log(f"      + บัญชีนี้ได้ {len(all_rows) - before} แถว (รวมสะสม {len(all_rows)} แถว)")
                # เซฟไฟล์รวมหลังจบแต่ละบัญชี (กันข้อมูลหายถ้าบัญชีถัดไปพัง)
                try:
                    _wa_save(all_rows, combined_path, include_account=True)
                except Exception as e:
                    log(f"      (เซฟไฟล์รวมระหว่างทางไม่สำเร็จ: {e})")
            except Exception as e:
                log(f"   ✗ บัญชี {acct['username']} ล้มเหลว: {e} — ข้ามไปบัญชีถัดไป")
                continue
        _wa_save(all_rows, combined_path, include_account=True)
        log(f"\n[Multi] เสร็จสิ้น — รวม {len(all_rows)} แถว จาก {ok_accounts}/{n} บัญชี → {combined_path}")
        return len(all_rows), combined_path

    # ── แยกไฟล์ต่อบัญชี (ค่าเริ่มต้นเดิม) ──
    total_rows = 0
    produced: list[Path] = []
    log(f"[Multi] เริ่มดึง e-Tracking จาก {n} บัญชี (แยกไฟล์ต่อบัญชี, ไฟล์ login: {login_excel.name})")
    for gi, (_ukey, acct) in enumerate(accounts.items(), start=1):
        if is_cancelled and is_cancelled():
            log("[!] ผู้ใช้ยกเลิก — หยุด"); break
        login_cfg = dict(cfg)
        login_cfg["username"] = acct["username"]
        login_cfg["password"] = acct["password"]
        login_cfg["user_type"] = acct["type"]
        # ระบบ login ต่อบัญชีจากคอลัมน์ 'ระบบ' ใน Excel (ถ้าระบุ)
        if acct.get("method"):
            login_cfg["method"] = acct["method"]
        safe_user = (_receipt_safe_name(acct["username"]).strip() or f"user{gi}")
        per_out = out_dir / f"{stem}_{safe_user}{suffix}"
        log(f"\n===== บัญชี {gi}/{n}: {acct['username']} ({acct['type']}) =====")
        try:
            cnt, _p = run_scrape(
                login_cfg, per_out, limit=limit,
                log=log, progress=progress, is_cancelled=is_cancelled,
            )
            total_rows += cnt
            produced.append(per_out)
        except Exception as e:
            log(f"   ✗ บัญชี {acct['username']} ล้มเหลว: {e} — ข้ามไปบัญชีถัดไป")
            continue
    log(f"\n[Multi] เสร็จสิ้น — รวม {total_rows} แถว จาก {len(produced)} บัญชี (โฟลเดอร์: {out_dir})")
    return total_rows, out_dir


# ─────────────────────────────────────────────────────────────────
# โหมด — ดึงรายงาน "ข้อมูลการขออนุญาต"
#   กรองสถานะคำขอ → เปิดหน้า detail ทีละรายการ → เก็บ:
#     • Tab "สถานะคำขอ" (#tab_default_1)  → รายการ "อนุมัติคำขอ" (วันที่ + ผล)
#     • Tab "คำขออนุญาต" (#tab_default_2) → บริษัท + จังหวัด จาก section
#       "สถานประกอบการ" (ไม่ใช่ "ข้อมูลนายจ้าง")
# ─────────────────────────────────────────────────────────────────

# regex จับ timestamp ไทยในไทม์ไลน์ เช่น "16 ก.ค. 2026 , 10:06"
_TH_FLOW_DATE_RE = re.compile(
    r"(\d{1,2}\s+[\u0e01-\u0e59.]+\s+\d{4}\s*,?\s*\d{1,2}:\d{2})"
)

# JS ดึงคู่ label/value + ข้อความดิบ จากแท็บ "คำขออนุญาต" (#tab_default_2)
# ใช้โครงสร้างจริงของหน้า: .head-step = หัวข้อ section, .label-form-info = label,
# .form-info = value  (เหมือน extract_label_value_pairs)
# คืนรายการแท็บทั้งหมดบนหน้า detail: [{label, pane}] — ใช้ map ชื่อแท็บ → id ของ pane
# (ตำแหน่ง tab_default_N ไม่คงที่: 'คำขออนุญาต' เป็น #tab_default_2 บางฟอร์ม, #tab_default_3 บางฟอร์ม)
_DETAIL_TABS_JS = r"""() => {
  const norm = s => (s || '').replace(/\s+/g, ' ').trim();
  const links = Array.from(document.querySelectorAll(
    'a[href^="#tab_default"], a[data-toggle="tab"][href^="#"], a[role="tab"][href^="#"]'
  ));
  const out = [];
  const seen = new Set();
  for (const a of links) {
    const href = (a.getAttribute('href') || '').replace(/^#/, '');
    if (!href || seen.has(href)) continue;
    seen.add(href);
    // index = ลำดับแท็บ (0-based) → ใช้ fallback อ่าน .tab-pane ตำแหน่งเดียวกัน
    // เมื่อ pane ไม่มี id (ฟอร์ม Change/Forms: nav ชี้ #tab-request แต่ pane ไม่มี id)
    out.push({label: norm(a.innerText || a.textContent || ''), pane: href, index: out.length});
  }
  return out;
}"""

# ดึง label/value pairs จาก pane ที่ระบุ (paneId) — selector ตาม pattern จริงของ e-WorkPermit
# (.head-step = หัวข้อ section, .label-form-info = label, .form-info = value) — ยืนยันจาก probe DOM
_PERMIT_TAB_JS = r"""(args) => {
  const paneId = (args && args.paneId) || '';
  const index = (args && args.index != null) ? args.index : -1;
  const norm = s => (s || '').replace(/\s+/g, ' ').trim();
  // หา pane: ตาม id ก่อน (ฟอร์มเก่า) ถ้าไม่มี ใช้ .tab-pane ตำแหน่ง index (ฟอร์ม Change)
  let pane = paneId ? document.getElementById(paneId) : null;
  if (!pane && index >= 0) {
    const panes = document.querySelectorAll('.tab-pane');
    if (index < panes.length) pane = panes[index];
  }
  if (!pane) return {raw: '', pairs: [], paneRef: ''};
  // เปิด accordion ที่พับอยู่ (ปุ่ม 'แสดงทั้งหมด') — best-effort, ไม่คลิกลิงก์ที่นำทางออก
  try {
    Array.from(pane.querySelectorAll('a,button,span,div,i')).forEach(b => {
      if (norm(b.textContent || '') === 'แสดงทั้งหมด') {
        const href = b.getAttribute('href');
        if (b.tagName !== 'A' || !href || href === '#') { try { b.click(); } catch (e) {} }
      }
    });
  } catch (e) {}
  const all = Array.from(pane.querySelectorAll('.head-step, .label-form-info, .form-info, [data-id="employerAddr.title"], [data-id="employerAddr.address"]'));
  const pairs = [];
  let cur = '';
  let pendingAddrLabel = '';
  for (let i = 0; i < all.length; i++) {
    const el = all[i];
    const cls = el.className || '';
    const did = el.getAttribute('data-id') || '';
    const text = norm(el.innerText || el.textContent || '');
    if (!text && did !== 'employerAddr.address') continue;
    if (cls.includes('head-step')) { cur = text; continue; }
    // บล็อกที่อยู่สถานประกอบการ (data-id="employerAddr.*") — โครงสร้างต่างจาก label/form-info
    // ที่อยู่จริงอยู่ใน attribute data_th (เช่น '...จังหวัด ชลบุรี...')
    if (did === 'employerAddr.title') { pendingAddrLabel = text; continue; }
    if (did === 'employerAddr.address') {
      const aval = norm(el.getAttribute('data_th') || el.innerText || el.textContent || '');
      if (aval) pairs.push({section: cur, label: pendingAddrLabel || 'สถานประกอบการ', value: aval});
      pendingAddrLabel = '';
      continue;
    }
    if (cls.includes('label-form-info')) {
      let val = '';
      for (let j = i + 1; j < all.length; j++) {
        const n = all[j];
        const nc = n.className || '';
        const ndid = n.getAttribute('data-id') || '';
        if (nc.includes('head-step') || nc.includes('label-form-info')) break;
        if (ndid === 'employerAddr.title' || ndid === 'employerAddr.address') break;
        if (nc.includes('form-info')) { val = norm(n.innerText || n.textContent || ''); break; }
      }
      pairs.push({section: cur, label: text, value: val});
    }
  }
  const raw = norm(pane.innerText || pane.textContent || '').slice(0, 12000);
  const paneRef = (paneId && document.getElementById(paneId)) ? ('#' + paneId) : ('.tab-pane[' + index + ']');
  return {raw, pairs, paneRef};
}"""


def parse_status_flow(raw: str) -> list[dict]:
    """แยกข้อความไทม์ไลน์ 'สถานะคำขอ' → รายการ [{date, text, step, result}] เรียงตามที่พบ
    ข้อความดิบมักเป็น: '<date>\nขั้นตอน\nผล\n<date>\nขั้นตอน\nผล...'
      • text   = ขั้นตอน + ผล (รวมเป็นช่องว่างเดียว) — ใช้ค้นหา 'อนุมัติคำขอ'
      • step   = บรรทัดแรกของ body (ชื่อขั้นตอน เช่น 'อนุมัติคำขอ')
      • result = บรรทัดที่เหลือ (ผล เช่น 'ผ่านการอนุมัติคำขอ')
    """
    if not raw:
        return []
    parts = _TH_FLOW_DATE_RE.split(raw.strip())
    # parts = [prefix, date1, body1, date2, body2, ...]
    entries: list[dict] = []
    i = 1
    while i < len(parts) - 1:
        date_txt = (parts[i] or "").strip()
        body_raw = (parts[i + 1] or "").strip(" :\n")
        body = re.sub(r"\s+", " ", body_raw).strip(" :\n")
        seg = [ln.strip() for ln in body_raw.splitlines() if ln.strip()]
        step = seg[0] if seg else body
        result = " ".join(seg[1:]).strip() if len(seg) > 1 else ""
        if date_txt:
            entries.append({"date": date_txt, "text": body, "step": step, "result": result})
        i += 2
    return entries


def _format_status_all(entries: list[dict]) -> str:
    """รวมทุกสถานะเป็นข้อความหลายบรรทัด — บรรทัดละ 'วันที่ -> ขั้นตอน -> ผล'
    (เรียงตามที่เว็บแสดง = ใหม่ไปเก่า). ช่องที่ว่างจะถูกข้าม"""
    lines: list[str] = []
    for e in entries:
        parts = [
            (e.get("date", "") or "").strip(),
            (e.get("step", "") or "").strip(),
            (e.get("result", "") or "").strip(),
        ]
        parts = [p for p in parts if p]
        if parts:
            lines.append(" -> ".join(parts))
    return "\n".join(lines)


def _find_approve_entry(entries: list[dict]) -> dict | None:
    """หา entry ของขั้นตอน 'อนุมัติคำขอ' (ไม่ใช่ 'พิจารณาคำขอ'/'รออนุมัติ')"""
    for e in entries:
        t = e.get("text", "") or ""
        # ต้องมีคำว่า 'อนุมัติคำขอ' และไม่ใช่ขั้น 'รออนุมัติ...' (ยังไม่อนุมัติ)
        if "อนุมัติคำขอ" in t and not t.strip().startswith("รอ"):
            return e
    return None


def _label_is_company(lab: str) -> bool:
    lab = (lab or "").strip().rstrip(":").strip()
    return (
        "ชื่อสถานประกอบการ" in lab
        or "ชื่อบริษัท" in lab
        or "ชื่อสถานที่" in lab
        or "ชื่อผู้ประกอบการ" in lab
        or "ชื่อนิติบุคคล" in lab
        # นิติบุคคล → ชื่อจริงอยู่ใต้ 'ชื่อหน่วยงาน' ; สมาคม/มูลนิธิ → 'ชื่อสมาคม...'
        or "ชื่อหน่วยงาน" in lab
        or "ชื่อสมาคม" in lab
        or "ชื่อมูลนิธิ" in lab
        or lab in ("บริษัท", "สถานประกอบการ", "ชื่อสถานประกอบการ")
    )


def _label_is_province(lab: str) -> bool:
    """label ที่เป็น 'จังหวัด' ของสถานประกอบการจริง ๆ — ไม่เอา
    'จังหวัดที่เข้ารับการอบรม...', 'ด่านเข้าเมือง...' ฯลฯ"""
    lab = (lab or "").strip().rstrip(":").strip()
    return lab in ("จังหวัด", "จังหวัด/เขต", "จังหวัดที่ตั้ง", "จังหวัดสถานประกอบการ")


def _province_from_addr(val: str) -> str:
    """ดึงชื่อจังหวัดจากสตริงที่อยู่ เช่น
    '... เขต/อำเภอ วังทองหลาง จังหวัด กรุงเทพมหานคร รหัสไปรษณีย์ 10310' → 'กรุงเทพมหานคร'
    (ชื่อจังหวัดไทยเป็นคำเดียวไม่มีช่องว่าง)"""
    if not val or "จังหวัด" not in val:
        return ""
    m = re.search(r"จังหวัด\s*([^\s]+)", val)
    if not m:
        return ""
    return m.group(1).strip(" :\u200b")


def _pick_establishment(pairs: list[dict]) -> tuple[str, str, str]:
    """เลือก บริษัท + จังหวัด จาก section สถานประกอบการ/สถานที่ทำงาน/ผู้รับอนุญาต
    โดยตัด 'ข้อมูลนายจ้าง' ออกเสมอ. คืน (company, province, section_used)

    รองรับหลาย form_type ตามที่ probe DOM จริงเจอ:
      • MOU renewal → section 'ข้อมูลผู้รับอนุญาตให้นำคนต่างด้าวมาทำงาน'
        (ชื่อสถานประกอบการ(ไทย) + จังหวัดฝังใน 'ที่อยู่สถานที่ทำงาน/สาขา')
      • Name List/นำเข้า → section 'สถานที่ทำงานของคนต่างด้าว' (label 'จังหวัด' แยกช่อง)
      • ฟอร์มนายจ้างตรง → section 'สถานประกอบการ'
    """
    def is_estab(sec: str) -> bool:
        s = sec or ""
        return any(
            k in s
            for k in (
                "สถานประกอบการ", "สถานที่ทำงาน", "สถานที่ประกอบการ",
                "สถานที่ตั้ง", "ผู้รับอนุญาต",
            )
        )

    def is_employer(sec: str) -> bool:
        return "นายจ้าง" in (sec or "")

    company = province = used_section = ""

    def scan(strict: bool) -> None:
        nonlocal company, province, used_section
        for p in pairs:
            sec = p.get("section", "") or ""
            lab = p.get("label", "") or ""
            val = (p.get("value", "") or "").strip()
            if not val:
                continue
            if is_employer(sec):
                continue  # ข้าม section 'ข้อมูลนายจ้าง' เสมอ
            if strict and not is_estab(sec):
                continue
            if not company and _label_is_company(lab):
                company, used_section = val, sec
            if not province and _label_is_province(lab):
                province, used_section = val, (used_section or sec)
            if not province and ("ที่อยู่" in lab or "สาขา" in lab or "สถานที่" in lab):
                prov = _province_from_addr(val)
                if prov:
                    province, used_section = prov, (used_section or sec)

    scan(strict=True)       # PASS 1: เฉพาะ section สถานประกอบการชัดเจน
    if not company or not province:
        scan(strict=False)  # PASS 2: ผ่อน — ทุก section ที่ไม่ใช่ 'นายจ้าง'
    return company.strip(), province.strip(), used_section.strip()


def _pick_employer(pairs: list[dict]) -> tuple[str, str]:
    """ดึงชื่อบริษัท + จังหวัด จาก section 'ข้อมูลนายจ้าง' โดยเฉพาะ
    (แยกคอลัมน์จากผู้รับอนุญาต/สถานประกอบการ ให้ผู้ใช้เลือกเองใน Excel).
    คืน (employer_company, employer_province)

    บาง form_type (เช่น MOU renewal ที่ 'ผู้รับอนุญาต' ว่าง) ชื่อบริษัทจริง
    อยู่ใต้ 'ข้อมูลนายจ้าง' เท่านั้น — เก็บไว้คนละช่องเผื่อผู้ใช้ต้องการ
    """
    company = province = ""
    for p in pairs:
        sec = p.get("section", "") or ""
        if "นายจ้าง" not in sec:
            continue
        lab = p.get("label", "") or ""
        val = (p.get("value", "") or "").strip()
        if not val:
            continue
        # ใน section 'ข้อมูลนายจ้าง' กรณีบุคคลธรรมดา ชื่ออยู่ใต้ 'ชื่อนายจ้าง (ไทย)'
        if not company and (_label_is_company(lab) or "ชื่อนายจ้าง" in lab):
            company = val
        if not province and _label_is_province(lab):
            province = val
        if not province and ("ที่อยู่" in lab or "สาขา" in lab or "สถานที่" in lab):
            prov = _province_from_addr(val)
            if prov:
                province = prov
    return company.strip(), province.strip()


def _pairs_from_structured(sf: dict) -> list[dict]:
    """แปลง structured_fields (dict คีย์ 'section | label') กลับเป็น list ของ
    {section, label, value} เพื่อใช้กับ _pick_establishment / _pick_employer ซ้ำได้
    """
    pairs: list[dict] = []
    for key, val in (sf or {}).items():
        if " | " in key:
            sec, lab = key.split(" | ", 1)
        else:
            sec, lab = "", key
        pairs.append({"section": sec, "label": lab, "value": val})
    return pairs


def _worker_name_from_structured(pairs: list[dict]) -> str:
    """ดึงชื่อแรงงาน (คนต่างด้าว) จาก structured_fields — ใช้เป็น fallback เมื่อ
    iframe การนัดหมายไม่มีชื่อ. ลอง section 'คนต่างด้าว/แรงงาน' + label ที่มีคำว่า 'ชื่อ'
    (ข้ามชื่อสถานประกอบการ/นายจ้าง/บริษัท)
    """
    def scan(latin_only: bool) -> str:
        for p in pairs:
            sec = p.get("section", "") or ""
            lab = p.get("label", "") or ""
            val = (p.get("value", "") or "").strip()
            if not val:
                continue
            if "คนต่างด้าว" not in sec and "แรงงาน" not in sec:
                continue
            if "ชื่อ" not in lab:
                continue
            if any(k in lab for k in ("สถานประกอบการ", "นายจ้าง", "บริษัท")):
                continue
            if latin_only and not re.search(r"[A-Za-z]", val):
                continue
            return val
        return ""
    return scan(True) or scan(False)


def _pick_workplace_province(pairs: list[dict]) -> str:
    """จังหวัดจาก section 'ข้อมูลการขออนุญาต' (ที่อยู่สถานประกอบการ/สถานที่ทำงาน) เท่านั้น
    ตามที่ผู้ใช้กำหนด — ไม่ดึงจาก 'ข้อมูลนายจ้าง'. คืนชื่อจังหวัด ('' ถ้าไม่มี)

    • MOU renewal: section 'ข้อมูลการขออนุญาต' → label 'สถานที่ทำงาน/สาขา'
      (จังหวัดฝังในที่อยู่ '...จังหวัด พระนครศรีอยุธยา รหัสไปรษณีย์...')
    • Name List: section 'ข้อมูลการขออนุญาต'/'สถานที่ทำงานของคนต่างด้าว' → label 'จังหวัด' แยกช่อง
    • ฟอร์ม Change ที่ไม่มี section นี้ → คืน '' (จังหวัดว่าง ตามที่ผู้ใช้ระบุ 'เท่านั้น')
    """
    def want_section(sec: str) -> bool:
        s = sec or ""
        if "นายจ้าง" in s:
            return False
        return any(k in s for k in ("ข้อมูลการขออนุญาต", "สถานประกอบการ", "สถานที่ทำงาน"))

    # PASS 1: label 'จังหวัด' แยกช่อง ใน section ที่ต้องการ
    for p in pairs:
        if not want_section(p.get("section", "")):
            continue
        if _label_is_province(p.get("label", "")):
            val = (p.get("value", "") or "").strip(" :\u200b")
            if val:
                return val
    # PASS 2: จังหวัดฝังในค่าที่อยู่ — สแกน 'ค่า' ทุกช่องใน section ที่ต้องการ
    #         (รองรับกรณี label ไม่สื่อความ เช่น '-' หรือว่าง: บางฟอร์มแสดงที่อยู่
    #          สถานประกอบการเป็นบรรทัดเดียว label='-' → เดิม PASS2 ที่เช็คเฉพาะ label จะข้ามไป)
    #         _province_from_addr คืนค่าเฉพาะเมื่อพบ 'จังหวัด <ชื่อ>' จึงปลอดภัยกับช่องที่ไม่ใช่ที่อยู่
    for p in pairs:
        if not want_section(p.get("section", "")):
            continue
        prov = _province_from_addr((p.get("value", "") or "").strip())
        if prov:
            return prov
    return ""


def _resolve_detail_tabs(page: Page) -> list[dict]:
    """คืนรายการแท็บบนหน้า detail: [{label, pane}] — ใช้ map ชื่อแท็บ → id ของ pane
    (ตำแหน่ง tab_default_N ไม่คงที่ แต่ละ form_type ต่างกัน)"""
    try:
        tabs = page.evaluate(_DETAIL_TABS_JS) or []
    except Exception:
        tabs = []
    return [t for t in tabs if isinstance(t, dict) and t.get("pane")]


def _find_pane_id(tabs: list[dict], *needles: str) -> str:
    """หา pane id จาก label ของแท็บ — 'ตรงเป๊ะ' ก่อน แล้วค่อย 'มีคำนั้นอยู่'"""
    for n in needles:
        for t in tabs:
            if (t.get("label") or "").strip() == n:
                return t.get("pane") or ""
    for n in needles:
        for t in tabs:
            if n in (t.get("label") or ""):
                return t.get("pane") or ""
    return ""


def _find_tab(tabs: list[dict], *needles: str) -> dict:
    """หา tab dict {label, pane, index} จาก label — 'ตรงเป๊ะ' ก่อน แล้วค่อย 'มีคำนั้นอยู่'.
    คืนทั้ง dict เพื่อให้ได้ทั้ง pane id และ index (ใช้ fallback เมื่อ pane ไม่มี id)"""
    for n in needles:
        for t in tabs:
            if (t.get("label") or "").strip() == n:
                return t
    for n in needles:
        for t in tabs:
            if n in (t.get("label") or ""):
                return t
    return {}


def scrape_permit_detail(page: Page, row: dict, cfg: dict | None = None, log=print) -> dict:
    """เปิดหน้า detail แล้วเก็บ:
      • Tab 'สถานะคำขอ'   → รายการ 'อนุมัติคำขอ' (วันที่ + ผล)
      • Tab 'คำขออนุญาต'  → บริษัท + จังหวัด จาก 'สถานประกอบการ' (ไม่เอา 'ข้อมูลนายจ้าง')
    หมายเหตุ: id ของ pane (#tab_default_N) ไม่คงที่ — 'คำขออนุญาต' เป็น #tab_default_2 บางฟอร์ม,
    #tab_default_3 บางฟอร์ม จึง resolve จาก 'ชื่อแท็บ' แทนการ hard-code id
    แต่ละ step มี try/except แยก — ถ้าหน้านี้ไม่มี tab/section ที่คาดหวัง จะข้ามเฉพาะส่วนนั้น
    """
    result: dict = {
        "approve_found": "",
        "approve_date": "",
        "approve_result": "",
        "status_all": "",
        "status_flow_raw": "",
        "company_main": "",
        "estab_company": "",
        "employer_company": "",
        "estab_province": "",
        "estab_section": "",
        "permit_raw": "",
        "tabs_read": "",
        "scrape_errors": "",
    }
    errors: list[str] = []

    # ---- นำทางเข้าหน้า detail (navigate ตรงผ่าน URL, fallback คลิกแถว/openDetail) ----
    ok = _open_detail_direct(page, row)
    if not ok and cfg is not None and _is_logged_out(page):
        if ensure_session(page, cfg, log=log):
            ok = _open_detail_direct(page, row)
    if not ok:
        errors.append("open_detail_failed")
        result["scrape_errors"] = " | ".join(errors)
        return result

    # ---- resolve ชื่อแท็บ → pane (id หรือ index) — โครงสร้างต่างกันตาม form_type ----
    #   ฟอร์มเก่า (MOU/MT_41): pane มี id #tab_default_N
    #   ฟอร์ม Change/Forms: pane ไม่มี id → map จากตำแหน่ง (index) ของ nav link
    tabs = _resolve_detail_tabs(page)
    status_tab = _find_tab(tabs, "สถานะคำขอ")
    permit_tab = _find_tab(tabs, "คำขออนุญาต")
    status_pane = status_tab.get("pane") or "tab_default_1"
    status_index = status_tab.get("index", 0)
    permit_pane = permit_tab.get("pane") or "tab_default_2"
    permit_index = permit_tab.get("index", -1)
    result["tabs_read"] = (
        f"สถานะคำขอ→#{status_pane}[{status_index}] | "
        f"คำขออนุญาต→#{permit_pane}[{permit_index}]"
    )

    # คลิกแท็บ: หา nav link จาก href=id ก่อน ถ้าไม่มีใช้ nav link ตำแหน่ง index
    _click_tab_js = r"""(args) => {
      let a = args.paneId ? document.querySelector('a[href="#' + args.paneId + '"]') : null;
      if (!a && args.index != null && args.index >= 0) {
        const links = Array.from(document.querySelectorAll(
          'a[href^="#tab_default"], a[data-toggle="tab"][href^="#"], a[role="tab"][href^="#"]'
        ));
        const seen = new Set(); const navs = [];
        for (const x of links) {
          const h = (x.getAttribute('href') || '').replace(/^#/, '');
          if (!h || seen.has(h)) continue; seen.add(h); navs.push(x);
        }
        if (args.index < navs.length) a = navs[args.index];
      }
      if (a) a.click();
    }"""

    # ---- 1) แท็บ 'สถานะคำขอ' → ไทม์ไลน์ → 'อนุมัติคำขอ' ----
    try:
        page.evaluate(_click_tab_js, {"paneId": status_pane, "index": status_index})
        page.wait_for_timeout(800)
        raw_flow = page.evaluate(
            r"""(args) => {
              let pane = args.paneId ? document.getElementById(args.paneId) : null;
              if (!pane && args.index != null && args.index >= 0) {
                const panes = document.querySelectorAll('.tab-pane');
                if (args.index < panes.length) pane = panes[args.index];
              }
              if (!pane) return '';
              return (pane.innerText || pane.textContent || '')
                .replace(/[ \t]+/g, ' ').replace(/\n{2,}/g, '\n').trim().slice(0, 8000);
            }""",
            {"paneId": status_pane, "index": status_index},
        ) or ""
        result["status_flow_raw"] = raw_flow
        entries = parse_status_flow(raw_flow)
        result["status_all"] = _format_status_all(entries)
        appr = _find_approve_entry(entries)
        if appr:
            result["approve_found"] = "พบ"
            result["approve_date"] = appr.get("date", "")
            t = appr.get("text", "") or ""
            m = re.match(r"\s*อนุมัติคำขอ\s*(.*)$", t)
            result["approve_result"] = (m.group(1).strip() if m else t) or t
        else:
            result["approve_found"] = "ไม่พบ"
    except Exception as e:
        errors.append(f"tab_status:{str(e).splitlines()[0][:120]}")

    # ---- 2) แท็บ 'คำขออนุญาต' → สถานประกอบการ (บริษัท/จังหวัด) ----
    try:
        page.evaluate(_click_tab_js, {"paneId": permit_pane, "index": permit_index})
        page.wait_for_timeout(1000)
        data = page.evaluate(_PERMIT_TAB_JS, {"paneId": permit_pane, "index": permit_index}) or {}
        result["permit_raw"] = data.get("raw", "") or ""
        pairs = data.get("pairs") or []
        company, province, section = _pick_establishment(pairs)
        emp_company, emp_province = _pick_employer(pairs)
        result["estab_company"] = company
        result["employer_company"] = emp_company
        # บริษัทหลัก: ใช้ชื่อจาก 'นายจ้าง' ก่อน (ข้อมูลจริงชื่อบริษัทอยู่ใต้ section นายจ้าง)
        # ถ้าว่างจริง ๆ ค่อย fallback ไปผู้รับอนุญาต/สถานประกอบการ
        result["company_main"] = emp_company or company
        # จังหวัด: เอาเฉพาะจาก section 'ข้อมูลการขออนุญาต' (สถานประกอบการ/สถานที่ทำงาน) เท่านั้น
        # ไม่ดึงจาก 'ข้อมูลนายจ้าง' (อาจคนละจังหวัดกับที่ตั้งสถานประกอบการจริง)
        result["estab_province"] = _pick_workplace_province(pairs)
        result["estab_section"] = section
    except Exception as e:
        errors.append(f"tab_permit:{str(e).splitlines()[0][:120]}")

    if errors:
        result["scrape_errors"] = " | ".join(errors)
    return result


def save_excel_permit_report(rows: list[dict], out_path: Path) -> None:
    """บันทึกรายงาน 'ข้อมูลการขออนุญาต' ลง Excel (1 แถว/คำขอ)
    คอลัมน์ตามที่ผู้ใช้กำหนด: บริษัท | จังหวัด | เลขคำขอ | สถานะคำขอ | หมายเหตุ
      • สถานะคำขอ = ทุกสถานะที่แสดงบนไทม์ไลน์ (วันที่ -> ขั้นตอน -> ผล) บรรทัดละสถานะ
      • หมายเหตุ  = เว้นว่าง (ข้อมูลสถานะรวมอยู่ในช่อง 'สถานะคำขอ' แล้ว)
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "ข้อมูลการขออนุญาต"

    cols = ["บริษัท", "จังหวัด", "เลขคำขอ", "สถานะคำขอ", "หมายเหตุ"]
    ws.append(cols)

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="C2185B")
    for c in range(1, len(cols) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)

    for r in rows:
        company = r.get("company_main", "") or r.get("employer_company", "")
        province = r.get("estab_province", "")
        req_no = r.get("reqNo", "")
        # สถานะคำขอ = ทุกสถานะบนไทม์ไลน์ (วันที่ -> ขั้นตอน -> ผล) บรรทัดละสถานะ
        status_val = (r.get("status_all", "") or "").strip()
        if not status_val:
            # fallback: ใช้ข้อความดิบ (ตัดหัว 'สถานะคำขอ') เผื่อ parse ไม่ได้
            raw = (r.get("status_flow_raw", "") or "").strip()
            if raw.startswith("สถานะคำขอ"):
                raw = raw[len("สถานะคำขอ"):].strip()
            status_val = raw
        remark_val = ""
        ws.append([company, province, req_no, status_val, remark_val])

    widths = {
        "บริษัท": 40, "จังหวัด": 16, "เลขคำขอ": 18,
        "สถานะคำขอ": 90, "หมายเหตุ": 20,
    }
    for c_idx, name in enumerate(cols, start=1):
        ws.column_dimensions[get_column_letter(c_idx)].width = widths.get(name, 20)
    for r_idx in range(2, ws.max_row + 1):
        for c_idx in range(1, ws.max_column + 1):
            ws.cell(row=r_idx, column=c_idx).alignment = Alignment(
                vertical="top", wrap_text=True
            )
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(out_path)


def run_permit_report(
    cfg: dict,
    out_path: Path,
    limit: int = 0,
    log=print,
    progress=None,
    is_cancelled=None,
) -> tuple[int, Path]:
    """โหมด 'ดึงรายงานข้อมูลการขออนุญาต'
    กรองสถานะ (+ รายการคำขอ + วันที่) → เปิด detail ทีละรายการ → เก็บ
    'อนุมัติคำขอ' (สถานะคำขอ) + บริษัท/จังหวัด (สถานประกอบการ) → Excel

    Returns: (จำนวนแถวที่ดึง, path ไฟล์ที่บันทึก)
    """
    out_path = _timestamped_path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    log(f"      ไฟล์รายงาน: {out_path.name}")
    scraped = _permit_report_collect(
        cfg, limit=limit, log=log, progress=progress,
        is_cancelled=is_cancelled, save_path=out_path,
    )
    save_excel_permit_report(scraped, out_path)
    log(f"[4/4] บันทึกไฟล์ Excel: {out_path}")
    log(f"[เสร็จสิ้น] รวม {len(scraped)} แถว")
    return len(scraped), out_path


def _permit_report_collect(
    cfg: dict,
    limit: int = 0,
    log=print,
    progress=None,
    is_cancelled=None,
    save_path=None,
) -> list[dict]:
    """เปิดเบราว์เซอร์ + login + กรองสถานะ/รายการคำขอ/วันที่ → เปิด detail ทีละรายการ
    → คืน list ของแถวที่ดึงได้ (ยังไม่เขียน Excel นอกจากระบุ save_path เพื่อเซฟระหว่างทาง)
    ใช้ร่วมกันทั้งโหมดบัญชีเดียว (run_permit_report) และหลายบัญชีรวมไฟล์เดียว
    (run_permit_report_multi)
    """
    def _cancelled() -> bool:
        return bool(is_cancelled and is_cancelled())

    with sync_playwright() as pw:
        browser = _launch_chromium(pw, cfg, ["--disable-blink-features=AutomationControlled"])
        ctx = browser.new_context(
            locale="th-TH", timezone_id="Asia/Bangkok",
            viewport={"width": 1500, "height": 900},
        )
        page = ctx.new_page()
        try:
            log("[1/4] กำลังเข้าสู่ระบบ...")
            login(page, cfg)
            log(f"      ล็อกอินสำเร็จ ({page.url})")

            log("[2/4] กำลังเปิดหน้า e-Tracking และตั้งค่าฟิลเตอร์...")
            goto_tracking(page)
            if _is_logged_out(page):
                ensure_session(page, cfg, log=log)
                goto_tracking(page)

            # รวบรวมรายการคำขอที่จะกรอง (multi-select > single)
            req_types_list: list[str] = []
            for c in (cfg.get("request_types") or []):
                s = (c or "").strip()
                if s and s not in {"0", "ALL", "all"} and s not in req_types_list:
                    req_types_list.append(s)
            if not req_types_list:
                single = (cfg.get("request_type") or "").strip()
                if single and single not in {"0", "ALL", "all"}:
                    req_types_list = [single]
            if not req_types_list:
                log("      ฟิลเตอร์รายการคำขอ: ทั้งหมด (ไม่กรอง)")
            elif len(req_types_list) == 1:
                log(f"      ฟิลเตอร์รายการคำขอ: {req_types_list[0]}")
            else:
                log(f"      ฟิลเตอร์รายการคำขอ ({len(req_types_list)} รายการ): "
                    + ", ".join(req_types_list))
            req_type = req_types_list[0] if req_types_list else ""

            # สถานะที่ติ๊กจาก UI — ถ้าไม่ติ๊กเลย ดึงทุกสถานะ (กันตกหล่น)
            status_ids = cfg.get("filter_status_ids") or ["WP", "WCOSNA", "WA", "AP", "SS"]
            log(f"      ติ๊ก checkbox สถานะ: {', '.join(status_ids)}")
            date_from = (cfg.get("date_from") or "").strip()
            date_to = (cfg.get("date_to") or "").strip()
            if date_from or date_to:
                log(f"      วันที่ยื่นคำขอ: {date_from or '(ต้นสุด)'} → {date_to or '(ล่าสุด)'}")

            def _fetch_rows_for(rt_code: str) -> list[dict]:
                apply_wa_filter(
                    page, rt_code, status_ids=status_ids,
                    date_from=date_from, date_to=date_to,
                )
                out_rows = collect_all_wa_rows(page, log=log)
                if not out_rows and _is_logged_out(page):
                    if ensure_tracking_ready(page, cfg, rt_code, log=log):
                        out_rows = collect_wa_rows(page)
                return out_rows

            if len(req_types_list) <= 1:
                rows = _fetch_rows_for(req_type)
            else:
                merged: list[dict] = []
                seen: set = set()
                for _i, _code in enumerate(req_types_list, 1):
                    if _cancelled():
                        break
                    log(f"      [{_i}/{len(req_types_list)}] กำลังดึงรายการคำขอ: {_code}")
                    chunk = _fetch_rows_for(_code)
                    added = 0
                    for r in chunk:
                        k = str(r.get("reqNo") or "")
                        if k and k in seen:
                            continue
                        if k:
                            seen.add(k)
                        merged.append(r)
                        added += 1
                    log(f"          ← ได้ {len(chunk)} แถว (ใหม่ {added}, รวม {len(merged)})")
                rows = merged
            total_collected = len(rows)

            # กรองตาม whitelist สถานะ (ถ้ามี) — เก็บ detail เฉพาะสถานะที่เลือก
            whitelist = cfg.get("status_whitelist")
            if whitelist:
                before = len(rows)
                rows = [r for r in rows if row_matches_whitelist(r, whitelist)]
                if before != len(rows):
                    log(f"      กรองตาม whitelist สถานะ: {before} → {len(rows)} รายการ")

            log(f"[3/4] พบรายการที่ตรงเงื่อนไข: {len(rows)} / {total_collected} รายการ")
            if limit and limit > 0:
                rows = rows[:limit]
                log(f"      จำกัดเฉพาะ {len(rows)} รายการแรก")

            total = len(rows)
            if progress:
                progress(0, total)

            scraped: list[dict] = []
            SAVE_EVERY = max(1, int(cfg.get("save_every") or 50))
            for i, row in enumerate(rows, 1):
                if _cancelled():
                    log("[!] ยกเลิกโดยผู้ใช้ — กำลังบันทึกสิ่งที่ดึงได้แล้ว")
                    break
                log(f"  [{i}/{total}] {row.get('reqNo', '')} - {row.get('requester', '')}")
                try:
                    detail = scrape_permit_detail(page, row, cfg=cfg, log=log)
                    # เปิด detail ไม่ได้ + session หลุด → ฟื้น tracking แล้วลองใหม่ 1 ครั้ง
                    if "open_detail_failed" in (detail.get("scrape_errors") or ""):
                        if ensure_tracking_ready(page, cfg, req_type, log=log):
                            log("      ↻ ลองดึงรายการนี้อีกครั้งหลังฟื้น session...")
                            detail = scrape_permit_detail(page, row, cfg=cfg, log=log)
                    row.update(detail)
                    if detail.get("scrape_errors"):
                        log(f"      ! ดึงบางส่วนไม่ได้ (ข้าม): {detail['scrape_errors']}")
                except Exception as e:
                    log(f"      !! ผิดพลาดร้ายแรง (ข้ามรายการนี้): {e}")
                    row.setdefault("scrape_errors", str(e))

                scraped.append(row)
                if save_path and len(scraped) % SAVE_EVERY == 0:
                    try:
                        save_excel_permit_report(scraped, save_path)
                        log(f"      💾 บันทึกความคืบหน้า: {len(scraped)} รายการ")
                    except Exception as e:
                        log(f"      (เซฟระหว่างทางไม่สำเร็จ: {e})")
                if progress:
                    progress(i, total)

            return scraped
        finally:
            ctx.close(); browser.close()


def run_permit_report_multi(
    cfg: dict,
    login_excel: Path,
    out_path: Path,
    limit: int = 0,
    log=print,
    progress=None,
    is_cancelled=None,
) -> tuple[int, Path]:
    """โหมด 'ดึงรายงานข้อมูลการขออนุญาต' หลายบัญชี — วน login ทุกบัญชีใน
    UsernameLogin.xlsx แล้ว รวมทุกแถวเข้าเป็นไฟล์ Excel เดียว (ไม่แยกไฟล์ต่อบัญชี)

    Args:
        cfg: ตัวเลือกรวม (headless, request_type(s), filter_status_ids, date ฯลฯ)
             — ไม่ต้องมี username/password (จะเติมจากไฟล์ Excel ต่อบัญชี)
        login_excel: ไฟล์ UsernameLogin.xlsx (คอลัมน์ Username, Password, Type[, ระบบ])
        out_path: ไฟล์ผลลัพธ์รวม (ไฟล์เดียวทุกบัญชี)

    Returns: (จำนวนแถวรวมทุกบัญชี, path ไฟล์รวม)
    """
    accounts = _read_login_accounts(login_excel)
    if not accounts:
        raise ValueError("ไม่พบบัญชีในไฟล์ UsernameLogin.xlsx")
    out_path = _timestamped_path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    n = len(accounts)
    all_rows: list[dict] = []
    ok_accounts = 0
    log(f"[Multi] เริ่มดึงรายงานข้อมูลการขออนุญาตจาก {n} บัญชี → รวมเป็นไฟล์เดียว: {out_path.name}")
    for gi, (_ukey, acct) in enumerate(accounts.items(), start=1):
        if is_cancelled and is_cancelled():
            log("[!] ผู้ใช้ยกเลิก — หยุด"); break
        login_cfg = dict(cfg)
        login_cfg["username"] = acct["username"]
        login_cfg["password"] = acct["password"]
        login_cfg["user_type"] = acct["type"]
        # ระบบ login ต่อบัญชีจากคอลัมน์ 'ระบบ' ใน Excel (ถ้าระบุ)
        if acct.get("method"):
            login_cfg["method"] = acct["method"]
        log(f"\n===== บัญชี {gi}/{n}: {acct['username']} ({acct['type']}) =====")
        try:
            rows = _permit_report_collect(
                login_cfg, limit=limit, log=log, progress=progress,
                is_cancelled=is_cancelled, save_path=None,
            )
            all_rows.extend(rows)
            ok_accounts += 1
            log(f"      + บัญชีนี้ได้ {len(rows)} แถว (รวมสะสม {len(all_rows)} แถว)")
            # เซฟไฟล์รวมหลังจบแต่ละบัญชี (กันข้อมูลหายถ้าบัญชีถัดไปพัง)
            try:
                save_excel_permit_report(all_rows, out_path)
            except Exception as e:
                log(f"      (เซฟไฟล์รวมระหว่างทางไม่สำเร็จ: {e})")
        except Exception as e:
            log(f"   ✗ บัญชี {acct['username']} ล้มเหลว: {e} — ข้ามไปบัญชีถัดไป")
            continue
    save_excel_permit_report(all_rows, out_path)
    log(f"\n[Multi] เสร็จสิ้น — รวม {len(all_rows)} แถว จาก {ok_accounts}/{n} บัญชี → {out_path}")
    return len(all_rows), out_path


# ─────────────────────────────────────────────────────────────────
# โหมดใหม่ — ดึงข้อมูลคนต่างด้าวจากหน้า "จัดการบัญชี → ข้อมูลคนต่างด้าว"
# (ตาราง #tb_aliens บน /profile_emp42) — เก็บแค่ที่เห็นในตาราง (เร็ว)
# ─────────────────────────────────────────────────────────────────
ALIEN_SUB_TABS: list[str] = [
    "ลูกจ้างที่มีใบอนุญาตทำงาน",
    "ลูกจ้างที่มีหนังสือรับแจ้ง",
]


def goto_account_aliens(page: Page, log=print) -> bool:
    """จาก state ไหนก็ตามหลัง login → ไปหน้า /profile_emp42 และเปิด tab ข้อมูลคนต่างด้าว"""
    try:
        if "/profile_emp42" not in page.url:
            log("    [aliens] เปิดหน้า 'จัดการบัญชี'...")
            try:
                page.evaluate("if (typeof openprofile === 'function') openprofile();")
                page.wait_for_url(lambda u: "/profile_emp42" in u, timeout=15_000)
            except Exception:
                page.goto(f"{DETAIL_BASE}/profile_emp42", wait_until="domcontentloaded", timeout=25_000)
        page.wait_for_timeout(2500)
        # คลิก tab ข้อมูลคนต่างด้าว
        page.evaluate(
            r"""() => {
              const t = document.querySelector('#employee-tab');
              if (t) t.click();
            }"""
        )
        # รอจน tab content visible
        try:
            page.wait_for_selector("#tb_aliens", state="visible", timeout=15_000)
        except Exception:
            pass
        # รอ AJAX load รอบแรกจน tbody มีแถว (ไม่ใช่ placeholder 'No data')
        _wait_alien_table_loaded(page, log=log, expect_min_rows=1, timeout_ms=20_000)
        return True
    except Exception as e:
        log(f"    [aliens] เปิดหน้าคนต่างด้าวล้มเหลว: {e}")
        return False


def _wait_alien_table_loaded(page: Page, log=print, expect_min_rows: int = 1, timeout_ms: int = 20_000) -> int:
    """รอจน #tb_aliens มี data row จริง (ไม่ใช่ 'No data available') คืนจำนวนแถวสุดท้าย"""
    deadline = time.time() + timeout_ms / 1000.0
    last_count = 0
    while time.time() < deadline:
        info = page.evaluate(
            r"""() => {
              const trs = document.querySelectorAll('#tb_aliens tbody tr');
              if (!trs.length) return {n: 0, empty: true};
              // 'No data' row จะมี td colspan ใหญ่ + class dataTables_empty
              const first = trs[0];
              const emptyCell = first.querySelector('td.dataTables_empty');
              const isEmpty = !!emptyCell;
              return {n: trs.length, empty: isEmpty};
            }"""
        )
        n = int(info.get("n", 0))
        if info.get("empty"):
            return 0  # ตารางว่างจริง
        if n >= expect_min_rows:
            last_count = n
            return n
        last_count = n
        page.wait_for_timeout(500)
    log(f"    [aliens] รอตารางโหลดครบไม่ทัน (เห็น {last_count} แถว)")
    return last_count


def _click_alien_subtab(page: Page, name: str, log=print) -> bool:
    """พยายามคลิก sub-tab ตามชื่อ (text-based) — ถ้า active อยู่แล้วจะข้าม"""
    res = page.evaluate(
        r"""(name) => {
          const root = document.querySelector('#employee, [aria-labelledby="employee-tab"]') || document.body;
          // หาเฉพาะ element ที่เป็นปุ่ม/แท็บ (ไม่เอา <li> ที่ครอบกว้างเกินไป)
          const cands = Array.from(root.querySelectorAll('button, a[role="tab"], [role="tab"], .nav-link'));
          for (const el of cands) {
            const t = (el.innerText || '').replace(/\s+/g,' ').trim();
            if (t === name) {
              if (/active/.test(el.className || '')) return 'already_active';
              try { el.click(); return 'clicked'; } catch(e) { return 'click_fail:'+e.message; }
            }
          }
          // fallback: หาแบบ startsWith
          for (const el of cands) {
            const t = (el.innerText || '').replace(/\s+/g,' ').trim();
            if (t.startsWith(name)) {
              if (/active/.test(el.className || '')) return 'already_active';
              try { el.click(); return 'clicked'; } catch(e) { return 'click_fail:'+e.message; }
            }
          }
          return 'not_found';
        }""",
        name,
    )
    log(f"    [aliens] sub-tab '{name}': {res}")
    if res == "already_active":
        # ไม่ต้องรอ AJAX ใหม่ — ตารางมีอยู่แล้ว
        return True
    if res == "clicked":
        page.wait_for_timeout(1500)
        _wait_alien_table_loaded(page, log=log, expect_min_rows=1, timeout_ms=15_000)
        return True
    return False


def _expand_alien_table(page: Page, log=print) -> int:
    """สั่งให้ DataTable #tb_aliens แสดงทั้งหมด คืนจำนวน row ใน tbody หลังขยาย"""
    info = page.evaluate(
        r"""() => {
          try {
            if (typeof $ !== 'undefined' && $.fn && $.fn.DataTable && $.fn.dataTable.isDataTable('#tb_aliens')) {
              const dt = $('#tb_aliens').DataTable();
              const pi = dt.page.info();
              // ใช้ draw() (reset paging) เพื่อบังคับ reload ajax ถ้าเป็น server-side
              dt.page.len(-1).draw();
              return {ok: true, recordsTotal: pi.recordsTotal, recordsDisplay: pi.recordsDisplay, loaded: dt.rows().count()};
            }
            const sel = document.querySelector('select[name="tb_aliens_length"]');
            if (sel) {
              if (!Array.from(sel.options).some(o => o.value === '-1' || o.value === '10000')) {
                const opt = new Option('แสดงทั้งหมด', '-1'); sel.add(opt);
              }
              sel.value = sel.querySelector('option[value="-1"]') ? '-1' : '10000';
              sel.dispatchEvent(new Event('change', {bubbles:true}));
              return {ok: true, via: 'select'};
            }
            return {ok: false};
          } catch(e) { return {ok: false, err: String(e)}; }
        }"""
    )
    log(f"    [aliens] ขยายตาราง: {info}")
    # รอจน processing เสร็จ + render เสร็จ
    target = int(info.get("recordsTotal") or 0)
    deadline = time.time() + 60.0
    prev = -1
    stable = 0
    while time.time() < deadline:
        status = page.evaluate(
            r"""() => {
              const proc = document.querySelector('#tb_aliens_processing');
              const processing = proc && getComputedStyle(proc).display !== 'none';
              const trs = document.querySelectorAll('#tb_aliens tbody tr');
              let n = 0, empty = false;
              for (const tr of trs) {
                if (tr.querySelector('td.dataTables_empty')) { empty = true; continue; }
                n++;
              }
              return {n, empty, processing};
            }"""
        )
        n = int(status.get("n", 0))
        if status.get("empty") and not status.get("processing"):
            return 0
        if status.get("processing"):
            page.wait_for_timeout(700)
            continue
        if n == prev and n > 0:
            stable += 1
            if (target and n >= target) or stable >= 4:
                break
        else:
            stable = 0
        prev = n
        page.wait_for_timeout(700)
    log(f"    [aliens] ตารางแสดง {max(prev,0)} แถว (target={target})")
    return max(prev, 0)


def _collect_via_pagination(page: Page, log=print, max_pages: int = 200) -> list[dict]:
    """fallback — คลิกปุ่ม 'ถัดไป' (#tb_aliens_next) บน UI ทีละหน้าแล้วเก็บข้อมูลทุกหน้า"""
    rows: list[dict] = []
    seen = set()
    consecutive_empty = 0

    def _wait_idle() -> None:
        deadline = time.time() + 20.0
        while time.time() < deadline:
            busy = page.evaluate(
                r"""() => {
                  const p = document.querySelector('#tb_aliens_processing');
                  const processing = p && getComputedStyle(p).display !== 'none';
                  const trs = document.querySelectorAll('#tb_aliens tbody tr');
                  let hasEmpty = false, n = 0;
                  for (const tr of trs) {
                    if (tr.querySelector('td.dataTables_empty')) { hasEmpty = true; }
                    else n++;
                  }
                  return {processing, n, hasEmpty};
                }"""
            )
            if not busy.get("processing") and (busy.get("n") or busy.get("hasEmpty")):
                return
            page.wait_for_timeout(250)

    for p in range(max_pages):
        _wait_idle()
        cur = _collect_alien_rows(page)
        new_in_page = 0
        for r in cur:
            # dedupe ด้วย identity ของ "คน" จริงๆ (ไม่รวม no ที่ DataTable renumber per page)
            key = (
                (r.get("ref_id") or "").strip(),
                (r.get("alien_id") or "").strip(),
                (r.get("wp_id") or "").strip(),
                (r.get("full_name") or "").strip(),
            )
            if not any(key):
                # ทุก field ว่าง — fallback ใช้ id_cell_raw + nationality + wp_end_date
                key = (
                    (r.get("id_cell_raw") or "").strip(),
                    (r.get("nationality") or "").strip(),
                    (r.get("wp_end_date") or "").strip(),
                )
            if key in seen:
                continue
            seen.add(key)
            rows.append(r)
            new_in_page += 1
        page_info = page.evaluate(
            r"""() => {
              let cur = '', total = 0, last = false;
              try {
                if (typeof $ !== 'undefined' && $.fn.dataTable.isDataTable('#tb_aliens')) {
                  const info = $('#tb_aliens').DataTable().page.info();
                  cur = String(info.page + 1);
                  total = info.pages;
                  last = info.page + 1 >= info.pages;
                }
              } catch(e) {}
              const li = document.querySelector('#tb_aliens_next');
              const liDisabled = !li || /(^|\s)disabled(\s|$)/.test(li.className) || li.getAttribute('aria-disabled') === 'true';
              return {cur, total, last, liDisabled};
            }"""
        )
        log(f"    [aliens] page {page_info.get('cur') or (p+1)}/{page_info.get('total') or '?'}: +{new_in_page} (รวม {len(rows)})")
        if new_in_page == 0:
            consecutive_empty += 1
            if consecutive_empty >= 5:
                log(f"    [aliens] 5 หน้าติดไม่ได้ข้อมูลใหม่ (server น่าจะคืน row ซ้ำ) — หยุดที่ {len(rows)} แถว unique")
                break
        else:
            consecutive_empty = 0
        if page_info.get("last"):
            log(f"    [aliens] หน้าสุดท้าย — รวม {len(rows)} แถว")
            break
        # คลิกปุ่ม 'ถัดไป' — site นี้ใช้ custom button .paginate-right (ภายใน #tb_aliens_paginate)
        # ส่ง page_action=next + page_num=<last alien_emp_rel_id> เป็น cursor-based pagination
        next_state = page.evaluate(
            r"""() => {
              const wrap = document.querySelector('#tb_aliens_paginate');
              if (!wrap) return {ok: false, reason: 'no-paginate-wrap'};
              const btn = wrap.querySelector('.paginate-right');
              if (!btn) return {ok: false, reason: 'no-next-btn'};
              const disabled = btn.disabled || btn.hasAttribute('disabled') || /(^|\s)disabled(\s|$)/.test(btn.className);
              return {ok: true, disabled};
            }"""
        )
        if not next_state.get("ok"):
            log(f"    [aliens] ไม่พบปุ่มถัดไป ({next_state.get('reason')}) — หยุด")
            break
        if next_state.get("disabled"):
            log(f"    [aliens] ปุ่มถัดไป disabled — หยุดที่ {len(rows)} แถว")
            break
        # ใช้ Playwright real click (trusted event) — JS .click() ไม่ trigger handler บน custom button
        try:
            page.locator("#tb_aliens_paginate .paginate-right").first.click(timeout=5000)
        except Exception as e:
            log(f"    [aliens] คลิกถัดไปไม่สำเร็จ: {e} — หยุด")
            break
        # รอ ajax ตอบกลับ — เช็คจาก first row identity เปลี่ยน
        prev_first = page.evaluate(
            r"""() => {
              try {
                const dt = $('#tb_aliens').DataTable();
                const rows = dt.rows({page:'current'}).data().toArray();
                return rows.length ? JSON.stringify([rows[0].alien_id, rows[0].alien_emp_rel_id]) : '';
              } catch(e) { return ''; }
            }"""
        )
        deadline = time.time() + 15.0
        while time.time() < deadline:
            cur_first = page.evaluate(
                r"""() => {
                  try {
                    const dt = $('#tb_aliens').DataTable();
                    const rows = dt.rows({page:'current'}).data().toArray();
                    return rows.length ? JSON.stringify([rows[0].alien_id, rows[0].alien_emp_rel_id]) : '';
                  } catch(e) { return ''; }
                }"""
            )
            if cur_first and cur_first != prev_first:
                break
            page.wait_for_timeout(200)
        page.wait_for_timeout(300)
    return rows


def _collect_alien_rows(page: Page) -> list[dict]:
    """เก็บข้อมูลทุกแถวที่เห็นใน #tb_aliens
    ใช้ DataTable API ก่อน (`dt.rows({page:'current'}).nodes()`) — ถ้าไม่มีก็ fallback DOM
    """
    return page.evaluate(
        r"""() => {
          const out = [];
          let trs = [];
          try {
            if (typeof $ !== 'undefined' && $.fn && $.fn.dataTable && $.fn.dataTable.isDataTable('#tb_aliens')) {
              const dt = $('#tb_aliens').DataTable();
              const nodes = dt.rows({page: 'current'}).nodes();
              for (let i = 0; i < nodes.length; i++) trs.push(nodes[i]);
            }
          } catch(e) {}
          if (!trs.length) {
            trs = Array.from(document.querySelectorAll('#tb_aliens tbody tr'));
          }
          trs.forEach(tr => {
            const tds = tr.querySelectorAll('td');
            if (!tds.length) return;
            if (tr.querySelector('td.dataTables_empty')) return;
            const cellText = (i) => tds[i] ? (tds[i].innerText||'').replace(/\s+/g,' ').trim() : '';
            const idCellRaw = tds[1] ? (tds[1].innerText||'').trim() : '';
            const refIdM = idCellRaw.match(/RA[A-Z0-9]+/i);
            const alienIdM = idCellRaw.match(/เลขประจำตัวคนต่างด้าว\s*[:：]\s*(\S+)/);
            const wpIdM = idCellRaw.match(/เลขที่ใบอนุญาตทำงาน\s*[:：]\s*(\S+)/);
            const nameRaw = tds[2] ? (tds[2].innerText||'').trim() : '';
            const nameLines = nameRaw.split(/\n+/).map(s => s.trim()).filter(Boolean);
            const fullName = nameLines[0] || '';
            const permitStatus = nameLines.slice(1).join(' ');
            out.push({
              no: cellText(0),
              ref_id: refIdM ? refIdM[0] : '',
              alien_id: alienIdM ? alienIdM[1] : '',
              wp_id: wpIdM ? wpIdM[1] : '',
              id_cell_raw: idCellRaw.replace(/\s+/g,' ').trim(),
              full_name: fullName,
              permit_status: permitStatus,
              nationality: cellText(3),
              wp_end_date: cellText(4),
              workplace: cellText(5),
              entry_exit_status: cellText(6),
            });
          });
          return out;
        }"""
    )


def _fetch_aliens_via_dt_api(page: Page, log=print, page_len: int = 10) -> list[dict]:
    """Fallback ระดับลึก — ใช้ DataTable API วน page(n).draw() ทีละหน้า
    รอ event 'xhr.dt' (server-side) ระหว่าง draw — เชื่อถือได้กว่าคลิก <a>
    ใช้ page_len=10 เป็น default เพราะ endpoint ของบาง sub-tab รองรับเฉพาะค่า default
    """
    info = page.evaluate(
        r"""async (pl) => {
          try {
            const dt = $('#tb_aliens').DataTable();
            // reset เป็น pageLength ที่กำหนด แล้วรอ draw แรกให้เสร็จ
            await new Promise(res => {
              const done = () => { dt.off('draw.dt', done); res(); };
              dt.on('draw.dt', done);
              dt.page.len(pl).page(0).draw();
              // safety timeout
              setTimeout(() => { dt.off('draw.dt', done); res(); }, 15000);
            });
            const pi = dt.page.info();
            return {ok: true, total: pi.recordsTotal, pages: pi.pages, length: pi.length};
          } catch(e) { return {ok: false, err: String(e)}; }
        }""",
        page_len,
    )
    if not info.get("ok"):
        log(f"    [aliens] dt-api init error: {info}")
        return []
    total = int(info.get("total") or 0)
    pages = int(info.get("pages") or 0)
    log(f"    [aliens] dt-api: total={total} pages={pages} length={info.get('length')}")
    if not total or not pages:
        return []

    all_rows: list[dict] = []
    seen_keys: set = set()
    consecutive_empty = 0
    for pidx in range(pages):
        page_data = page.evaluate(
            r"""async (n) => {
              const dt = $('#tb_aliens').DataTable();
              if (dt.page() !== n) {
                await new Promise(res => {
                  let xhrSeen = false, drawSeen = false, done = false;
                  const tryFinish = () => {
                    if (done) return;
                    if (xhrSeen && drawSeen) {
                      done = true;
                      dt.off('draw.dt', onDraw); dt.off('xhr.dt', onXhr);
                      res();
                    }
                  };
                  const onXhr = () => { xhrSeen = true; tryFinish(); };
                  const onDraw = () => { drawSeen = true; tryFinish(); };
                  dt.on('xhr.dt', onXhr);
                  dt.on('draw.dt', onDraw);
                  dt.page(n).draw('page');
                  // safety: ถ้า 30s ยังไม่ครบ event ทั้ง 2 → ปล่อยผ่านเก็บที่ได้
                  setTimeout(() => { if (!done) { done = true;
                    dt.off('draw.dt', onDraw); dt.off('xhr.dt', onXhr); res(); } }, 30000);
                });
              }
              const arr = dt.rows({page:'current'}).data().toArray();
              const pi = dt.page.info();
              return {arr, cur: pi.page, len: pi.length};
            }""",
            pidx,
        ) or {}
        rows_in = page_data.get("arr") or []
        cur = page_data.get("cur")
        added = 0
        for r in rows_in:
            if not isinstance(r, dict):
                continue
            # ใช้ field ที่น่าจะ unique ที่สุดเป็น dedupe key — ถ้าทุก field null → fallback page+index
            key = (
                r.get("alien_id") or "",
                r.get("alen_id_card") or r.get("alien_id_card") or "",
                r.get("alien_permit_no") or "",
                r.get("alien_emp_rel_id") or "",
            )
            if not any(key):
                # ทุก field ว่าง → ไม่ dedupe เลย ใช้ตำแหน่งใน sequence
                key = (pidx, len(all_rows))
            if key in seen_keys:
                continue
            seen_keys.add(key)
            all_rows.append(r)
            added += 1
        log(f"    [aliens] dt-api page {pidx+1}/{pages} (dt.page={cur}): +{added} (รวม {len(all_rows)}/{total})")
        if added == 0:
            consecutive_empty += 1
            if consecutive_empty >= 5:
                log(f"    [aliens] dt-api 5 หน้าติดไม่ได้ข้อมูลใหม่ — หยุด")
                break
        else:
            consecutive_empty = 0
        # หน่วงเล็กน้อยกัน server throttle
        page.wait_for_timeout(400)
    return _parse_alien_raw_rows(all_rows)


def _fetch_aliens_via_api_direct(page: Page, log=print, page_len: int = 10) -> list[dict]:
    """Bypass DataTable: ดึง ajax URL/params จาก DT แล้ว POST ตรงผ่าน playwright request
    เชื่อถือได้สุด — ไม่พึ่ง DT state ที่อาจเสียจาก error ก่อนหน้า
    """
    meta = page.evaluate(
        r"""() => {
          try {
            const dt = $('#tb_aliens').DataTable();
            const url = dt.ajax.url();
            const params = dt.ajax.params();
            // ตัด field ที่ใหญ่/ไม่จำเป็น
            return {ok: true, url, params};
          } catch(e) { return {ok: false, err: String(e)}; }
        }"""
    )
    if not meta.get("ok"):
        log(f"    [aliens] api-direct meta error: {meta}")
        return []
    raw_url = meta.get("url") or ""
    raw_params = meta.get("params") or {}
    # ทำเป็น absolute URL
    if raw_url.startswith("/"):
        # ใช้ origin จาก page.url
        from urllib.parse import urlparse
        u = urlparse(page.url)
        full_url = f"{u.scheme}://{u.netloc}{raw_url}"
    else:
        full_url = raw_url
    # แปลง params ของ DataTables (มี columns[…] / order[…] / search[value]) ให้เป็น flat form
    def _flatten(prefix: str, val) -> list[tuple[str, str]]:
        out: list[tuple[str, str]] = []
        if isinstance(val, dict):
            for k, v in val.items():
                key = f"{prefix}[{k}]" if prefix else str(k)
                out.extend(_flatten(key, v))
        elif isinstance(val, list):
            for i, v in enumerate(val):
                key = f"{prefix}[{i}]"
                out.extend(_flatten(key, v))
        else:
            out.append((prefix, "" if val is None else str(val)))
        return out
    base_form = _flatten("", raw_params)
    # ดึง recordsTotal ก่อนด้วย start=0
    def _post(start: int, length: int, draw: int) -> dict:
        form: list[tuple[str, str]] = []
        for k, v in base_form:
            if k == "start":
                form.append((k, str(start)))
            elif k == "length":
                form.append((k, str(length)))
            elif k == "draw":
                form.append((k, str(draw)))
            else:
                form.append((k, v))
        # ถ้า base ไม่มี start/length/draw → เติม
        keys_present = {k for k, _ in form}
        for missing_k, missing_v in [("start", str(start)), ("length", str(length)), ("draw", str(draw))]:
            if missing_k not in keys_present:
                form.append((missing_k, missing_v))
        try:
            resp = page.context.request.post(
                full_url,
                form=dict(form),  # playwright รวม dup key เป็นค่าหลัง — DT ส่ง key หลายชั้นไม่ซ้ำกัน OK
                headers={
                    "X-Requested-With": "XMLHttpRequest",
                    "Accept": "application/json, text/javascript, */*; q=0.01",
                    "Referer": page.url,
                },
            )
            if not resp.ok:
                return {"ok": False, "status": resp.status, "err": f"HTTP {resp.status}"}
            try:
                j = resp.json()
            except Exception:
                txt = resp.text()[:200]
                return {"ok": False, "err": f"non-json: {txt}"}
            return {"ok": True, "total": int(j.get("recordsTotal") or 0), "data": j.get("data") or []}
        except Exception as e:
            return {"ok": False, "err": str(e)}

    first = _post(0, page_len, 1)
    if not first.get("ok"):
        log(f"    [aliens] api-direct first error: {first}")
        return []
    total = int(first.get("total") or 0)
    data: list = list(first.get("data") or [])
    log(f"    [aliens] api-direct: total={total} page_len={page_len} — เริ่มดึง")
    log(f"    [aliens] api-direct chunk 1: +{len(data)} (รวม {len(data)}/{total})")
    if not total:
        return _parse_alien_raw_rows(data)

    # dedupe ด้วย identity จริงของแต่ละแถว — กัน server bug ที่คืน row ชุดเดิมซ้ำ
    def _row_key(d: dict) -> tuple:
        return (
            str(d.get("alien_id") or ""),
            str(d.get("alien_emp_rel_id") or ""),
            str(d.get("alien_permit_no") or ""),
            str(d.get("alien_id_card") or d.get("alen_id_card") or ""),
        )
    seen: set = set()
    unique_data: list = []
    for d in data:
        k = _row_key(d)
        if any(k) and k in seen:
            continue
        seen.add(k); unique_data.append(d)
    data = unique_data

    draw = 2
    start = page_len  # ขอ chunk ถัดไปด้วย start=page_len เสมอ ไม่อิง len(data) (กันกรณี dedupe ตัดทิ้ง)
    consecutive_fail = 0
    consecutive_dup = 0  # นับ chunk ติดที่ได้แต่ row ซ้ำเดิม
    while start < total:
        r = _post(start, page_len, draw)
        draw += 1
        got = r.get("data") or [] if r.get("ok") else []
        if not r.get("ok"):
            log(f"    [aliens] api-direct error @start={start}: {r}")
            consecutive_fail += 1
            if consecutive_fail >= 3:
                break
            start += page_len
            continue
        if not got:
            log(f"    [aliens] api-direct @start={start} ว่าง — หยุด")
            break
        # นับเฉพาะ row ใหม่จริงๆ
        new_in_chunk = 0
        for d in got:
            k = _row_key(d)
            if any(k) and k in seen:
                continue
            seen.add(k); data.append(d); new_in_chunk += 1
        consecutive_fail = 0
        if new_in_chunk == 0:
            consecutive_dup += 1
            if consecutive_dup >= 2:
                log(f"    [aliens] api-direct @start={start}: server คืน row ซ้ำเดิมทั้งหมด → server bug — หยุดที่ {len(data)} แถว unique")
                break
        else:
            consecutive_dup = 0
            if (draw % 20) == 0 or len(data) >= total:
                log(f"    [aliens] api-direct @start={start}: +{new_in_chunk} ใหม่ (รวม {len(data)}/{total})")
        start += page_len
    log(f"    [aliens] api-direct รวม {len(data)} แถว unique (server บอก total={total})")
    return _parse_alien_raw_rows(data)


def _fetch_aliens_via_ajax(page: Page, log=print, length: int = 10000) -> list[dict]:
    """เรียก ajax endpoint ของ DataTable ตรงๆ ดึงทุกแถวมา
    ถ้า server ตัด response (เช่นรับสูงสุด ~500) จะ chunk ขอเป็นช่วงๆ จนครบ
    คืน list[dict] schema เดียวกับ _collect_alien_rows
    """
    def _call(start: int, ln: int) -> dict:
        return page.evaluate(
            r"""async ({start, ln}) => {
              try {
                const dt = $('#tb_aliens').DataTable();
                const url = dt.ajax.url();
                const params = $.extend(true, {}, dt.ajax.params(), {length: ln, start: start});
                const resp = await $.ajax({url, method: 'POST', data: params, dataType: 'json'});
                return {ok: true, total: resp.recordsTotal, filtered: resp.recordsFiltered, data: resp.data || []};
              } catch(e) {
                return {ok: false, err: String(e && e.message || e), status: e && e.status};
              }
            }""",
            {"start": start, "ln": ln},
        )

    # ครั้งแรก — ดูว่า server ตอบกลับมากี่แถว และ recordsTotal เท่าไร
    first = _call(0, length)
    if not first.get("ok"):
        log(f"    [aliens] ajax error: {first}")
        return []
    total = int(first.get("total") or 0)
    data: list = list(first.get("data") or [])
    log(f"    [aliens] ajax chunk 1: {len(data)}/{total} แถว")

    # ถ้ายังไม่ครบ → chunk เพิ่ม (server อาจจำกัด ~500/ครั้ง)
    if total and len(data) < total:
        chunk_size = max(len(data), 500)  # ใช้ขนาดที่ server ยอมรับ
        start = len(data)
        # ขนาด chunk สำรองเมื่อได้ 0 — บาง endpoint จำกัด length เล็กกว่ามาก
        fallback_sizes = [500, 250, 100, 50, 10]
        while start < total:
            r = _call(start, chunk_size)
            got = r.get("data") or [] if r.get("ok") else []
            if not r.get("ok"):
                log(f"    [aliens] ajax chunk error @start={start}: {r}")
            # ถ้าได้ 0 → ลอง chunk เล็กลงเรื่อยๆ ก่อนยอมแพ้
            if not got:
                recovered = False
                for fs in fallback_sizes:
                    if fs >= chunk_size:
                        continue
                    r2 = _call(start, fs)
                    g2 = r2.get("data") or [] if r2.get("ok") else []
                    if g2:
                        log(f"    [aliens] ajax @start={start} ลอง length={fs} → +{len(g2)}")
                        got = g2
                        chunk_size = fs
                        recovered = True
                        break
                if not recovered:
                    log(f"    [aliens] ajax chunk @start={start} ว่าง — หยุด")
                    break
            data.extend(got)
            log(f"    [aliens] ajax chunk @start={start}: +{len(got)} (รวม {len(data)}/{total})")
            start += len(got)
            if len(got) < chunk_size and start < total:
                # server คืนน้อยกว่าขอ — อาจถึงจุดสิ้นสุด หรือเปลี่ยน chunk เล็กลง
                chunk_size = max(len(got), 100)

    log(f"    [aliens] ajax คืน {len(data)}/{total} แถว")

    return _parse_alien_raw_rows(data)


def _parse_alien_raw_rows(data: list) -> list[dict]:
    """แปลง raw rows (จาก ajax หรือ DataTable API) → schema มาตรฐานของ scraper"""
    import re as _re

    def _strip_html(s: str) -> str:
        if not s:
            return ""
        s = _re.sub(r"<br\s*/?>", "\n", s, flags=_re.I)
        s = _re.sub(r"<[^>]+>", " ", s)
        s = (
            s.replace("&nbsp;", " ").replace("&amp;", "&")
            .replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
        )
        lines = [_re.sub(r"[ \t]+", " ", ln).strip() for ln in s.split("\n")]
        return "\n".join([ln for ln in lines if ln])

    def _g(d: dict, *keys) -> str:
        for k in keys:
            v = d.get(k)
            if v is None:
                continue
            s = str(v).strip()
            if s and s.lower() != "null":
                return s
        return ""

    rows: list[dict] = []
    for i, r in enumerate(data, start=1):
        if not isinstance(r, dict):
            # fallback: array of HTML strings
            cells = [_strip_html(str(c or "")) for c in (list(r)[:7] + [""] * 7)][:7]
            no, id_cell, name_cell, nat, wp_end, workplace, entry_exit = cells
            m_ref = _re.search(r"RA[A-Z0-9]+", id_cell, _re.I)
            m_alien = _re.search(r"เลขประจำตัวคนต่างด้าว\s*[:：]\s*(\S+)", id_cell)
            m_wp = _re.search(r"เลขที่ใบอนุญาตทำงาน\s*[:：]\s*(\S+)", id_cell)
            rows.append({
                "no": no or str(i),
                "ref_id": m_ref.group(0) if m_ref else "",
                "alien_id": m_alien.group(1) if m_alien else "",
                "wp_id": m_wp.group(1) if m_wp else "",
                "id_cell_raw": _re.sub(r"\s+", " ", id_cell).strip(),
                "full_name": name_cell.split("\n")[0] if name_cell else "",
                "permit_status": " ".join(name_cell.split("\n")[1:]) if name_cell else "",
                "nationality": nat,
                "wp_end_date": wp_end,
                "workplace": workplace,
                "entry_exit_status": entry_exit,
            })
            continue

        # ประกอบชื่อภาษาไทย (prefix + first + middle + last)
        full_name_th = _g(r, "full_name_th")
        if not full_name_th:
            name_th_parts = [
                _g(r, "prefix_name_th"),
                _g(r, "firstname_th"),
                _g(r, "middlename_th"),
                _g(r, "lastname_th"),
            ]
            full_name_th = " ".join([p for p in name_th_parts if p]).strip()
        full_name_en = _g(r, "full_name_en")
        if not full_name_en:
            name_en_parts = [
                _g(r, "prefix_name_en"),
                _g(r, "firstname_en"),
                _g(r, "middlename_en"),
                _g(r, "lastname_en"),
            ]
            full_name_en = " ".join([p for p in name_en_parts if p]).strip()
        full_name = full_name_th or full_name_en
        # สถานที่ทำงาน
        workplace = _g(r, "address_th", "addr_name_th") or _g(r, "address_en", "addr_name_en")
        # สัญชาติ
        nationality = _g(r, "nationality_th") or _g(r, "nationality_en")
        # วันที่สิ้นสุดใบอนุญาต (ตัด 00:00:00)
        wp_end = _g(r, "stay_permis_expire_dt")
        wp_end = _re.sub(r"\s+00:00:00$", "", wp_end)
        # เลขที่ใบอนุญาตทำงาน
        wp_id = _g(r, "alien_permit_no")
        # เลขประจำตัวคนต่างด้าว (13 หลัก)
        alien_personal_no = _g(r, "alen_id_card", "alien_id_card")
        # รหัสอ้างอิง (RA…)
        ref_id = _g(r, "alien_id")
        # สถานะใบอนุญาต / สถานะแจ้งเข้า-ออก
        permit_status = _g(r, "wp_status_name_th", "wp_status_name_en")
        entry_exit_status = _g(r, "working_status_th", "working_status_en")
        rows.append({
            "no": str(i),
            "ref_id": ref_id,
            "alien_id": alien_personal_no,
            "wp_id": wp_id,
            "id_cell_raw": f"{ref_id} | {alien_personal_no} | {wp_id}".strip(" |"),
            "full_name": full_name,
            "full_name_en": full_name_en,
            "permit_status": permit_status,
            "nationality": nationality,
            "wp_end_date": wp_end,
            "workplace": workplace,
            "entry_exit_status": entry_exit_status,
            "form_type_id": _g(r, "form_type_id"),
            "alien_emp_rel_id": _g(r, "alien_emp_rel_id"),
        })
    return rows


def scrape_aliens(page: Page, sub_tabs: list[str], log=print) -> dict[str, list[dict]]:
    """เก็บข้อมูลจากตาราง #tb_aliens สำหรับ sub-tab ที่ระบุ
    คืน {sub_tab_name: [rows]}
    """
    result: dict[str, list[dict]] = {}
    if not goto_account_aliens(page, log=log):
        return result
    for st in sub_tabs:
        log(f"  [sub-tab] {st}")
        clicked = _click_alien_subtab(page, st, log=log)
        if not clicked:
            log(f"    [aliens] ข้าม sub-tab '{st}' (ไม่พบ)")
            result[st] = []
            continue
        # อ่าน recordsTotal เพื่อรู้จำนวนเป้าหมาย
        target = page.evaluate(
            r"""() => {
              try {
                if (typeof $ !== 'undefined' && $.fn.dataTable.isDataTable('#tb_aliens')) {
                  return $('#tb_aliens').DataTable().page.info().recordsTotal;
                }
              } catch(e) {}
              return 0;
            }"""
        ) or 0
        target = int(target)
        log(f"    [aliens] เป้าหมาย {target} แถว")
        # วิธีหลัก: ajax ตรง (server-side DataTable) — เร็วและครบสำหรับ sub-tab ปกติ
        rows = _fetch_aliens_via_ajax(page, log=log, length=max(target or 0, 10000))
        # ถ้า ajax ไม่ครบ → fallback ใช้ UI click 'ถัดไป' (เลียนแบบคน)
        if not rows or (target and len(rows) < target):
            log(f"    [aliens] ajax ได้ {len(rows)}/{target} → ลอง UI pagination (คลิกถัดไป)")
            page.evaluate(
                r"""() => {
                  try {
                    if (typeof $ !== 'undefined' && $.fn.dataTable.isDataTable('#tb_aliens')) {
                      $('#tb_aliens').DataTable().page.len(10).page(0).draw();
                    }
                  } catch(e) {}
                }"""
            )
            page.wait_for_timeout(1000)
            rows_ui = _collect_via_pagination(page, log=log, max_pages=max((target or 0) // 10 + 10, 200))
            if len(rows_ui) > len(rows):
                rows = rows_ui
        log(f"    [aliens] เก็บได้ {len(rows)} แถว")
        result[st] = rows
    return result


def save_excel_aliens(data: dict[str, list[dict]], out_path: Path) -> None:
    """บันทึกแต่ละ sub-tab เป็น sheet แยกใน workbook เดียว"""
    wb = Workbook()
    wb.remove(wb.active)
    cols = [
        ("no", "ลำดับ", 8),
        ("ref_id", "รหัสอ้างอิง (RA)", 24),
        ("alien_id", "เลขประจำตัวคนต่างด้าว", 22),
        ("wp_id", "เลขที่ใบอนุญาตทำงาน", 22),
        ("full_name", "ชื่อลูกจ้าง", 30),
        ("full_name_en", "ชื่อลูกจ้าง (EN)", 30),
        ("permit_status", "สถานะใบอนุญาต", 20),
        ("nationality", "สัญชาติ", 18),
        ("wp_end_date", "วันที่สิ้นสุดใบอนุญาต", 20),
        ("workplace", "สถานที่ทำงาน", 35),
        ("entry_exit_status", "สถานะแจ้งเข้า/แจ้งออก", 22),
        ("form_type_id", "รหัสประเภทคำขอ", 18),
    ]
    bold = Font(bold=True)
    fill = PatternFill("solid", fgColor="FFE5E5")
    for sheet_name, rows in data.items():
        # sheet name max 31 chars, ห้าม / \ ? * [ ]
        safe = re.sub(r"[\\/?*\[\]:]", "_", sheet_name)[:31] or "Sheet"
        ws = wb.create_sheet(safe)
        ws.append([c[1] for c in cols])
        for c in range(1, len(cols) + 1):
            cell = ws.cell(row=1, column=c)
            cell.font = bold
            cell.fill = fill
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        for r in rows:
            ws.append([r.get(c[0], "") for c in cols])
        for i, (_, _, w) in enumerate(cols, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        for r_idx in range(2, ws.max_row + 1):
            for c_idx in range(1, ws.max_column + 1):
                ws.cell(row=r_idx, column=c_idx).alignment = Alignment(
                    vertical="top", wrap_text=True
                )
        ws.freeze_panes = "A2"
        if ws.max_row >= 1:
            ws.auto_filter.ref = ws.dimensions
    if not wb.sheetnames:
        wb.create_sheet("Empty")
    wb.save(out_path)


REGISTER_URL = "https://eworkpermit.doe.go.th/Login/Register"
REGISTER_ALIEN_URL = "https://eworkpermit.doe.go.th/Register/Alien"


def _parse_row_range(spec: str | None, total: int) -> list[int]:
    """แปลง '1-10,15,20-25' เป็น list ของ index 1-based"""
    if not spec:
        return list(range(1, total + 1))
    out: list[int] = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            a, b = int(a.strip()), int(b.strip())
            out.extend(range(min(a, b), max(a, b) + 1))
        else:
            out.append(int(part))
    return [i for i in out if 1 <= i <= total]


def _safe_filename(s: Any) -> str:
    """แปลงข้อความเป็นชื่อไฟล์ที่ปลอดภัย (เก็บอักษรไทย/อังกฤษ/ตัวเลข, ตัวอื่นเป็น _)"""
    s = str(s or "").strip()
    s = re.sub(r"[\\/:*?\"<>|\s]+", "_", s)
    s = s.strip("._")
    return s[:50] or "x"


def _make_basename(rec: dict[str, Any]) -> str:
    """สร้าง basename สำหรับ screenshot จาก No. + TaxID + Name"""
    no = _safe_filename(rec.get("No.") or rec.get("No") or "")
    tax = _safe_filename(rec.get("TaxID") or "")
    name = _safe_filename(rec.get("Name") or "")
    parts = [p for p in (no, tax, name) if p and p != "x"]
    return "_".join(parts) or "row"


def _read_register_excel(path: Path) -> list[dict[str, Any]]:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    hdr = [c.value for c in ws[1]]
    rows: list[dict[str, Any]] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(v not in (None, "") for v in row):
            continue
        rec = {hdr[i]: row[i] for i in range(len(hdr)) if hdr[i]}
        rows.append(rec)
    return rows


def _capture_register_alert(page: Page) -> str:
    """Capture alert/confirmation modal text (SweetAlert or Bootstrap).
    Exclude the search_alien_modal (form modal) to avoid capturing form content."""
    try:
        return page.evaluate(
            r"""() => {
              // SweetAlert2 (priority)
              for (const el of document.querySelectorAll('.swal2-popup, .swal2-container')) {
                if (el.offsetParent === null) continue;
                const t = el.querySelector('.swal2-title');
                const c = el.querySelector('.swal2-html-container, .swal2-content');
                return ((t?t.textContent.trim():'') + ' | ' + (c?c.textContent.trim():'')).slice(0, 600);
              }
              // Bootstrap modal (but exclude search_alien_modal)
              for (const el of document.querySelectorAll('.modal.show, .modal[style*="display: block"]')) {
                if (el.id === 'search_alien_modal') continue;  // Skip the search form modal
                const t = el.querySelector('.modal-title');
                const b = el.querySelector('.modal-body');
                return ((t?t.textContent.trim():'') + ' | ' + (b?b.textContent.trim().slice(0,500):''));
              }
              return '';
            }"""
        ) or ""
    except Exception:
        return ""


def _close_register_alert(page: Page) -> None:
    try:
        page.evaluate(
            r"""() => {
              document.querySelectorAll('.swal2-popup .swal2-confirm, .swal2-container .swal2-confirm').forEach(b => b.click());
              document.querySelectorAll('.modal.show .btn-primary, .modal.show .close, .modal[style*="display: block"] .btn-primary').forEach(b => b.click());
            }"""
        )
        page.wait_for_timeout(700)
    except Exception:
        pass


def _register_init_form(page: Page) -> None:
    page.goto(REGISTER_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(2000)
    page.locator("#User_typeRegister").select_option(value="1")
    page.wait_for_timeout(500)
    page.locator("input#alienOption1").check(force=True)
    page.wait_for_timeout(300)


def _register_attempt_id(page: Page, id_type_value: str, id_value: str, log=print) -> str:
    """พยายามลงทะเบียนด้วยเลข id ที่กำหนด คืน '' ถ้าผ่าน (ไปหน้า /Register/Alien) มิฉะนั้นคืนข้อความ alert"""
    page.locator("#typeAlienUser").select_option(value=id_type_value)
    page.wait_for_timeout(300)
    page.locator("#inputAlien1").fill(id_value)
    page.wait_for_timeout(200)
    try:
        page.locator("#validate_register").click()
    except Exception as e:
        return f"กดปุ่มลงทะเบียนไม่ได้: {e}"
    # รอเปลี่ยนหน้า หรือ alert
    for _ in range(20):
        page.wait_for_timeout(300)
        if "/Register/Alien" in page.url:
            return ""
        a = _capture_register_alert(page)
        if a:
            return a
    if "/Register/Alien" in page.url:
        return ""
    return _capture_register_alert(page) or "ไม่ทราบสาเหตุ (ไม่เปลี่ยนหน้าและไม่มี alert)"


def _register_one_alien(
    page: Page,
    rec: dict[str, Any],
    screenshot_dir: Path,
    log=print,
) -> dict[str, Any]:
    no = rec.get("No.") or rec.get("No") or "?"
    name = str(rec.get("Name") or "").strip()
    tax = str(rec.get("TaxID") or "").strip()
    passport = str(rec.get("PassportNo.") or rec.get("PassportNo") or "").strip()
    pwd = str(rec.get("Password") or "").strip()
    email = str(rec.get("Email") or "").strip()
    base = _make_basename(rec)
    out: dict[str, Any] = {
        "No.": no, "Name": name, "TaxID": tax, "PassportNo.": passport, "Email": email,
        "Status": "", "AttemptedWith": "", "AlertText": "",
        "ScreenshotAlert": "", "ScreenshotSuccess": "", "Error": "",
        "Basename": base,
    }

    log(f"  → ลงทะเบียน No.{no} {name} (TaxID={tax}, Passport={passport})")

    # 1.1 - 1.5 ใช้ TaxID ก่อน
    try:
        _register_init_form(page)
    except Exception as e:
        out["Status"] = "ERROR"; out["Error"] = f"init form failed: {e}"
        return out

    alert_tax = _register_attempt_id(page, "1", tax, log=log)
    used_id_type = "TaxID"
    if alert_tax:
        ss1 = screenshot_dir / f"{base}_alert_taxid.png"
        try: page.screenshot(path=str(ss1), full_page=True)
        except Exception: pass
        log(f"     TaxID ไม่ผ่าน: {alert_tax[:120]}")
        _close_register_alert(page)
        # ลอง Passport
        try:
            _register_init_form(page)
        except Exception as e:
            out["Status"] = "FAIL"; out["AlertText"] = f"TaxID: {alert_tax}"
            out["ScreenshotAlert"] = ss1.name; out["Error"] = f"reinit failed: {e}"
            return out
        alert_pp = _register_attempt_id(page, "3", passport, log=log)
        if alert_pp:
            ss2 = screenshot_dir / f"{base}_alert_passport.png"
            try: page.screenshot(path=str(ss2), full_page=True)
            except Exception: pass
            log(f"     Passport ไม่ผ่าน: {alert_pp[:120]}")
            out["Status"] = "FAIL"
            out["AttemptedWith"] = "TaxID,Passport"
            out["AlertText"] = f"TaxID: {alert_tax} || Passport: {alert_pp}"
            # ใช้ไฟล์ล่าสุด (passport) เป็น link หลัก เพื่อเลี่ยงปัญหาหลายลิงก์ในเซลล์เดียว
            out["ScreenshotAlert"] = ss2.name
            return out
        used_id_type = "Passport"
    out["AttemptedWith"] = used_id_type

    # 1.6 ยอมรับเงื่อนไข
    try:
        page.locator("#checkbox_em").wait_for(state="attached", timeout=15000)
        page.locator("#checkbox_em").evaluate("e => e.click()")
        page.wait_for_timeout(500)
        page.locator("#gonext_0").wait_for(state="visible", timeout=10000)
        # รอจน enable
        for _ in range(20):
            disabled = page.evaluate("() => { const b=document.getElementById('gonext_0'); return b? b.disabled : true; }")
            if not disabled:
                break
            page.wait_for_timeout(200)
        page.locator("#gonext_0").click()
        page.wait_for_timeout(2500)

        # 1.7 ฟอร์มข้อมูลผู้ใช้งาน
        page.locator("#educationLevel2").wait_for(state="visible", timeout=15000)
        page.locator("#educationLevel2").select_option(value="8")
        page.wait_for_timeout(300)
        page.locator("#workExperien_Alien2").fill("2 ปี")
        page.wait_for_timeout(200)
        page.locator("#validateFormAlien").click()
        page.wait_for_timeout(2500)
        a = _capture_register_alert(page)
        if a:
            ss = screenshot_dir / f"{base}_alert_form.png"
            try: page.screenshot(path=str(ss), full_page=True)
            except Exception: pass
            out["Status"] = "FAIL"; out["AlertText"] = f"FORM: {a}"; out["ScreenshotAlert"] = ss.name
            return out

        # 1.8 สรุปข้อมูล
        page.locator("#summaryInfoAlien").wait_for(state="visible", timeout=15000)
        page.locator("#summaryInfoAlien").click()
        page.wait_for_timeout(2500)

        # 1.9a รหัสผ่าน
        page.locator("#setPassword").wait_for(state="visible", timeout=15000)
        page.locator("#setPassword").fill(pwd)
        page.locator("#confirmPassword").fill(pwd)
        page.wait_for_timeout(300)
        page.locator("#validateSetAlienPassword").click()
        page.wait_for_timeout(2500)
        a = _capture_register_alert(page)
        if a:
            ss = screenshot_dir / f"{base}_alert_password.png"
            try: page.screenshot(path=str(ss), full_page=True)
            except Exception: pass
            out["Status"] = "FAIL"; out["AlertText"] = f"PASSWORD: {a}"; out["ScreenshotAlert"] = ss.name
            return out

        # 1.9b อีเมล + ข้าม OTP
        page.locator("#otp_skip").wait_for(state="attached", timeout=15000)
        page.locator("#otp_skip").evaluate("e => e.click()")
        page.wait_for_timeout(300)
        page.locator("#user_email").fill(email)
        page.wait_for_timeout(300)
        page.locator("#validateEmailOTPAlien").click()
        page.wait_for_timeout(5000)
        a = _capture_register_alert(page)
        if a and "สำเร็จ" not in a:
            ss = screenshot_dir / f"{base}_alert_email.png"
            try: page.screenshot(path=str(ss), full_page=True)
            except Exception: pass
            out["Status"] = "FAIL"; out["AlertText"] = f"EMAIL: {a}"; out["ScreenshotAlert"] = ss.name
            return out

        # capture หน้าสำเร็จ
        ss_ok = screenshot_dir / f"{base}_success.png"
        try: page.screenshot(path=str(ss_ok), full_page=True)
        except Exception: pass
        out["Status"] = "SUCCESS"
        out["ScreenshotSuccess"] = ss_ok.name
        log(f"     ✅ สำเร็จ ({used_id_type})")
    except Exception as e:
        ss_err = screenshot_dir / f"{base}_error.png"
        try: page.screenshot(path=str(ss_err), full_page=True)
        except Exception: pass
        out["Status"] = "ERROR"; out["Error"] = str(e)[:400]; out["ScreenshotAlert"] = ss_err.name
        log(f"     ❌ ERROR: {e}")
    return out


def _switch_lang_th(page: Page) -> None:
    try:
        page.locator("#sltLang").first.select_option(value="th")
        page.wait_for_timeout(2000)
    except Exception:
        pass


# label ที่จะดึงจากหน้า ProfileDigitalWorkPermit (ไทย)
PERMIT_LABELS = [
    "เลขประจำตัวคนต่างด้าว",
    "ใบอนุญาตทำงานเลขที่",
    "ชื่อ(Eng)",
    "วันเดือนปีเกิด",
    "สัญชาติ",
    "วันที่อนุญาตทำงาน",
    "วันที่สิ้นสุดใบอนุญาตทำงาน",
    "ชื่อสถานประกอบการ",
    "สถานที่ทำงาน",
    "ประเภทกิจการ",
    "ประเภทงาน",
    "ลักษณะงาน",
    "ชื่อนายทะเบียน",
    "ลายมือชื่อ",
    "สถานที่ออกใบอนุญาต",
]


def _parse_permit_text(body_text: str) -> dict[str, str]:
    """Pattern: label\\n[เว้นบรรทัด]\\nvalue\\n — เก็บ label→ค่า next-non-empty-line"""
    lines = [ln.strip() for ln in (body_text or "").splitlines()]
    out: dict[str, str] = {}
    for i, ln in enumerate(lines):
        for label in PERMIT_LABELS:
            if ln == label and label not in out:
                # หา next non-empty line ที่ไม่ใช่ label อีกตัว
                for j in range(i + 1, min(i + 6, len(lines))):
                    val = lines[j]
                    if val and val not in PERMIT_LABELS and val != ":":
                        out[label] = val
                        break
                break
    return out


def _logout_safely(page: Page) -> None:
    try:
        page.evaluate("typeof logout === 'function' && logout(new Event('click'))")
        page.wait_for_timeout(2000)
    except Exception:
        pass
    try:
        page.context.clear_cookies()
    except Exception:
        pass


def _capture_alien_permit_info(
    page: Page,
    email: str,
    pwd: str,
    basename: str,
    screenshot_dir: Path,
    log=print,
) -> dict[str, Any]:
    """หลังลงทะเบียนสำเร็จ → login ใหม่ + capture หน้า ข้อมูลใบอนุญาต"""
    out: dict[str, Any] = {"PermitFields": {}, "PermitText": "", "PermitScreenshot": "", "PermitError": ""}
    _logout_safely(page)
    try:
        page.goto(LOGIN_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(2500)
        page.locator("#User_typeLogin").select_option(value="1")
        page.wait_for_timeout(1500)
        page.locator("#username_login").fill(email)
        page.locator("#password_login").fill(pwd)
        page.locator("#validate_login").click()
        page.wait_for_timeout(6000)
        if "/Login" in (page.url or ""):
            alert = _capture_register_alert(page) or "login ไม่สำเร็จ"
            out["PermitError"] = f"login: {alert[:200]}"
            ss = screenshot_dir / f"{basename}_login_fail.png"
            try: page.screenshot(path=str(ss), full_page=True)
            except Exception: pass
            return out
        _switch_lang_th(page)
        page.goto("https://eworkpermit.doe.go.th/Profile/ProfileDigitalWorkPermit", wait_until="domcontentloaded")
        page.wait_for_timeout(4000)
        _switch_lang_th(page)
        ss = screenshot_dir / f"{basename}_permit_info.png"
        try: page.screenshot(path=str(ss), full_page=True)
        except Exception: pass
        out["PermitScreenshot"] = ss.name
        try:
            text = page.evaluate("() => document.body.innerText") or ""
        except Exception:
            text = ""
        out["PermitText"] = text[:8000]
        out["PermitFields"] = _parse_permit_text(text)
        log(f"     📑 ดึงข้อมูลใบอนุญาตแล้ว ({len(out['PermitFields'])} ฟิลด์)")
        # เก็บข้อมูลเพิ่มจากหน้า /profile_alien (รหัสอ้างอิง ฯลฯ)
        try:
            page.goto("https://eworkpermit.doe.go.th/profile_alien",
                      wait_until="domcontentloaded", timeout=45000)
            page.wait_for_timeout(3000)
            _switch_lang_th(page)
            page.wait_for_timeout(1500)
            ss2 = screenshot_dir / f"{basename}_profile_alien.png"
            try: page.screenshot(path=str(ss2), full_page=True)
            except Exception: pass
            out["ProfileScreenshot"] = ss2.name
            try:
                ptext = page.evaluate("() => document.body.innerText") or ""
            except Exception:
                ptext = ""
            import re as _re
            extra: dict[str, str] = {}
            patterns = {
                "ReferenceCode": r"รหัสอ้างอิง\s*[:：]\s*([A-Za-z0-9]+)",
                "AlienCode": r"รหัสคนต่างด้าว\s*[:：]\s*([0-9]+)",
                "ApplicationCode": r"อยู่ระหว่างการยื่นคำขอ\s*[:：]\s*([A-Za-z0-9]+)",
            }
            for key, pat in patterns.items():
                m = _re.search(pat, ptext)
                if m:
                    extra[key] = m.group(1).strip()
            # สถานะใบอนุญาต (badge ด้านบน) — หาจากบรรทัดที่มีคำว่า ใบอนุญาต และไม่มี ":"
            for line in ptext.splitlines()[:30]:
                s = line.strip()
                if s in ("ใบอนุญาตหมดอายุ", "ใบอนุญาตปกติ",
                         "อยู่ระหว่างการยื่นคำขอ", "ไม่มีใบอนุญาต"):
                    extra["PermitStatus"] = s
                    break
            out["PermitFields"].update(extra)
            log(f"     🔖 รหัสอ้างอิง: {extra.get('ReferenceCode', '-')}")
        except Exception as e:
            log(f"     ⚠ profile_alien err: {e}")
    except Exception as e:
        out["PermitError"] = str(e)[:300]
        log(f"     ⚠ permit capture err: {e}")
    return out


def _save_register_report(rows: list[dict[str, Any]], out_path: Path, log=print) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "ผลลงทะเบียน"
    headers = ["No.", "Name", "TaxID", "PassportNo.", "Email", "Status",
               "AttemptedWith", "AlertText", "ScreenshotAlert", "ScreenshotSuccess", "Error"]
    ws.append(headers)
    head_fill = PatternFill("solid", fgColor="305496")
    link_font = Font(color="0563C1", underline="single")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    text_cols = {"TaxID", "PassportNo.", "No."}
    link_cols = {"ScreenshotAlert", "ScreenshotSuccess"}
    ss_subdir = "screenshots/register"

    def _hyperlink(fname: str) -> str:
        # ใช้ formula =HYPERLINK("path","display") — เสถียรกว่า cell.hyperlink เมื่อมีหลายลิงก์
        target = f"{ss_subdir}/{fname}".replace('"', "")
        display = fname.replace('"', "")
        return f'=HYPERLINK("{target}","{display}")'

    for r in rows:
        ws.append([r.get(h, "") for h in headers])
        row_idx = ws.max_row
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=row_idx, column=c_idx)
            if h in text_cols and cell.value not in (None, ""):
                cell.value = str(cell.value)
                cell.number_format = "@"
            if h in link_cols and cell.value:
                cell.value = _hyperlink(str(cell.value))
                cell.font = link_font

    widths = [6, 30, 18, 16, 32, 10, 16, 60, 30, 30, 50]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions

    # Sheet 2: ข้อมูลใบอนุญาต (จาก section 2)
    ws2 = wb.create_sheet("ข้อมูลใบอนุญาต")
    extra_cols = ["ReferenceCode", "AlienCode", "ApplicationCode", "PermitStatus"]
    permit_headers = (
        ["No.", "Name", "Email"]
        + extra_cols
        + ["PermitScreenshot", "ProfileScreenshot", "PermitError"]
        + PERMIT_LABELS
        + ["PermitText"]
    )
    ws2.append(permit_headers)
    for cell in ws2[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    permit_text_cols = {"No.", "AlienCode", "ApplicationCode",
                        "เลขประจำตัวคนต่างด้าว", "ใบอนุญาตทำงานเลขที่"}
    permit_link_cols = {"PermitScreenshot", "ProfileScreenshot"}
    for r in rows:
        if r.get("Status") != "SUCCESS":
            continue
        fields = r.get("PermitFields", {}) or {}
        row_vals = (
            [r.get("No.", ""), r.get("Name", ""), r.get("Email", "")]
            + [fields.get(k, "") for k in extra_cols]
            + [r.get("PermitScreenshot", ""), r.get("ProfileScreenshot", ""),
               r.get("PermitError", "")]
            + [fields.get(lbl, "") for lbl in PERMIT_LABELS]
            + [r.get("PermitText", "")]
        )
        ws2.append(row_vals)
        row_idx = ws2.max_row
        for c_idx, h in enumerate(permit_headers, start=1):
            cell = ws2.cell(row=row_idx, column=c_idx)
            if h in permit_text_cols and cell.value not in (None, ""):
                cell.value = str(cell.value)
                cell.number_format = "@"
            if h in permit_link_cols and cell.value:
                cell.value = _hyperlink(str(cell.value))
                cell.font = link_font
    p_widths = (
        [6, 30, 32]
        + [22, 18, 22, 18]
        + [30, 30, 30]
        + [22] * len(PERMIT_LABELS)
        + [80]
    )
    for i, w in enumerate(p_widths, start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.freeze_panes = "A2"
    if ws2.max_row > 1:
        ws2.auto_filter.ref = ws2.dimensions

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    log(f"บันทึกรายงาน: {out_path}")


def run_register(
    cfg: dict,
    excel_input: Path,
    out_path: Path,
    row_range: str | None = None,
    log=print,
    progress=None,
    is_cancelled=None,
) -> tuple[int, Path]:
    """ลงทะเบียนคนต่างด้าวจาก Excel ทีละแถว
    Excel ต้องมีคอลัมน์: No., Name, TaxID, PassportNo., Password, Email
    """
    records = _read_register_excel(excel_input)
    total = len(records)
    indices = _parse_row_range(row_range, total)
    out_path = _timestamped_path(out_path)
    log(f"[1/3] อ่าน Excel: {excel_input} ({total} แถว) → จะลงทะเบียน {len(indices)} แถว: {row_range or 'ทั้งหมด'}")
    log(f"      ไฟล์รายงาน: {out_path.name}")
    if progress:
        try: progress(0, len(indices))
        except Exception: pass

    screenshot_dir = out_path.parent / "screenshots" / "register"
    screenshot_dir.mkdir(parents=True, exist_ok=True)

    results: list[dict[str, Any]] = []
    with sync_playwright() as pw:
        browser = _launch_chromium(pw, cfg, ["--ignore-certificate-errors", "--start-maximized"])
        ctx = browser.new_context(
            locale="th-TH",
            ignore_https_errors=True,
            viewport={"width": 1920, "height": 1080},
        )
        page = ctx.new_page()
        try:
            for k, idx in enumerate(indices, start=1):
                if is_cancelled and is_cancelled():
                    log("[!] ผู้ใช้ยกเลิก — หยุด"); break
                rec = records[idx - 1]
                log(f"[2/3] ({k}/{len(indices)}) แถว {idx}")
                res = _register_one_alien(page, rec, screenshot_dir, log=log)
                res["RowIndex"] = idx
                # Section 2: capture ข้อมูลใบอนุญาต ถ้าลงทะเบียนสำเร็จ
                if res.get("Status") == "SUCCESS":
                    email = str(rec.get("Email") or "").strip()
                    pwd = str(rec.get("Password") or "").strip()
                    base = res.get("Basename") or _make_basename(rec)
                    permit = _capture_alien_permit_info(page, email, pwd, base, screenshot_dir, log=log)
                    res.update(permit)
                results.append(res)
                if progress:
                    try: progress(k, len(indices))
                    except Exception: pass
        finally:
            ctx.close(); browser.close()

    log(f"[3/3] บันทึกรายงาน: {out_path}")
    _save_register_report(results, out_path, log=log)
    success = sum(1 for r in results if r.get("Status") == "SUCCESS")
    fail = sum(1 for r in results if r.get("Status") == "FAIL")
    err = sum(1 for r in results if r.get("Status") == "ERROR")
    log(f"สรุป: SUCCESS={success}, FAIL={fail}, ERROR={err} (รวม {len(results)} แถว)")
    return success, out_path


# ─────────────────────────────────────────────────────────────────
# โหมดดาวน์โหลดเอกสาร (ใบเสร็จ 400 / บต.44 / บต.22)
#   อ่าน RequestData.xlsx (เลขคำขอ + username + passport) + UsernameLogin.xlsx (รหัสผ่าน + type)
#   จัดกลุ่มตาม username → login ครั้งเดียว/บัญชี → ค้นหาเลขคำขอ → เปิด detail
#   → ดาวน์โหลดเอกสารตามที่เลือก (รองรับหลายประเภท)
#   ตั้งชื่อไฟล์: {PASSPORT}_{SUFFIX}.pdf  เช่น MD586245_RECEIPT400.pdf, MD586245_BT44.pdf
# ─────────────────────────────────────────────────────────────────
RECEIPT_AMOUNT = "400"
RECEIPT_SUFFIX = "RECEIPT400"

# กำหนดประเภทเอกสารที่ดาวน์โหลดได้จากหน้า detail
#   key (สำหรับ CLI / GUI): "receipt", "bt44", "bt22"
#   suffix: ส่วนท้ายของชื่อไฟล์ (ก่อน .pdf)
#   label: ข้อความที่ใช้แสดงใน log/report
#   tab_pattern: regex (JS) สำหรับหาชื่อแท็บที่ต้องคลิก
#   find: วิธีหาปุ่มดาวน์โหลดในแท็บ
#     - "button": คลิกปุ่ม/ลิงก์ที่ข้อความตรงกับ button_pattern (เช่น ใบเสร็จ)
#     - "row_link": หาแถวที่ข้อความตรงกับ label_pattern แล้วคลิกลิงก์ GetDocumentConfirm ในแถวนั้น
#                   (ใช้กับแท็บ 'เอกสารตอบรับจากระบบ' ที่ปุ่มเป็นไอคอนไม่มี text)
DOC_TYPES: dict[str, dict[str, Any]] = {
    "receipt": {
        "suffix": "RECEIPT",
        "label": "ใบเสร็จ",
        "tab_pattern": r"ชำระเงิน",
        "find": "button",
        "button_pattern": r"หลักฐานการชำระเงิน",
        # ใบเสร็จมีได้หลายใบ/หลายราคา → ดาวน์โหลดทุกใบ แล้วตั้งชื่อตามราคาใน PDF
        "multi": True,
    },
    "bt44": {
        "suffix": "BT44",
        "label": "แบบ บต.44",
        "tab_pattern": r"เอกสารตอบรับ",
        "find": "row_link",
        "label_pattern": r"บต\.?\s*44",
    },
    "bt22": {
        "suffix": "BT22",
        "label": "แบบ บต.22",
        "tab_pattern": r"เอกสารตอบรับ",
        "find": "row_link",
        "label_pattern": r"บต\.?\s*22",
    },
    "bt53": {
        "suffix": "BT53",
        "label": "แบบ บต.53",
        "tab_pattern": r"เอกสารตอบรับ",
        "find": "row_link",
        "label_pattern": r"บต\.?\s*53",
    },
    "bt55": {
        "suffix": "BT55",
        "label": "แบบ บต.55",
        "tab_pattern": r"เอกสารตอบรับ",
        "find": "row_link",
        "label_pattern": r"บต\.?\s*55",
        # 1 คำขออาจมีเอกสาร บต.55 แยกตามคนหลายฉบับ → ดาวน์โหลดทุกฉบับ
        # แล้วตั้งชื่อไฟล์จากข้อมูลใน PDF (ชื่อ/เลขที่/ใบอนุญาตทำงานเลขที่)
        "per_person": True,
    },
    "bt52": {
        "suffix": "BT52",
        "label": "แบบ บต.52",
        "tab_pattern": r"เอกสารตอบรับ",
        "find": "row_link",
        # บนเว็บไม่เขียน 'บต.52' — ใช้ชื่อเต็ม 'แบบแจ้งการจ้างคนต่างด้าวทำงาน ตามมาตรา 13 วรรคหนึ่ง'
        "label_pattern": r"แบบแจ้งการจ้างคนต่างด้าวทำงาน",
    },
    "bt56": {
        "suffix": "BT56",
        "label": "แบบ บต.56",
        "tab_pattern": r"เอกสารตอบรับ",
        "find": "row_link",
        "label_pattern": r"บต\.?\s*56",
    },
}
DOC_TYPES_DEFAULT: list[str] = ["receipt", "bt44", "bt22"]
_TITLE_PREFIX_RE = re.compile(
    r"^(?:miss|mrs|mr|ms|master|mstr|นาย|นาง|นางสาว|เด็กชาย|เด็กหญิง)\.?\s+",
    re.IGNORECASE,
)


def _receipt_safe_name(s: Any) -> str:
    """แปลงเป็นชื่อไฟล์ปลอดภัย โดยคงช่องว่างไว้ (ตัดเฉพาะอักขระต้องห้าม)"""
    s = str(s or "").strip()
    s = re.sub(r'[\\/:*?"<>|\r\n\t]+', " ", s)
    s = re.sub(r"\s{2,}", " ", s).strip()
    return s or "x"


def _normalize_name_suffix(s: Any) -> str:
    """normalize ส่วนต่อท้ายชื่อไฟล์ที่ผู้ใช้กำหนดเอง
    - ว่าง → คืน "" (ไม่ต่อท้าย)
    - ตัดอักขระต้องห้ามในชื่อไฟล์ออก
    - ถ้าไม่ได้ขึ้นต้นด้วยตัวคั่น (_ หรือ -) → เติม "_" ให้อัตโนมัติ
      เช่น "IO" → "_IO", "_IO" → "_IO"
    """
    s = re.sub(r'[\\/:*?"<>|\r\n\t]+', "", str(s or "").strip())
    if not s:
        return ""
    if s[0] not in ("_", "-"):
        s = "_" + s
    return s


def _strip_title(name: str) -> str:
    """ตัดคำนำหน้า (Miss/Mr/นางสาว ฯลฯ) ออกจากชื่อ"""
    return _TITLE_PREFIX_RE.sub("", (name or "").strip()).strip()


def _read_request_data(path: Path) -> list[dict[str, Any]]:
    """อ่าน RequestData.xlsx → [{seq, name_eng, req_no, username, passport, row_index}]
    คอลัมน์: ลำดับ, ชื่อคนต่างด้าว(Eng), เลขที่คำขอ, Username, PASSPORT NUMBER
    """
    p = Path(path)
    if not str(p).strip():
        raise FileNotFoundError(
            "ไม่ได้ระบุไฟล์ RequestData.xlsx — โหมด 'ดาวน์โหลดใบเสร็จ' ต้องมีไฟล์รายการคำขอ\n"
            "กด 'เลือก...' เพื่อชี้ไปที่ไฟล์ Excel (คอลัมน์: ลำดับ, ชื่อคนต่างด้าว(Eng), เลขที่คำขอ, Username, PASSPORT NUMBER)"
        )
    if not p.exists():
        raise FileNotFoundError(
            f"ไม่พบไฟล์ '{p}' — โหมด 'ดาวน์โหลดใบเสร็จ' ต้องการไฟล์ RequestData.xlsx\n"
            "ตรวจสอบ path ให้ถูก หรือกด 'เลือก...' เพื่อชี้ไปที่ไฟล์ Excel"
        )
    if p.suffix.lower() not in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        raise ValueError(
            f"ไฟล์ '{p.name}' ไม่ใช่ Excel — รองรับเฉพาะ .xlsx / .xlsm / .xltx / .xltm"
        )
    wb = load_workbook(p, data_only=True)
    ws = wb.active
    hdr = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    def col(*names: str) -> int:
        for nm in names:
            for i, h in enumerate(hdr):
                if h == nm:
                    return i
        # fallback แบบ substring
        for nm in names:
            for i, h in enumerate(hdr):
                if nm in h:
                    return i
        return -1

    i_seq = col("ลำดับ", "No.", "No")
    i_name = col("ชื่อคนต่างด้าว(Eng)", "ชื่อคนต่างด้าว", "Name")
    i_req = col("เลขที่คำขอ", "เลขคำขอ")
    i_user = col("Username", "username", "ref username", "Ref")
    i_passport = col("PASSPORT NUMBER", "Passport Number", "PassportNo.", "Passport No.", "passport", "Passport")

    rows: list[dict[str, Any]] = []
    for ridx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=1):
        if not any(v not in (None, "") for v in row):
            continue

        def g(i: int) -> str:
            return str(row[i]).strip() if 0 <= i < len(row) and row[i] is not None else ""

        rec = {
            "seq": g(i_seq) or str(ridx),
            "name_eng": g(i_name),
            "req_no": g(i_req),
            "username": g(i_user),
            "passport": g(i_passport),
            "row_index": ridx,
        }
        if rec["req_no"]:
            rows.append(rec)
    return rows


def _read_ref_numbers(path: Path) -> list[dict[str, Any]]:
    """อ่าน Ref_number.xlsx → [{seq, name_eng:'', req_no, username, passport:'', row_index}]
    คอลัมน์: Ref_number (เลขคำขอ), user (Username สำหรับ login)
    คืนรูปแบบเดียวกับ _read_request_data เพื่อให้ใช้ flow ดาวน์โหลดร่วมกันได้
    """
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    hdr = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    def col(*names: str) -> int:
        for nm in names:
            for i, h in enumerate(hdr):
                if h.lower() == nm.lower():
                    return i
        for nm in names:
            for i, h in enumerate(hdr):
                if nm.lower() in h.lower():
                    return i
        return -1

    i_req = col("Ref_number", "Ref number", "RefNumber", "Ref", "เลขที่คำขอ", "เลขคำขอ")
    i_user = col("user", "Username", "username", "ผู้ใช้งาน", "บัญชี")

    rows: list[dict[str, Any]] = []
    for ridx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=1):
        if not any(v not in (None, "") for v in row):
            continue

        def g(i: int) -> str:
            if not (0 <= i < len(row)) or row[i] is None:
                return ""
            v = row[i]
            # Excel เก็บเลขเป็น number → ตัด .0 ท้าย (เลขคำขอ/username เป็นสตริงตัวเลขล้วน)
            if isinstance(v, float) and v.is_integer():
                v = int(v)
            return str(v).strip()

        rec = {
            "seq": str(ridx),
            "name_eng": "",
            "req_no": g(i_req),
            "username": g(i_user),
            "passport": "",
            "row_index": ridx,
        }
        if rec["req_no"]:
            rows.append(rec)
    return rows


def _normalize_login_method(text: str, default: str = "E-Workpermit") -> str:
    """แปลงค่าคอลัมน์ 'ระบบ' จาก Excel → ชื่อ method ที่ login() เข้าใจ
    - มีคำว่า service → 'E-Service'  (radio ค่า 1)
    - มีคำว่า tracking / workpermit → 'E-Workpermit'  (radio ค่า 2)
    - ว่าง/ไม่รู้จัก → คืน default
    """
    t = (text or "").strip().lower().replace("-", "").replace(" ", "")
    if not t:
        return default
    if "service" in t:
        return "E-Service"
    if "tracking" in t or "workpermit" in t:
        return "E-Workpermit"
    return default


def _read_login_accounts(path: Path) -> dict[str, dict[str, str]]:
    """อ่าน UsernameLogin.xlsx → {username_lower: {username, password, type, method}}
    คอลัมน์: Username, Password, Type, ระบบ (e-Service / e-Tracking)
    - คอลัมน์ 'ระบบ' (System/Method) ระบุระบบ login ของแต่ละบัญชี
      'e-Service' → เข้าระบบ DOE e-Service, 'e-Tracking' → เข้าระบบ E-Workpermit
      ถ้าไม่มีคอลัมน์นี้/เว้นว่าง จะ fallback ไปใช้ค่าที่เลือกในหน้าโปรแกรม
    """
    p = Path(path) if path else Path("")
    if not str(p).strip():
        raise FileNotFoundError(
            "ไม่ได้ระบุไฟล์ UsernameLogin.xlsx — โหมดนี้ต้องมีไฟล์รายการบัญชี login\n"
            "กด 'เลือก...' เพื่อชี้ไปที่ไฟล์ Excel (คอลัมน์: Username, Password, Type, ระบบ)"
        )
    if not p.exists():
        raise FileNotFoundError(
            f"ไม่พบไฟล์ '{p}' — ต้องการไฟล์ UsernameLogin.xlsx\n"
            "ตรวจสอบ path ให้ถูก หรือกด 'เลือก...' เพื่อชี้ไปที่ไฟล์ Excel"
        )
    if p.suffix.lower() not in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        raise ValueError(
            f"ไฟล์ '{p.name}' ไม่ใช่ Excel — รองรับเฉพาะ .xlsx / .xlsm / .xltx / .xltm"
        )
    wb = load_workbook(p, data_only=True)
    ws = wb.active
    hdr = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    def col(*names: str) -> int:
        for nm in names:
            for i, h in enumerate(hdr):
                if h == nm:
                    return i
        for nm in names:
            for i, h in enumerate(hdr):
                if nm.lower() in h.lower():
                    return i
        return -1

    i_user = col("Username", "username")
    i_pwd = col("Password", "password")
    i_type = col("Type", "ประเภทผู้ใช้งาน", "ประเภทผู้ใช้")
    i_method = col("ระบบ", "System", "Method", "method")

    out: dict[str, dict[str, str]] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(v not in (None, "") for v in row):
            continue

        def g(i: int) -> str:
            return str(row[i]).strip() if 0 <= i < len(row) and row[i] is not None else ""

        u = g(i_user)
        if not u:
            continue
        out[u.lower()] = {
            "username": u,
            "password": g(i_pwd),
            "type": g(i_type) or "ผู้กระทำการแทน",
            # "" = ไม่ระบุในไฟล์ → ผู้เรียกจะ fallback ไปใช้ค่าจากหน้าโปรแกรม
            "method": _normalize_login_method(g(i_method), default=""),
        }
    return out


def _search_request(page: Page, req_no: str) -> None:
    """กรอกเลขคำขอในช่องค้นหา #search_group_id แล้วกดปุ่มค้นหา"""
    page.evaluate(
        r"""(reqNo) => {
            const t = document.querySelector('#search_group_id');
            if (t) {
                t.value = reqNo;
                t.dispatchEvent(new Event('input', {bubbles:true}));
                t.dispatchEvent(new Event('keyup', {bubbles:true}));
                t.dispatchEvent(new Event('change', {bubbles:true}));
            }
        }""",
        req_no,
    )
    page.wait_for_timeout(600)
    page.evaluate(
        r"""() => {
            const btns = [...document.querySelectorAll('button, a.btn, input[type=button], input[type=submit]')]
                .filter(b => b.offsetParent !== null);
            const f = btns.find(b => /ค้นหา/.test((b.innerText || b.value || '')));
            if (f) f.click();
        }"""
    )
    page.wait_for_timeout(3500)


def _open_first_detail(page: Page) -> bool:
    """คลิกลิงก์ openDetail แถวแรกของผลค้นหา → ไปหน้า detail. คืน True ถ้าสำเร็จ"""
    has = page.evaluate(
        r"""() => !!document.querySelector('a[onclick*="openDetail"]')"""
    )
    if not has:
        return False
    try:
        with page.expect_navigation(timeout=20_000, wait_until="domcontentloaded"):
            page.evaluate(
                r"""() => { const a = document.querySelector('a[onclick*="openDetail"]'); if (a) a.click(); }"""
            )
        page.wait_for_timeout(2500)
        return True
    except Exception:
        return False


def _extract_alien_eng_name(page: Page) -> str:
    """คลิกแท็บ 'ข้อมูลคนต่างด้าว' แล้วดึงค่า 'ชื่อคนต่างด้าว(Eng)'
    retry 2 รอบ (fast-fail ~1.6s) — คืน '' ถ้าดึงไม่ได้ (อย่า fallback ไปชื่อ login)
    """
    read_js = r"""() => {
        const active = [...document.querySelectorAll('.tab-pane')]
            .find(p => p.classList.contains('active') || p.offsetParent !== null);
        const text = active ? (active.innerText || '') : document.body.innerText || '';
        const lines = text.split(/\n+/).map(s => s.trim()).filter(Boolean);
        for (let i = 0; i < lines.length; i++) {
            if (/ชื่อคนต่างด้าว\s*\(\s*Eng\s*\)/i.test(lines[i])) {
                // ค่าถัดไปคือชื่อ (ข้ามบรรทัดว่าง)
                for (let j = i + 1; j < lines.length && j <= i + 3; j++) {
                    const v = lines[j];
                    if (v && !/ชื่อคนต่างด้าว/.test(v)) return v;
                }
            }
        }
        return '';
    }"""
    # PASS 0: อ่านจาก tab ปัจจุบันก่อน (0 wait) — เผื่อ tab 'ข้อมูลคนต่างด้าว' active อยู่แล้ว
    try:
        v = (page.evaluate(read_js) or "").strip()
        if v:
            return v
    except Exception:
        pass
    # PASS 1-2: คลิก tab แล้วรอ (fast-fail — 2 รอบ × 800ms = ~1.6s worst case)
    for attempt in range(2):
        page.evaluate(
            r"""() => {
                const f = [...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a')]
                    .find(a => /ข้อมูลคนต่างด้าว/.test(a.innerText || ''));
                if (f) f.click();
            }"""
        )
        page.wait_for_timeout(800)
        try:
            name = page.evaluate(read_js)
        except Exception:
            name = ""
        name = (name or "").strip()
        if name:
            return name
    return ""


def _extract_employer_name_th(page: Page) -> str:
    """ดึงค่า 'ชื่อสถานประกอบการ(ไทย)' (ในกลุ่ม 'ข้อมูลนายจ้าง')

    Optimized path: ค่านี้อยู่บนทั้ง tab 'ข้อมูลคนต่างด้าว' และ 'คำขออนุญาต'
    ปกติเรามาที่นี่หลัง _extract_alien_eng_name (ซึ่งเปิด tab 'ข้อมูลคนต่างด้าว' ค้างไว้)
    → อ่านจาก tab ปัจจุบันก่อน (0 click, 0 wait) — ประหยัดเวลามาก
    → ถ้าไม่เจอ ค่อย fallback ไปคลิก tab 'คำขออนุญาต' (retry แค่ 2 ครั้ง)

    ใช้สำหรับตั้งชื่อโฟลเดอร์กลุ่มไฟล์ตามบริษัทของนายจ้าง (โหมดดาวน์โหลดใบแจ้งผล)
    """
    read_js = r"""() => {
        const norm = s => (s||'').replace(/\s+/g,' ').trim();
        // 1) ลองแบบ label-form-info / form-info ก่อน (โครงสร้าง Bootstrap ของหน้า detail)
        const root = document.querySelector('.tab-content, .container, main, body') || document.body;
        const all = Array.from(root.querySelectorAll('.label-form-info, .form-info'));
        for (let i = 0; i < all.length; i++) {
            const el = all[i];
            const cls = el.className || '';
            const text = norm(el.innerText || el.textContent || '');
            if (!cls.includes('label-form-info')) continue;
            if (!/ชื่อสถานประกอบการ\s*\(\s*ไทย\s*\)/.test(text)) continue;
            for (let j = i + 1; j < all.length; j++) {
                const next = all[j];
                const ncls = next.className || '';
                if (ncls.includes('label-form-info')) break;
                if (ncls.includes('form-info')) {
                    const v = norm(next.innerText || next.textContent || '');
                    if (v) return v;
                }
            }
        }
        // 2) fallback: parse text ทั้ง tab-pane ที่ active
        const active = [...document.querySelectorAll('.tab-pane')]
            .find(p => p.classList.contains('active') || p.offsetParent !== null);
        const text = active ? (active.innerText || '') : document.body.innerText || '';
        const lines = text.split(/\n+/).map(s => s.trim()).filter(Boolean);
        for (let i = 0; i < lines.length; i++) {
            if (/ชื่อสถานประกอบการ\s*\(\s*ไทย\s*\)/.test(lines[i])) {
                for (let j = i + 1; j < lines.length && j <= i + 3; j++) {
                    const v = lines[j];
                    if (v && !/ชื่อสถานประกอบการ/.test(v)) return v;
                }
            }
        }
        return '';
    }"""

    # PASS 1: อ่านจาก tab ปัจจุบัน (มักเป็น 'ข้อมูลคนต่างด้าว' ที่ยังค้างอยู่)
    # ไม่คลิก tab เพิ่ม ไม่รอ — ใช้ได้ทันทีถ้าฟิลด์อยู่ตรงนั้น
    try:
        v = (page.evaluate(read_js) or "").strip()
        if v:
            return v
    except Exception:
        pass

    # PASS 2: fallback คลิก tab 'คำขออนุญาต' — retry 2 รอบพอ
    for attempt in range(2):
        try:
            page.evaluate(
                r"""() => {
                    const f = [...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a, button')]
                        .find(a => /^\s*คำขออนุญาต\s*$/.test(a.innerText || a.textContent || ''));
                    if (f) f.click();
                }"""
            )
        except Exception:
            pass
        page.wait_for_timeout(800 if attempt == 0 else 600)
        try:
            v = (page.evaluate(read_js) or "").strip()
        except Exception:
            v = ""
        if v:
            return v
    return ""


def _company_folder_name(s: str, max_len: int = 80) -> str:
    """ทำชื่อโฟลเดอร์ปลอดภัยจากชื่อบริษัท (ตัดอักขระต้องห้าม + จำกัดความยาว)
    - Windows ห้าม: \\ / : * ? " < > | และ trailing space/dot
    - คงคำภาษาไทย/อังกฤษ/ตัวเลข/ช่องว่างไว้เพื่ออ่านง่าย
    """
    s = str(s or "").strip()
    if not s:
        return ""
    s = re.sub(r'[\\/:*?"<>|\r\n\t]+', " ", s)
    s = re.sub(r"\s{2,}", " ", s).strip(" .-_")
    if len(s) > max_len:
        s = s[:max_len].rstrip(" .-_")
    return s


def _extract_doc_url(raw: str, base_url: str) -> str:
    """ดึง URL ของเอกสารจากค่า onclick/href ที่อ่านมาจากปุ่ม/ลิงก์
    รองรับรูปแบบ window.open('...'), href ตรง ๆ ฯลฯ คืน absolute URL หรือ '' ถ้าหาไม่เจอ
    (ใช้ทำทางลัด fetch ตรงในหน้าเดิม เพื่อความเร็ว)
    """
    if not raw:
        return ""
    import re as _re
    from urllib.parse import urljoin
    # หา URL ในเครื่องหมายคำพูดก่อน (เช่นใน window.open('xxx'))
    cands: list[str] = []
    for m in _re.finditer(r"""['"]([^'"]+)['"]""", raw):
        cands.append(m.group(1))
    # เผื่อ href เปล่า ๆ ที่ไม่มี quote
    for part in raw.split("|"):
        part = part.strip()
        if part and ("/" in part or part.lower().endswith(".pdf")):
            cands.append(part)
    best = ""
    for c in cands:
        cl = c.lower()
        if not c or c in ("#", "javascript:void(0)", "about:blank"):
            continue
        if "getdocument" in cl or ".pdf" in cl or "/document" in cl or cl.startswith("http") or cl.startswith("/"):
            best = c
            if "getdocument" in cl or ".pdf" in cl:
                break
    if not best:
        return ""
    try:
        return urljoin(base_url, best)
    except Exception:
        return best


def _popup_fetch_bytes(popup, url: str) -> tuple[bytes | None, str]:
    """fetch ไฟล์ภายในหน้าต่าง popup (มี cookie/referer/session เดียวกับที่เปิดจริง)
    รองรับทั้ง blob: และ http(s). คืน (body, error)
    """
    if not url or "about:blank" in url:
        return None, "url ว่าง"
    try:
        res = popup.evaluate(
            r"""async (u) => {
                try {
                    const r = await fetch(u, {credentials: 'include'});
                    if (!r.ok) return {ok: false, status: r.status};
                    const buf = await r.arrayBuffer();
                    const a = new Uint8Array(buf);
                    let s = ''; const chunk = 0x8000;
                    for (let i = 0; i < a.length; i += chunk) {
                        s += String.fromCharCode.apply(null, a.subarray(i, i + chunk));
                    }
                    return {ok: true, b64: btoa(s)};
                } catch (e) {
                    return {ok: false, err: String(e)};
                }
            }""",
            url,
        )
    except Exception as e:
        return None, str(e).splitlines()[0][:150]
    if not res or not res.get("ok"):
        if res and res.get("status"):
            return None, f"HTTP {res['status']}"
        return None, (res.get("err") if res else "fetch ล้มเหลว") or "fetch ล้มเหลว"
    b64 = res.get("b64")
    if not b64:
        return None, "อ่านไฟล์ไม่สำเร็จ"
    import base64
    try:
        return base64.b64decode(b64), ""
    except Exception as e:
        return None, str(e)[:120]


def _grab_pdf_after_click(
    page: Page, do_click, log=print, prefetch_url: str = "",
) -> tuple[bytes | None, str]:
    """คลิกปุ่ม/ลิงก์เปิดเอกสาร แล้วคว้าไฟล์ PDF ให้ได้ไม่ว่าระบบจะส่งแบบไหน:
    (1) popup เปิด PDF inline → fetch ภายใน popup
    (2) เซิร์ฟเวอร์ส่งเป็น attachment → download event
    (3) อย่างอื่น → response ที่ content-type เป็น pdf
    คืน (body, error). ครอบคลุมความต่างของ Chromium/OS แต่ละเครื่อง
    """
    ctx = page.context
    captured: list = []

    def _on_resp(resp):
        try:
            url = (resp.url or "").lower()
            hdr = resp.headers or {}
        except Exception:
            return
        ct = (hdr.get("content-type") or "").lower()
        cd = (hdr.get("content-disposition") or "").lower()
        if "pdf" in ct or url.endswith(".pdf") or "attachment" in cd or "getdocument" in url:
            captured.append(resp)

    download_holder: dict = {}

    def _on_download(d):
        download_holder["d"] = d

    # ── ทางลัด: fetch URL ของเอกสารตรงในหน้าเดิม (เร็วสุด ไม่ต้องเปิด popup) ──
    if prefetch_url:
        body, _err = _popup_fetch_bytes(page, prefetch_url)
        if body and len(body) >= 500 and body[:5] == b"%PDF-":
            return body, ""

    # ลงทะเบียน listener "ก่อน" คลิก เพื่อไม่พลาด download/response ที่มาทันที
    ctx.on("response", _on_resp)
    page.on("download", _on_download)
    popup = None

    try:
        # คลิกพร้อมพยายามจับ popup (กรณี viewer). ถ้าระบบส่งเป็น download/inline
        # โดยไม่เปิด popup ก็จะ timeout แล้วไปทาง fallback
        try:
            with page.expect_popup(timeout=15_000) as pinfo:
                do_click()
            popup = pinfo.value
        except PWTimeoutError:
            pass
        except Exception:
            pass

        if popup:
            try:
                popup.on("download", _on_download)
            except Exception:
                pass

        # (1) popup เปิด PDF inline → รอ url จริง (บาง viewer เปิด about:blank ก่อน)
        #     แล้ว fetch ภายใน popup (มี session/referer เดียวกัน) พร้อม retry
        if popup:
            purl = ""
            for _ in range(20):  # รอสูงสุด ~20 วินาที ให้เอกสารถูกสร้าง/นำทาง
                try:
                    purl = popup.url or ""
                except Exception:
                    purl = ""
                if purl and "about:blank" not in purl:
                    break
                try:
                    popup.wait_for_load_state("load", timeout=1_000)
                except Exception:
                    pass
                page.wait_for_timeout(1_000)
            if purl and "about:blank" not in purl:
                for _ in range(3):  # retry fetch เผื่อ PDF ยังสตรีมไม่เสร็จ
                    body, _err = _popup_fetch_bytes(popup, purl)
                    if body and len(body) >= 500:
                        return body, ""
                    page.wait_for_timeout(1_500)

        # ให้เวลาทาง download/response ทำงาน (กรณีไม่มี popup หรือ popup ว่าง)
        for _ in range(8):
            if download_holder.get("d") is not None or captured:
                break
            page.wait_for_timeout(1_000)

        # (2) attachment → download event
        dl = download_holder.get("d")
        if dl is not None:
            try:
                p = dl.path()
                if p:
                    b = Path(p).read_bytes()
                    if b and len(b) >= 500:
                        return b, ""
            except Exception:
                pass

        # (3) response ที่เป็น PDF
        for resp in reversed(captured):
            try:
                b = resp.body()
                if b and len(b) >= 500:
                    return b, ""
            except Exception:
                continue

        return None, "ไม่พบไฟล์ PDF จากการคลิก (อาจยังไม่ถูกสร้าง)"
    finally:
        try:
            ctx.remove_listener("response", _on_resp)
        except Exception:
            pass
        try:
            page.remove_listener("download", _on_download)
        except Exception:
            pass
        try:
            if popup:
                popup.close()
        except Exception:
            pass





def _download_doc_pdf(
    page: Page,
    doc_cfg: dict[str, str],
    out_path: Path,
    log=print,
) -> str:
    """คลิกแท็บที่ match tab_pattern → หาปุ่มดาวน์โหลดตาม find mode → ดาวน์โหลด PDF (popup)
    คืน '' ถ้าสำเร็จ, มิฉะนั้นคืนข้อความ error
    """
    tab_pattern = doc_cfg["tab_pattern"]
    find_mode = doc_cfg.get("find", "button")

    # 1) คลิกแท็บ
    clicked = page.evaluate(
        r"""(pat) => {
            const re = new RegExp(pat);
            const f = [...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a')]
                .find(a => re.test(a.innerText || ''));
            if (f) { f.click(); return true; }
            return false;
        }""",
        tab_pattern,
    )
    if not clicked:
        return f"ไม่พบแท็บที่ตรงกับ '{tab_pattern}'"
    page.wait_for_timeout(1800)

    # 2) เตรียม JS สำหรับ "เช็คว่ามีปุ่ม" และ "คลิกปุ่ม" ตาม find mode
    if find_mode == "row_link":
        pattern = doc_cfg["label_pattern"]
        has_js = r"""(pat) => {
            const re = new RegExp(pat);
            const pane = document.querySelector('#tab-response') || document;
            let found = false;
            pane.querySelectorAll('*').forEach(el => {
                if (found) return;
                const txt = (el.innerText || '').trim();
                if (!txt || !re.test(txt)) return;
                const links = [...el.querySelectorAll('a, button, [onclick]')]
                    .filter(b => (b.getAttribute('onclick') || '').includes('GetDocumentConfirm'));
                if (links.length === 1) found = true;
            });
            return found;
        }"""
        click_js = r"""(pat) => {
            const re = new RegExp(pat);
            const pane = document.querySelector('#tab-response') || document;
            let best = null, bestLen = Infinity;
            pane.querySelectorAll('*').forEach(el => {
                const txt = (el.innerText || '').trim();
                if (!txt || !re.test(txt)) return;
                const links = [...el.querySelectorAll('a, button, [onclick]')]
                    .filter(b => (b.getAttribute('onclick') || '').includes('GetDocumentConfirm'));
                if (links.length === 1 && txt.length < bestLen) {
                    best = links[0]; bestLen = txt.length;
                }
            });
            if (best) { best.click(); return true; }
            return false;
        }"""
        url_js = r"""(pat) => {
            const re = new RegExp(pat);
            const pane = document.querySelector('#tab-response') || document;
            let best = null, bestLen = Infinity;
            pane.querySelectorAll('*').forEach(el => {
                const txt = (el.innerText || '').trim();
                if (!txt || !re.test(txt)) return;
                const links = [...el.querySelectorAll('a, button, [onclick]')]
                    .filter(b => (b.getAttribute('onclick') || '').includes('GetDocumentConfirm'));
                if (links.length === 1 && txt.length < bestLen) {
                    best = links[0]; bestLen = txt.length;
                }
            });
            if (!best) return '';
            return (best.getAttribute('onclick') || '') + ' | ' + (best.getAttribute('href') || '');
        }"""
        not_found_msg = f"ไม่พบเอกสาร '{doc_cfg['label']}' ในแท็บเอกสารตอบรับ (อาจยังไม่ถูกสร้าง)"
    else:  # "button"
        pattern = doc_cfg["button_pattern"]
        has_js = r"""(pat) => {
            const re = new RegExp(pat);
            return !![...document.querySelectorAll('button, a, [onclick]')].find(b =>
                b.offsetParent !== null && re.test(b.innerText || ''));
        }"""
        click_js = r"""(pat) => {
            const re = new RegExp(pat);
            const f = [...document.querySelectorAll('button, a, [onclick]')].find(b =>
                b.offsetParent !== null && re.test(b.innerText || ''));
            if (f) f.click();
        }"""
        url_js = r"""(pat) => {
            const re = new RegExp(pat);
            const f = [...document.querySelectorAll('button, a, [onclick]')].find(b =>
                b.offsetParent !== null && re.test(b.innerText || ''));
            if (!f) return '';
            return (f.getAttribute('onclick') || '') + ' | ' + (f.getAttribute('href') || '');
        }"""
        not_found_msg = f"ไม่พบปุ่ม/ลิงก์ที่ตรงกับ '{pattern}' ในแท็บนี้"

    if not page.evaluate(has_js, pattern):
        return not_found_msg

    # 3) ลองดึง URL เอกสารจากปุ่ม/ลิงก์ เพื่อ fetch ตรง (เร็วกว่าเปิด popup)
    prefetch_url = ""
    try:
        raw = page.evaluate(url_js, pattern) or ""
        prefetch_url = _extract_doc_url(raw, page.url)
    except Exception:
        prefetch_url = ""

    # 4) คลิกปุ่ม แล้วคว้าไฟล์ PDF (รองรับ popup/inline/attachment ทุกเครื่อง)
    body, err = _grab_pdf_after_click(
        page, lambda: page.evaluate(click_js, pattern), log=log,
        prefetch_url=prefetch_url,
    )
    if err:
        return f"ดาวน์โหลด PDF ไม่สำเร็จ: {err}"
    if not body or len(body) < 500:
        return f"ไฟล์ PDF เล็กผิดปกติ ({len(body) if body else 0} bytes)"
    try:
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_bytes(body)
        return ""
    except Exception as e:
        return f"บันทึก PDF ล้มเหลว: {str(e).splitlines()[0][:150]}"





def _extract_pdf_amount(pdf_bytes: bytes) -> str:
    """อ่านยอดเงินจากไฟล์ PDF ใบเสร็จ → คืนเป็นสตริง เช่น '100', '400', '1800'
    (ว่าง = หาไม่เจอ)
    - 1 ไฟล์ PDF อาจมีใบเสร็จหลายหน้า/หลายใบ (เช่น หน้า1 ค่ายื่นคำขอ 100 + หน้า2 ค่าธรรมเนียม 300)
      → รวมยอด 'รวมเป็นเงินทั้งสิ้น' ของทุกใบเข้าด้วยกัน (100+300 = 400)
    - ถ้าหาวลี 'รวมเป็นเงินทั้งสิ้น' ไม่เจอ → fallback ใช้ตัวเลขเงินที่มากสุด (พฤติกรรมเดิม)
    """
    try:
        import io
        from pypdf import PdfReader

        reader = PdfReader(io.BytesIO(pdf_bytes))
        text = "\n".join((pg.extract_text() or "") for pg in reader.pages)
    except Exception:
        return ""

    def _fmt(total: float) -> str:
        if total <= 0:
            return ""
        if total == int(total):
            return str(int(total))
        return f"{total:.2f}".rstrip("0").rstrip(".").replace(".", "_")

    _money = r"\d{1,3}(?:,\d{3})*\.\d{2}"
    # ยอด 'รวมเป็นเงินทั้งสิ้น' ต่อ 1 ใบเสร็จ (เลขตามหลังวลีทันที) → รวมทุกใบในไฟล์เดียว
    totals = re.findall(rf"รวมเป็นเงินทั้งสิ้น[^\d]{{0,40}}({_money})", text)
    vals = [v for a in totals if (v := float(a.replace(",", ""))) > 0]
    if vals:
        return _fmt(sum(vals))

    # fallback: ไม่พบวลีรวมยอด → เลือกตัวเลขเงินที่มากสุด (กันเลขคำขอ/พาสปอร์ตปน)
    fvals = [v for a in re.findall(_money, text) if (v := float(a.replace(",", ""))) > 0]
    if not fvals:
        return ""
    return _fmt(max(fvals))


# เดือนไทย + รูปแบบชื่อที่ยอมรับ (ใช้พาร์สข้อมูลในเอกสาร บต.55)
_BT55_TH_MONTHS = (
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม",
)
_BT55_NAME_OK = re.compile(r"^[A-Za-zก-๙][A-Za-zก-๙\s.\-']{1,58}$")


def _bt55_parse_pdf(pdf_bytes: bytes) -> dict[str, str]:
    """อ่านข้อมูลในกรอบสีแดงของเอกสาร บต.55 (เป็น PDF text-base) เพื่อนำไปตั้งชื่อไฟล์
    คืน {"name", "doc_no", "work_permit"} — ช่องไหนหาไม่เจอจะเป็น "" (ไม่ต้องระบุ)
      - name        : ชื่อคนต่างด้าว (เช่น MISS SWE ZIN MYINT)
      - doc_no      : เลขที่ / เลขหนังสือเดินทาง (เช่น MJ197150)
      - work_permit : ใบอนุญาตทำงานเลขที่ (เลข 13 หลัก เช่น 1400670022679)
    """
    out = {"name": "", "doc_no": "", "work_permit": ""}
    try:
        import io
        from pypdf import PdfReader

        reader = PdfReader(io.BytesIO(pdf_bytes))
        text = "\n".join((pg.extract_text() or "") for pg in reader.pages)
    except Exception:
        return out
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    # เลขที่ (หนังสือเดินทาง): 1-2 ตัวพิมพ์ใหญ่ ตามด้วยตัวเลข 6-9 หลัก
    m = re.search(r"\b([A-Z]{1,2}\d{6,9})\b", text)
    if m:
        out["doc_no"] = m.group(1)

    # ใบอนุญาตทำงานเลขที่: เลข 13 หลักที่ไม่ขึ้นต้นด้วย 0
    # (เลขทะเบียนนิติบุคคลของบริษัทก็ 13 หลักแต่ขึ้นต้นด้วย 0 → ตัดทิ้ง)
    thirteens = re.findall(r"(?<!\d)(\d{13})(?!\d)", text)
    non_company = [t for t in thirteens if not t.startswith("0")]
    out["work_permit"] = (
        non_company[0] if non_company else (thirteens[-1] if thirteens else "")
    )

    # ชื่อ: บรรทัดถัดจาก "วัน/เดือนไทย/ปี พ.ศ." ที่ถูกแยกเป็น 3 บรรทัดชุดแรก
    for i in range(len(lines) - 3):
        if (
            re.fullmatch(r"\d{1,2}", lines[i])
            and lines[i + 1] in _BT55_TH_MONTHS
            and re.fullmatch(r"\d{4}", lines[i + 2])
        ):
            cand = lines[i + 3]
            if _BT55_NAME_OK.match(cand) and cand not in _BT55_TH_MONTHS:
                out["name"] = cand
            break
    return out


def _download_all_receipts(
    page: Page,
    receipts_dir: Path,
    passport_safe: str,
    log=print,
    name_suffix: str = "",
) -> list[dict[str, str]]:
    """ดาวน์โหลด 'หลักฐานการชำระเงิน' ทุกใบในแท็บการชำระเงิน
    คืน list ของ {amount, file, status, error} (1 รายการต่อ 1 ใบเสร็จ)
    ตั้งชื่อไฟล์ตามราคาที่อ่านได้จากใน PDF: {PASSPORT}_RECEIPT{ราคา}{name_suffix}.pdf
    """
    results: list[dict[str, str]] = []

    # 1) คลิกแท็บการชำระเงิน
    clicked = page.evaluate(
        r"""() => {
            const re = /ชำระเงิน/;
            const f = [...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a')]
                .find(a => re.test(a.innerText || ''));
            if (f) { f.click(); return true; }
            return false;
        }"""
    )
    if not clicked:
        return [{"amount": "", "file": "", "status": "FAIL", "error": "ไม่พบแท็บการชำระเงิน"}]
    page.wait_for_timeout(1600)

    # 2) นับจำนวนปุ่ม 'หลักฐานการชำระเงิน' ที่มองเห็น
    count = page.evaluate(
        r"""() => [...document.querySelectorAll('button, a, [onclick]')]
            .filter(b => b.offsetParent !== null && /หลักฐานการชำระเงิน/.test(b.innerText || '')).length"""
    )
    if not count:
        return [{"amount": "", "file": "", "status": "FAIL",
                 "error": "ไม่พบปุ่มหลักฐานการชำระเงิน (อาจยังไม่มีใบเสร็จ)"}]

    used_names: set[str] = set()
    for idx in range(count):
        try:
            def _click_receipt(i=idx):
                page.evaluate(
                    r"""(i) => {
                        const btns = [...document.querySelectorAll('button, a, [onclick]')]
                            .filter(b => b.offsetParent !== null && /หลักฐานการชำระเงิน/.test(b.innerText || ''));
                        if (btns[i]) btns[i].click();
                    }""",
                    i,
                )

            # ลองดึง URL ของใบเสร็จใบนี้ เพื่อ fetch ตรง (เร็วกว่าเปิด popup)
            prefetch_url = ""
            try:
                raw = page.evaluate(
                    r"""(i) => {
                        const btns = [...document.querySelectorAll('button, a, [onclick]')]
                            .filter(b => b.offsetParent !== null && /หลักฐานการชำระเงิน/.test(b.innerText || ''));
                        const b = btns[i];
                        if (!b) return '';
                        return (b.getAttribute('onclick') || '') + ' | ' + (b.getAttribute('href') || '');
                    }""",
                    idx,
                ) or ""
                prefetch_url = _extract_doc_url(raw, page.url)
            except Exception:
                prefetch_url = ""

            body, err = _grab_pdf_after_click(
                page, _click_receipt, log=log, prefetch_url=prefetch_url,
            )
            if err:
                results.append({"amount": "", "file": "", "status": "FAIL",
                                "error": f"ใบที่ {idx + 1}: ดาวน์โหลดไม่สำเร็จ ({err})"})
                continue
            if not body or len(body) < 500:
                results.append({"amount": "", "file": "", "status": "FAIL",
                                "error": f"ใบที่ {idx + 1}: ไฟล์ PDF เล็กผิดปกติ ({len(body) if body else 0} bytes)"})
                continue

            amount = _extract_pdf_amount(body) or f"x{idx + 1}"
            base = f"{passport_safe}_RECEIPT{amount}{name_suffix}"
            name = base + ".pdf"
            n = 2
            while name in used_names:  # ราคาซ้ำกันในการรันเดียว → เพิ่มเลขท้าย
                name = f"{base}_{n}.pdf"
                n += 1
            used_names.add(name)
            receipts_dir.mkdir(parents=True, exist_ok=True)
            (receipts_dir / name).write_bytes(body)
            results.append({"amount": amount, "file": name, "status": "SUCCESS", "error": ""})
            log(f"     ✓ ใบเสร็จ {amount} บาท: {name}")
        except Exception as e:
            results.append({"amount": "", "file": "", "status": "FAIL",
                            "error": f"ใบที่ {idx + 1}: {str(e).splitlines()[0][:120]}"})
    return results


def _norm_id(s: Any) -> str:
    """normalize เลขที่เอกสาร/passport เพื่อเทียบ: ตัดช่องว่าง-ขีด แล้วทำเป็นพิมพ์ใหญ่"""
    return re.sub(r"[\s\-]", "", str(s or "")).upper()


def _bt55_filename_stem(
    info: dict[str, str],
    req_no: str,
    people: list[dict[str, Any]] | None,
    fallback: str = "",
) -> str:
    """ตั้งชื่อไฟล์ บต.55 (ส่วน stem ไม่รวม _BT55.pdf) รูปแบบ {PASSPORT}_{NAME}_{เลขคำขอ}
    - จับคู่กับแถวใน Excel (people) ด้วย passport = เลขที่ในเอกสาร หรือเลข 13 หลัก
      → ใช้ passport+ชื่อ จาก Excel
    - ถ้าไม่เจอคู่ → ใช้ เลขที่(passport)+ชื่อ จากเนื้อหา PDF แทน
    - ถ้ายังว่าง → คืน fallback
    """
    doc_no = info.get("doc_no", "")
    work_permit = info.get("work_permit", "")
    matched = None
    for p in (people or []):
        key = _norm_id(p.get("passport", ""))
        if key and key in (_norm_id(doc_no), _norm_id(work_permit)):
            matched = p
            break
    if matched:
        passport, name = matched.get("passport", ""), matched.get("name_eng", "")
    else:
        passport, name = doc_no, info.get("name", "")
    # ระบุตัวบุคคลไม่ได้เลย (ไม่มีทั้ง passport และชื่อ) → ใช้ fallback กันไฟล์ชนกันในคำขอเดียวกัน
    if not (passport or name):
        return fallback or req_no
    stem = "_".join(_safe_filename(x) for x in (passport, name, req_no) if x)
    return stem or fallback


def _download_bt55_per_person(
    page: Page,
    receipts_dir: Path,
    passport_safe: str,
    req_no: str = "",
    people: list[dict[str, Any]] | None = None,
    log=print,
    name_suffix: str = "",
) -> list[dict[str, str]]:
    """ดาวน์โหลดเอกสาร 'แบบ บต.55' ทุกฉบับในแท็บเอกสารตอบรับ
    (1 คำขออาจมีเอกสาร บต.55 แยกตามคนหลายฉบับ — ตัวเดิมดึงได้แค่ฉบับเดียว/พลาด)
    ตั้งชื่อไฟล์: {PASSPORT}_{NAME}_{เลขคำขอ}_BT55.pdf
    - จับคู่แต่ละไฟล์กับแถวใน Excel (people) ด้วย passport/เลข 13 หลัก → ใช้ passport+ชื่อจาก Excel
    - ถ้าไม่เจอคู่ → ใช้ เลขที่(passport)+ชื่อ จากเนื้อหา PDF แทน
    คืน list ของ {name, doc_no, work_permit, file, status, error} (1 รายการ/1 ฉบับ)
    """
    pat = r"บต\.?\s*55"

    # 1) คลิกแท็บเอกสารตอบรับ
    clicked = page.evaluate(
        r"""() => {
            const re = /เอกสารตอบรับ/;
            const f = [...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a')]
                .find(a => re.test(a.innerText || ''));
            if (f) { f.click(); return true; }
            return false;
        }"""
    )
    if not clicked:
        return [{"name": "", "doc_no": "", "work_permit": "", "file": "",
                 "status": "FAIL", "error": "ไม่พบแท็บเอกสารตอบรับ"}]
    page.wait_for_timeout(1800)

    # 2) นับจำนวนเอกสาร บต.55
    #    โครงสร้างจริง (#tab_default_4 → #DetailDocumentList): เป็น grid แบบ flat
    #    คอลัมน์ชื่อเอกสาร <div class="col-7/col-8"> กับปุ่มดาวน์โหลด
    #    <div class="row col-3"><a onclick="GetDocumentConfirm(...)"> เป็น "พี่น้องติดกัน"
    #    (ปุ่มไม่ได้อยู่ใน element เดียวกับ label) → จับคู่ผ่าน parentElement.previousElementSibling
    #    แล้วกรองเฉพาะ label ที่เป็น บต.55 (กัน บต.52/บต.56 และแถวแม่ที่ไม่มีปุ่ม) + dedup ด้วย onclick
    enum_js = r"""(pat) => {
        const re = new RegExp(pat);
        const seen = new Set();
        let n = 0;
        document.querySelectorAll('[onclick*="GetDocumentConfirm"]').forEach(a => {
            const oc = a.getAttribute('onclick') || '';
            if (!oc) return;
            const grp = a.parentElement;
            const lbl = (grp && grp.previousElementSibling)
                ? (grp.previousElementSibling.innerText || '') : '';
            if (!re.test(lbl)) return;
            if (seen.has(oc)) return;
            seen.add(oc); n++;
        });
        return n;
    }"""
    count = page.evaluate(enum_js, pat)
    if not count:
        return [{"name": "", "doc_no": "", "work_permit": "", "file": "",
                 "status": "FAIL",
                 "error": "ไม่พบเอกสาร บต.55 ในแท็บเอกสารตอบรับ (อาจยังไม่ถูกสร้าง)"}]

    # JS เลือกลิงก์ลำดับที่ i ด้วยลำดับ/ตัวกรอง dedup เดียวกับ enum_js
    # (mode='url' อ่าน onclick/href, mode='click' สั่งคลิก)
    pick_js = r"""(args) => {
        const { pat, i, mode } = args;
        const re = new RegExp(pat);
        const seen = new Set(); const list = [];
        document.querySelectorAll('[onclick*="GetDocumentConfirm"]').forEach(a => {
            const oc = a.getAttribute('onclick') || '';
            if (!oc) return;
            const grp = a.parentElement;
            const lbl = (grp && grp.previousElementSibling)
                ? (grp.previousElementSibling.innerText || '') : '';
            if (!re.test(lbl)) return;
            if (seen.has(oc)) return;
            seen.add(oc); list.push(a);
        });
        const b = list[i];
        if (!b) return mode === 'url' ? '' : false;
        if (mode === 'url') return (b.getAttribute('onclick') || '') + ' | ' + (b.getAttribute('href') || '');
        b.click(); return true;
    }"""

    results: list[dict[str, str]] = []
    for idx in range(count):
        try:
            # ลองดึง URL เอกสารฉบับนี้ เพื่อ fetch ตรง (เร็วกว่าเปิด popup)
            prefetch_url = ""
            try:
                raw = page.evaluate(pick_js, {"pat": pat, "i": idx, "mode": "url"}) or ""
                prefetch_url = _extract_doc_url(raw, page.url)
            except Exception:
                prefetch_url = ""

            body, err = _grab_pdf_after_click(
                page,
                lambda i=idx: page.evaluate(pick_js, {"pat": pat, "i": i, "mode": "click"}),
                log=log,
                prefetch_url=prefetch_url,
            )
            if err:
                results.append({"name": "", "doc_no": "", "work_permit": "", "file": "",
                                "status": "FAIL",
                                "error": f"ฉบับที่ {idx + 1}: ดาวน์โหลดไม่สำเร็จ ({err})"})
                continue
            if not body or len(body) < 500:
                results.append({"name": "", "doc_no": "", "work_permit": "", "file": "",
                                "status": "FAIL",
                                "error": f"ฉบับที่ {idx + 1}: ไฟล์ PDF เล็กผิดปกติ "
                                         f"({len(body) if body else 0} bytes)"})
                continue

            info = _bt55_parse_pdf(body)
            stem = _bt55_filename_stem(info, req_no, people, fallback=f"{passport_safe}_{idx + 1}")
            name = f"{stem}_BT55{name_suffix}.pdf"
            # ชื่อซ้ำ (รันซ้ำ/หลายแถวคำขอเดียวกัน) → ทับไฟล์เดิม ไม่หลบชื่อ จะได้ไฟล์เดียว
            receipts_dir.mkdir(parents=True, exist_ok=True)
            (receipts_dir / name).write_bytes(body)
            results.append({**info, "file": name, "status": "SUCCESS", "error": ""})
            log(f"     ✓ บต.55: {name}")
        except Exception as e:
            results.append({"name": "", "doc_no": "", "work_permit": "", "file": "",
                            "status": "FAIL",
                            "error": f"ฉบับที่ {idx + 1}: {str(e).splitlines()[0][:120]}"})
    return results


def _save_receipt_report(
    rows: list[dict[str, Any]],
    out_path: Path,
    doc_types: list[str],
    log=print,
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "ผลดาวน์โหลดเอกสาร"
    base_headers = ["ลำดับ", "PASSPORT", "ชื่อ(Excel)", "ชื่อ(Eng จากระบบ)",
                    "เลขที่คำขอ", "Username", "Status รวม"]
    headers: list[str] = list(base_headers)
    for dt in doc_types:
        label = DOC_TYPES[dt]["label"]
        headers.append(f"Status {label}")
        headers.append(f"ไฟล์ {label}")
    headers.append("Error")
    ws.append(headers)

    head_fill = PatternFill("solid", fgColor="305496")
    link_font = Font(color="0563C1", underline="single")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    text_cols = {"ลำดับ", "PASSPORT", "เลขที่คำขอ"}
    link_cols = {f"ไฟล์ {DOC_TYPES[dt]['label']}" for dt in doc_types
                 if not DOC_TYPES[dt].get("multi") and not DOC_TYPES[dt].get("per_person")}

    for r in rows:
        row_vals: list[Any] = [
            r.get("seq", ""), r.get("passport", ""), r.get("name_excel", ""),
            r.get("name_eng", ""), r.get("req_no", ""), r.get("username", ""),
            r.get("status", ""),
        ]
        docs = r.get("docs", {})
        for dt in doc_types:
            d = docs.get(dt, {})
            row_vals.append(d.get("status", ""))
            row_vals.append(d.get("pdf_file", ""))
        row_vals.append(r.get("error", ""))
        ws.append(row_vals)

        row_idx = ws.max_row
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=row_idx, column=c_idx)
            if h in text_cols and cell.value not in (None, ""):
                cell.value = str(cell.value)
                cell.number_format = "@"
            if h in link_cols and cell.value:
                fname = str(cell.value)
                cell.value = f'=HYPERLINK("receipts/{fname}","{fname}")'
                cell.font = link_font

    base_widths = [8, 22, 28, 28, 20, 30, 12]
    doc_widths: list[int] = []
    for _ in doc_types:
        doc_widths.extend([14, 36])
    widths = base_widths + doc_widths + [55]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    log(f"บันทึกรายงาน: {out_path}")


def run_receipts(
    cfg: dict,
    request_excel: Path,
    login_excel: Path,
    out_path: Path,
    row_range: str | None = None,
    doc_types: list[str] | None = None,
    name_suffix: str = "",
    make_subfolder: bool = False,
    log=print,
    progress=None,
    is_cancelled=None,
) -> tuple[int, Path]:
    """ดาวน์โหลดเอกสาร (ใบเสร็จ 400 / บต.44 / บต.22) ตามเลขคำขอใน RequestData.xlsx
    - จัดกลุ่มตาม Username → login ครั้งเดียว/บัญชี
    - รองรับ session timeout: login ใหม่ด้วยบัญชีเดิมแล้วทำต่อ
    - resume ได้: ข้ามไฟล์ PDF ที่มีอยู่แล้ว (แยกตามประเภท)
    - doc_types: รายการคีย์จาก DOC_TYPES (default = ทั้งหมด)
    - name_suffix: ส่วนต่อท้ายชื่อไฟล์ที่ผู้ใช้กำหนด เช่น "_IO" → {PASSPORT}_BT22_IO.pdf (ว่าง = ไม่ต่อท้าย)
    - make_subfolder: True = แยกไฟล์ลงโฟลเดอร์ย่อยชื่อ {PASSPORT} ของแต่ละคน
    """
    doc_types = [d for d in (doc_types or DOC_TYPES_DEFAULT) if d in DOC_TYPES]
    if not doc_types:
        raise ValueError("ต้องเลือกประเภทเอกสารอย่างน้อย 1 อย่าง")
    name_suffix = _normalize_name_suffix(name_suffix)
    out_path = _timestamped_path(out_path)
    records = _read_request_data(request_excel)
    accounts = _read_login_accounts(login_excel)
    total = len(records)
    indices = _parse_row_range(row_range, total)
    selected = [records[i - 1] for i in indices]
    log(f"[1/3] อ่าน RequestData: {request_excel} ({total} แถว) → จะทำ {len(selected)} แถว: {row_range or 'ทั้งหมด'}")
    log(f"      บัญชี login: {login_excel} ({len(accounts)} บัญชี)")
    log(f"      เอกสารที่จะดาวน์โหลด: {', '.join(DOC_TYPES[d]['label'] for d in doc_types)}")
    if name_suffix:
        log(f"      ต่อท้ายชื่อไฟล์: '{name_suffix}' (เช่น {{PASSPORT}}_BT44{name_suffix}.pdf)")
    if make_subfolder:
        log("      แยกโฟลเดอร์ตาม PASSPORT: receipts/{PASSPORT}/...")

    receipts_dir = out_path.parent / "receipts"
    receipts_dir.mkdir(parents=True, exist_ok=True)

    # จัดกลุ่มตาม username (คงลำดับเดิม)
    groups: dict[str, list[dict[str, Any]]] = {}
    for rec in selected:
        groups.setdefault(rec["username"], []).append(rec)

    results: list[dict[str, Any]] = []
    done_count = 0
    success = 0
    # เอกสารแยกตามคน (บต.55) ดาวน์โหลด "ครั้งเดียวต่อ 1 เลขคำขอ" — กันโหลดซ้ำเมื่อหลายแถวเป็นคำขอเดียวกัน
    per_person_done: dict[tuple[str, str], str] = {}
    # map เลขคำขอ → รายชื่อคนใน Excel (passport+ชื่อ) เพื่อจับคู่ตั้งชื่อไฟล์ บต.55
    req_people: dict[str, list[dict[str, Any]]] = {}
    for _rec in records:
        req_people.setdefault(_rec.get("req_no", ""), []).append(
            {"passport": _rec.get("passport", ""), "name_eng": _rec.get("name_eng", "")}
        )
    # หา PASSPORT ที่ถูกใช้ซ้ำกับหลายเลขคำขอ (เช่นคอลัมน์ PASSPORT เป็นชื่อ LOT/ก้อน ไม่ใช่พาสปอร์ตจริง)
    # ชื่อไฟล์ปกติเป็น {PASSPORT}_BT44.pdf → ถ้า passport ซ้ำ ไฟล์จะชนกันจนถูกมองว่า SKIP_EXISTS
    _pp_to_reqs: dict[str, set[str]] = {}
    for _rec in records:
        _pp = _receipt_safe_name(_rec.get("passport", "") or _rec.get("seq", "") or "")
        _pp_to_reqs.setdefault(_pp, set()).add(_rec.get("req_no", ""))
    ambiguous_passports = {pp for pp, rq in _pp_to_reqs.items() if len(rq) > 1}
    if ambiguous_passports:
        log(f"      ⚠ คอลัมน์ PASSPORT ซ้ำข้ามหลายเลขคำขอ {len(ambiguous_passports)} ค่า "
            f"→ จะเติมเลขคำขอในชื่อไฟล์กันชนกัน (เช่น {{PASSPORT}}_{{เลขคำขอ}}_BT44.pdf)")
    if progress:
        try: progress(0, len(selected))
        except Exception: pass

    headless = bool(cfg.get("headless", False))
    with sync_playwright() as pw:
        browser = _launch_chromium(pw, cfg, ["--ignore-certificate-errors", "--start-maximized"])
        ctx = browser.new_context(
            locale="th-TH",
            ignore_https_errors=True,
            viewport={"width": 1920, "height": 1080},
            accept_downloads=True,
        )
        page = ctx.new_page()
        try:
            for gi, (username, recs) in enumerate(groups.items(), start=1):
                if is_cancelled and is_cancelled():
                    log("[!] ผู้ใช้ยกเลิก — หยุด"); break

                acct = accounts.get((username or "").strip().lower())
                if not acct:
                    log(f"[กลุ่ม {gi}/{len(groups)}] ✗ ไม่พบบัญชีของ '{username}' ใน {login_excel.name} — ข้าม {len(recs)} รายการ")
                    for rec in recs:
                        results.append({**rec, "name_excel": rec.get("name_eng", ""),
                                        "docs": {dt: {"status": "", "pdf_file": "", "error": ""} for dt in doc_types},
                                        "status": "NO_ACCOUNT",
                                        "error": f"ไม่พบ username '{username}' ใน UsernameLogin"})
                        done_count += 1
                    if progress:
                        try: progress(done_count, len(selected))
                        except Exception: pass
                    continue

                login_cfg = {
                    "username": acct["username"],
                    "password": acct["password"],
                    "user_type": acct["type"],
                    "method": acct.get("method") or cfg.get("method", "E-Workpermit"),
                }
                log(f"[กลุ่ม {gi}/{len(groups)}] เข้าสู่ระบบ: {acct['username']} ({acct['type']}) — {len(recs)} รายการ")
                try:
                    if gi > 1:
                        _logout_safely(page)
                        page.wait_for_timeout(800)
                    login(page, login_cfg)
                except Exception as e:
                    log(f"      ✗ login ไม่สำเร็จ: {e} — ข้ามกลุ่มนี้")
                    for rec in recs:
                        results.append({**rec, "name_excel": rec.get("name_eng", ""),
                                        "docs": {dt: {"status": "", "pdf_file": "", "error": ""} for dt in doc_types},
                                        "status": "LOGIN_FAIL", "error": str(e)[:200]})
                        done_count += 1
                    continue

                for ri, rec in enumerate(recs, start=1):
                    if is_cancelled and is_cancelled():
                        log("[!] ผู้ใช้ยกเลิก — หยุด"); break
                    res = _process_one_receipt(
                        page, rec, login_cfg, receipts_dir, doc_types, log=log,
                        per_person_done=per_person_done, req_people=req_people,
                        name_suffix=name_suffix, make_subfolder=make_subfolder,
                        ambiguous_passports=ambiguous_passports,
                    )
                    results.append(res)
                    done_count += 1
                    if res.get("status") == "SUCCESS":
                        success += 1
                    if progress:
                        try: progress(done_count, len(selected))
                        except Exception: pass

                # เซฟ report เป็นระยะ (กันข้อมูลหายระหว่างทาง)
                _save_receipt_report(results, out_path, doc_types, log=lambda *_: None)
        finally:
            ctx.close(); browser.close()

    _save_receipt_report(results, out_path, doc_types, log=log)
    fail = sum(1 for r in results if r.get("status") not in ("SUCCESS",))
    log(f"[3/3] สรุป: สำเร็จ {success} / {len(results)} (ล้มเหลว/บางส่วน {fail})")
    return success, out_path


def _process_one_receipt(
    page: Page,
    rec: dict[str, Any],
    login_cfg: dict,
    receipts_dir: Path,
    doc_types: list[str],
    log=print,
    per_person_done: dict[tuple[str, str], str] | None = None,
    req_people: dict[str, list[dict[str, Any]]] | None = None,
    name_suffix: str = "",
    make_subfolder: bool = False,
    ambiguous_passports: set[str] | None = None,
) -> dict[str, Any]:
    """ค้นหา 1 เลขคำขอ → เปิด detail → ดาวน์โหลดเอกสารตามที่เลือก
    มี session recovery: ถ้าหลุด login ระหว่างทาง → login ใหม่ด้วยบัญชีเดิม แล้วลองอีกครั้ง
    per_person_done: เก็บ (เลขคำขอ, ประเภทเอกสาร) ที่โหลดแบบ "แยกตามคน" (บต.55) ไปแล้ว
      → คำขอเดียวกันในแถวถัดไปจะข้าม ไม่โหลดซ้ำ (ดาวน์โหลดครั้งเดียวต่อคำขอ)
    name_suffix: ส่วนต่อท้ายชื่อไฟล์ (เช่น "_IO") — ต่อก่อน ".pdf"
    make_subfolder: True = เก็บไฟล์ลงโฟลเดอร์ย่อยชื่อ {PASSPORT}
    """
    if per_person_done is None:
        per_person_done = {}
    seq = rec.get("seq", "")
    req_no = rec.get("req_no", "")
    name_excel = rec.get("name_eng", "")
    passport = rec.get("passport", "") or seq  # fallback = ลำดับ
    passport_safe = _receipt_safe_name(passport)
    # ถ้า passport นี้ถูกใช้ซ้ำกับหลายเลขคำขอ → เติมเลขคำขอในชื่อไฟล์ กันไฟล์ชนกัน (SKIP_EXISTS ผิด ๆ)
    req_safe = _receipt_safe_name(req_no)
    name_key = f"{passport_safe}_{req_safe}" if (req_safe and passport_safe in (ambiguous_passports or set())) else passport_safe
    # โฟลเดอร์ปลายทาง: ถ้าติ๊กสร้างโฟลเดอร์ → receipts/{PASSPORT}/ ของแต่ละคน
    # ไม่สร้างโฟลเดอร์ตรงนี้ — ปล่อยให้ path ของไฟล์สร้างแบบ lazy ตอนเขียนไฟล์จริงเท่านั้น
    # (ถ้า record ไหนไม่พบ/ดาวน์โหลดไม่สำเร็จ จะไม่มีโฟลเดอร์ว่างค้างไว้)
    target_dir = receipts_dir / passport_safe if make_subfolder else receipts_dir

    docs_state: dict[str, dict[str, str]] = {
        dt: {"status": "", "pdf_file": "", "error": ""} for dt in doc_types
    }
    res: dict[str, Any] = {
        **rec, "name_excel": name_excel, "name_eng": "",
        "docs": docs_state, "status": "", "error": "",
    }
    log(f"   • ลำดับ {seq} | passport {passport} | เลขคำขอ {req_no} | {name_excel}")

    # resume: เช็คทีละ doc — ถ้าไฟล์มีอยู่แล้วก็ skip
    todo: list[str] = []
    for dt in doc_types:
        cfg_dt = DOC_TYPES[dt]
        if cfg_dt.get("per_person"):
            # บต.55 = เอกสารแยกตามคน "ของทั้งคำขอ" → ดาวน์โหลดครั้งเดียวต่อ 1 เลขคำขอ
            # ถ้าคำขอนี้โหลดแบบแยกตามคนไปแล้ว (แถวก่อนของคำขอเดียวกัน) → ข้าม ไม่โหลดซ้ำ
            if (req_no, dt) in per_person_done:
                docs_state[dt]["status"] = "SKIP_EXISTS"
                docs_state[dt]["pdf_file"] = per_person_done[(req_no, dt)]
                log(f"     ↷ {cfg_dt['label']}: คำขอนี้ดาวน์โหลดแล้ว — ข้าม (โหลดครั้งเดียวต่อคำขอ)")
            else:
                todo.append(dt)
            continue
        if cfg_dt.get("multi"):
            existing = sorted(target_dir.glob(f"{name_key}_RECEIPT*.pdf"))
            if existing:
                docs_state[dt]["status"] = "SKIP_EXISTS"
                docs_state[dt]["pdf_file"] = ", ".join(p.name for p in existing)
                log(f"     ↷ {cfg_dt['label']}: มีไฟล์แล้ว ({len(existing)} ใบ) — ข้าม")
            else:
                todo.append(dt)
            continue
        out_pdf = target_dir / f"{name_key}_{cfg_dt['suffix']}{name_suffix}.pdf"
        if out_pdf.exists():
            docs_state[dt]["status"] = "SKIP_EXISTS"
            docs_state[dt]["pdf_file"] = out_pdf.name
            log(f"     ↷ {cfg_dt['label']}: มีไฟล์แล้ว ({out_pdf.name}) — ข้าม")
        else:
            todo.append(dt)

    def _attempt(targets: list[str]) -> list[str]:
        """ทำ 1 รอบ คืน list ของ doc_type ที่ยังควรลองใหม่ (เช่น session หลุด)"""
        goto_tracking(page)
        page.wait_for_timeout(1500)
        _search_request(page, req_no)
        if not _open_first_detail(page):
            for dt in targets:
                docs_state[dt]["status"] = "FAIL"
                docs_state[dt]["error"] = "ไม่พบผลค้นหา / เปิดรายละเอียดไม่ได้"
            return []
        name_eng = _extract_alien_eng_name(page) or name_excel
        res["name_eng"] = name_eng
        retry: list[str] = []
        for dt in targets:
            cfg_dt = DOC_TYPES[dt]
            if cfg_dt.get("per_person"):
                people = (req_people or {}).get(req_no, [])
                pp_results = _download_bt55_per_person(
                    page, target_dir, name_key, req_no, people, log=log,
                    name_suffix=name_suffix,
                )
                ok = [r for r in pp_results if r["status"] == "SUCCESS"]
                errs_here = [r["error"] for r in pp_results if r.get("error")]
                if ok:
                    docs_state[dt]["status"] = "SUCCESS"
                    docs_state[dt]["pdf_file"] = ", ".join(r["file"] for r in ok)
                    docs_state[dt]["error"] = " | ".join(errs_here)
                    per_person_done[(req_no, dt)] = docs_state[dt]["pdf_file"]
                    log(f"     ✓ {cfg_dt['label']}: ดาวน์โหลด {len(ok)} ฉบับ (แยกตามคน)")
                else:
                    docs_state[dt]["status"] = "FAIL"
                    docs_state[dt]["error"] = " | ".join(errs_here) or "ดาวน์โหลด บต.55 ไม่สำเร็จ"
                    log(f"     ✗ {cfg_dt['label']}: {docs_state[dt]['error']}")
                    if _is_logged_out(page):
                        retry.append(dt)
                        break
                continue
            if cfg_dt.get("multi"):
                rec_results = _download_all_receipts(page, target_dir, name_key, log=log, name_suffix=name_suffix)
                ok = [r for r in rec_results if r["status"] == "SUCCESS"]
                errs_here = [r["error"] for r in rec_results if r.get("error")]
                if ok:
                    docs_state[dt]["status"] = "SUCCESS"
                    docs_state[dt]["pdf_file"] = ", ".join(r["file"] for r in ok)
                    docs_state[dt]["error"] = " | ".join(errs_here)
                    log(f"     ✓ {cfg_dt['label']}: ดาวน์โหลด {len(ok)} ใบ")
                else:
                    docs_state[dt]["status"] = "FAIL"
                    docs_state[dt]["error"] = " | ".join(errs_here) or "ดาวน์โหลดใบเสร็จไม่สำเร็จ"
                    log(f"     ✗ {cfg_dt['label']}: {docs_state[dt]['error']}")
                    if _is_logged_out(page):
                        retry.append(dt)
                        break
                continue
            out_pdf = target_dir / f"{name_key}_{cfg_dt['suffix']}{name_suffix}.pdf"
            err = _download_doc_pdf(page, cfg_dt, out_pdf, log=log)
            if err:
                docs_state[dt]["status"] = "FAIL"
                docs_state[dt]["error"] = err
                log(f"     ✗ {cfg_dt['label']}: {err}")
                if _is_logged_out(page):
                    retry.append(dt)
                    break  # เลิก loop รอบนี้ ค่อย login ใหม่แล้วทำ retry list ต่อ
            else:
                docs_state[dt]["status"] = "SUCCESS"
                docs_state[dt]["pdf_file"] = out_pdf.name
                docs_state[dt]["error"] = ""
                log(f"     ✓ {cfg_dt['label']}: {out_pdf.name}")
        return retry

    if todo:
        try:
            retry = _attempt(todo)
            if retry and _is_logged_out(page):
                log("     ⚠ session หมดอายุ — login ใหม่แล้วลองอีกครั้ง")
                login(page, login_cfg)
                _attempt(retry)
        except Exception as e:
            msg = str(e).splitlines()[0][:200]
            try:
                if _is_logged_out(page):
                    log("     ⚠ session หมดอายุ (exception) — login ใหม่แล้วลองอีกครั้ง")
                    login(page, login_cfg)
                    pending = [dt for dt in todo if docs_state[dt]["status"] != "SUCCESS"]
                    _attempt(pending)
                else:
                    for dt in todo:
                        if not docs_state[dt]["status"]:
                            docs_state[dt]["status"] = "FAIL"
                            docs_state[dt]["error"] = msg
                    log(f"     ✗ {msg}")
            except Exception as e2:
                for dt in todo:
                    if docs_state[dt]["status"] not in ("SUCCESS",):
                        docs_state[dt]["status"] = "FAIL"
                        docs_state[dt]["error"] = (
                            f"{msg} | recover ล้มเหลว: {str(e2).splitlines()[0][:120]}"
                        )

    # สรุปสถานะรวม + รวม error
    statuses = [docs_state[dt]["status"] for dt in doc_types]
    if all(s in ("SUCCESS", "SKIP_EXISTS") for s in statuses):
        res["status"] = "SUCCESS"
    elif any(s == "SUCCESS" for s in statuses):
        res["status"] = "PARTIAL"
    else:
        res["status"] = "FAIL"
    errs = [
        f"{DOC_TYPES[dt]['label']}: {docs_state[dt]['error']}"
        for dt in doc_types if docs_state[dt].get("error")
    ]
    res["error"] = " | ".join(errs)
    if not res["name_eng"]:
        res["name_eng"] = name_excel
    return res


# ─────────────────────────────────────────────────────────────────
# โหมดใหม่ — ดาวน์โหลดเอกสารผลอนุญาต (ใบแจ้งผล / ใบรับคำขอ)
#   • ไม่ใช้ RequestData — ดึงรายการคำขอจาก e-Tracking (filter สถานะ + รายการคำขอ)
#   • login จาก UsernameLogin.xlsx (วนทุกบัญชี)
#   • ตั้งชื่อไฟล์ {ชื่อคนต่างด้าว(Eng)}_{ชื่อเอกสารที่แสดงบนเว็บ}.pdf
# ─────────────────────────────────────────────────────────────────
RESULT_DOC_TYPES: dict[str, dict[str, str]] = {
    "result_notice": {
        "label": "ใบแจ้งผลใบอนุญาตทำงาน",
        "label_pattern": r"ใบแจ้งผลใบอนุญาตทำงาน",
    },
    "request_receipt": {
        "label": "ใบรับคำขอใบอนุญาตทำงาน",
        "label_pattern": r"ใบรับคำขอใบอนุญาตทำงาน",
    },
    "bt50": {
        "label": "แบบ บต.50 อ.6",
        "label_pattern": r"บต\.?\s*50",
    },
}
RESULT_DOC_TYPES_DEFAULT: list[str] = ["result_notice", "request_receipt"]


def _open_detail_for_row(page: Page, row: dict) -> bool:
    """เปิดหน้า detail ของแถวที่เก็บมาจาก collect_all_wa_rows (ต้องอยู่หน้า e-Tracking)
    คืน True ถ้าเปิดสำเร็จ
    """
    # 1) คลิกลิงก์ในตาราง (ถ้าแถวยังอยู่ใน DOM)
    group_id = str(row.get("group_id", "") or "")
    if group_id:
        try:
            cand = page.locator(f'a[onclick*="openDetail"][onclick*="{group_id}"]').first
            if cand.count() > 0:
                with page.expect_navigation(timeout=20_000, wait_until="domcontentloaded"):
                    cand.click()
                page.wait_for_timeout(2500)
                return True
        except Exception:
            pass
    # 2) เรียก openDetail() ผ่าน JS โดยตรง
    try:
        args = [
            row.get("user_id", ""), row.get("group_id", ""), row.get("status", ""),
            row.get("form_type", ""), row.get("institution_id", ""), row.get("id", ""),
        ]
        with page.expect_navigation(timeout=15_000, wait_until="domcontentloaded"):
            page.evaluate(
                """(a) => { if (typeof openDetail === 'function') openDetail(a[0],a[1],a[2],a[3],a[4],a[5]); }""",
                args,
            )
        page.wait_for_timeout(2500)
        return True
    except Exception:
        return False


def _is_detail_loaded(page: Page) -> bool:
    """ตรวจว่าหน้า detail โหลดสำเร็จ (มีแท็บเอกสาร/ข้อมูลคนต่างด้าว) ไม่ใช่หน้า error/login"""
    if _is_logged_out(page):
        return False
    try:
        url = (page.url or "").lower()
    except Exception:
        url = ""
    # endpoint หน้า detail มีหลายชื่อตาม form_type แต่ทุกอันมีคำว่า 'detail'
    # (DetailTracking, DetailRequest41, DetailFormRenewMOU, DetailFormRenew59, ...)
    if "detail" not in url:
        return False
    try:
        return bool(page.evaluate(
            r"""() => !![...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a')]
                .find(a => /เอกสารตอบรับ|ข้อมูลคนต่างด้าว|คำขออนุญาต|สถานะคำขอ/.test(a.innerText || ''))"""
        ))
    except Exception:
        return False


def _open_detail_direct(page: Page, row: dict) -> bool:
    """เปิดหน้า detail โดย navigate ตรงผ่าน build_detail_url (เร็วกว่าคลิกตาราง+go_back)
    ถ้า navigate ตรงไม่สำเร็จ → fallback ไปวิธีคลิกในตาราง/เรียก openDetail()
    คืน True ถ้าเปิด detail สำเร็จ
    """
    try:
        url = build_detail_url(row)
    except Exception:
        url = ""
    if url:
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=25_000)
            page.wait_for_timeout(1500)
            if _is_detail_loaded(page):
                return True
        except Exception:
            pass
    # fallback — ต้องอยู่หน้า e-Tracking ถึงจะคลิกแถวได้
    return _open_detail_for_row(page, row)


def _download_response_doc_named(
    page: Page,
    out_dir: Path,
    name_eng_safe: str,
    doc_cfg: dict[str, str],
    log=print,
    req_no: str = "",
) -> dict[str, str]:
    """ไปแท็บ 'เอกสารตอบรับจากระบบ' → หาเอกสารตาม label_pattern → ดาวน์โหลด
    ตั้งชื่อ {เลขคำขอ}_{name_eng}_{ชื่อเอกสารบนเว็บ}.pdf  คืน {status, file, label, error}
    (ใส่เลขคำขอไว้หน้าสุดเพื่อกันชื่อซ้ำกรณีคนต่างด้าวชื่อเหมือนกัน)
    """
    pattern = doc_cfg["label_pattern"]
    res = {"status": "", "file": "", "label": "", "error": ""}

    # คลิกแท็บเอกสารตอบรับ
    clicked = page.evaluate(
        r"""() => {
            const re = /เอกสารตอบรับ/;
            const f = [...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a')]
                .find(a => re.test(a.innerText || ''));
            if (f) { f.click(); return true; }
            return false;
        }"""
    )
    if not clicked:
        res["status"] = "FAIL"
        res["error"] = "ไม่พบแท็บเอกสารตอบรับ"
        return res
    page.wait_for_timeout(1500)

    # อ่านชื่อเอกสาร (label) ที่ตรง pattern และมีลิงก์ GetDocumentConfirm 1 ลิงก์
    label_text = page.evaluate(
        r"""(pat) => {
            const re = new RegExp(pat);
            const pane = document.querySelector('#tab-response') || document;
            let bestText = '', bestLen = Infinity;
            pane.querySelectorAll('*').forEach(el => {
                const txt = (el.innerText || '').trim();
                if (!txt || !re.test(txt)) return;
                const links = [...el.querySelectorAll('a, button, [onclick]')]
                    .filter(b => (b.getAttribute('onclick') || '').includes('GetDocumentConfirm'));
                if (links.length === 1 && txt.length < bestLen) {
                    bestText = txt; bestLen = txt.length;
                }
            });
            return bestText;
        }""",
        pattern,
    )
    if not label_text:
        res["status"] = "NOT_FOUND"
        res["error"] = f"ไม่พบเอกสาร '{doc_cfg['label']}' ในแท็บเอกสารตอบรับ"
        return res

    label = re.sub(r"\s+", " ", str(label_text)).strip()
    # ตัดข้อความที่นำหน้าชื่อเอกสารจริงออก (เช่น เลขลำดับแถว "2 ", "3 ")
    m = re.search(pattern, label)
    if m:
        label = label[m.start():].strip()
    label = label[:120]
    res["label"] = label
    req_safe = _receipt_safe_name(str(req_no or "").strip())
    if req_safe:
        stem = _receipt_safe_name(f"{req_safe}_{name_eng_safe}_{label}")
    else:
        stem = _receipt_safe_name(f"{name_eng_safe}_{label}")
    out_pdf = out_dir / f"{stem}.pdf"
    if out_pdf.exists():
        res["status"] = "SKIP_EXISTS"
        res["file"] = out_pdf.name
        log(f"     ↷ {doc_cfg['label']}: มีไฟล์แล้ว ({out_pdf.name}) — ข้าม")
        return res

    # ใช้กลไกดาวน์โหลดเดิม (row_link ในแท็บเอกสารตอบรับ)
    dl_cfg = {
        "label": doc_cfg["label"],
        "tab_pattern": r"เอกสารตอบรับ",
        "find": "row_link",
        "label_pattern": pattern,
    }
    err = _download_doc_pdf(page, dl_cfg, out_pdf, log=log)
    if err:
        res["status"] = "FAIL"
        res["error"] = err
        log(f"     ✗ {doc_cfg['label']}: {err}")
    else:
        res["status"] = "SUCCESS"
        res["file"] = out_pdf.name
        log(f"     ✓ {doc_cfg['label']}: {out_pdf.name}")
    return res


def _save_result_docs_report(
    rows: list[dict[str, Any]],
    out_path: Path,
    doc_keys: list[str],
    log=print,
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "ผลดาวน์โหลดเอกสารผลอนุญาต"
    base_headers = ["ลำดับ", "ชื่อคนต่างด้าว(Eng)", "ชื่อสถานประกอบการ(ไทย)", "เลขที่คำขอ", "Username", "สถานะคำขอ", "Status รวม"]
    headers: list[str] = list(base_headers)
    for dk in doc_keys:
        label = RESULT_DOC_TYPES[dk]["label"]
        headers.append(f"Status {label}")
        headers.append(f"ไฟล์ {label}")
    headers.append("Error")
    ws.append(headers)

    head_fill = PatternFill("solid", fgColor="305496")
    link_font = Font(color="0563C1", underline="single")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    text_cols = {"ลำดับ", "เลขที่คำขอ"}
    link_cols = {f"ไฟล์ {RESULT_DOC_TYPES[dk]['label']}" for dk in doc_keys}

    for r in rows:
        row_vals: list[Any] = [
            r.get("seq", ""), r.get("name_eng", ""), r.get("employer_th", ""),
            r.get("req_no", ""), r.get("username", ""), r.get("status_text", ""), r.get("status", ""),
        ]
        docs = r.get("docs", {})
        for dk in doc_keys:
            d = docs.get(dk, {})
            row_vals.append(d.get("status", ""))
            row_vals.append(d.get("file", ""))
        row_vals.append(r.get("error", ""))
        ws.append(row_vals)

        # subfolder prefix สำหรับ HYPERLINK — ตรงตามที่ไฟล์ถูกบันทึกจริง
        subdir = str(r.get("subdir", "") or "").strip()
        link_prefix = f"result_docs/{subdir}/" if subdir else "result_docs/"

        row_idx = ws.max_row
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=row_idx, column=c_idx)
            if h in text_cols and cell.value not in (None, ""):
                cell.value = str(cell.value)
                cell.number_format = "@"
            if h in link_cols and cell.value:
                fname = str(cell.value)
                cell.value = f'=HYPERLINK("{link_prefix}{fname}","{fname}")'
                cell.font = link_font

    base_widths = [8, 30, 40, 20, 26, 22, 12]
    doc_widths: list[int] = []
    for _ in doc_keys:
        doc_widths.extend([16, 46])
    widths = base_widths + doc_widths + [55]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    log(f"บันทึกรายงาน: {out_path}")


def _process_one_result_by_ref(
    page: Page,
    rec: dict[str, Any],
    login_cfg: dict,
    docs_dir: Path,
    doc_keys: list[str],
    log=print,
    prebuilt_row: dict | None = None,
) -> dict[str, Any]:
    """ค้นหา 1 เลขคำขอ → เปิด detail → ดาวน์โหลดใบแจ้งผล/ใบรับคำขอ (เอกสารตอบรับจากระบบ)
    มี session recovery: ถ้าหลุด login ระหว่างทาง → login ใหม่ด้วยบัญชีเดิม แล้วลองอีกครั้ง
    ถ้ามี prebuilt_row (เก็บจาก e-Tracking มาแล้ว) → navigate ตรงผ่าน URL (เร็ว ไม่ต้องค้นหา)
    """
    seq = rec.get("seq", "")
    req_no = rec.get("req_no", "")
    res: dict[str, Any] = {
        "seq": seq, "req_no": req_no, "username": rec.get("username", ""),
        "name_eng": "", "employer_th": "", "subdir": "", "status_text": "",
        "docs": {dk: {"status": "", "file": "", "error": ""} for dk in doc_keys},
        "status": "", "error": "",
    }
    log(f"   • ลำดับ {seq} | เลขคำขอ {req_no} | user {rec.get('username','')}")

    def _attempt() -> bool:
        """เปิด detail. คืน True ถ้าสำเร็จ
        - ถ้ามี prebuilt_row → navigate ตรงผ่าน URL ก่อน (เร็วสุด)
        - ไม่งั้น/ถ้าไม่สำเร็จ → fallback ไปค้นหาเลขคำขอแล้วเปิดแถวแรก
        """
        if prebuilt_row:
            if _open_detail_direct(page, prebuilt_row):
                return True
        goto_tracking(page)
        page.wait_for_timeout(1200)
        _search_request(page, req_no)
        return _open_first_detail(page)

    try:
        opened = _attempt()
        if not opened and _is_logged_out(page):
            log("     ⚠ session หมดอายุ — login ใหม่แล้วลองอีกครั้ง")
            login(page, login_cfg)
            opened = _attempt()
    except Exception as e:
        if _is_logged_out(page):
            try:
                log("     ⚠ session หมดอายุ (exception) — login ใหม่แล้วลองอีกครั้ง")
                login(page, login_cfg)
                opened = _attempt()
            except Exception as e2:
                res["status"] = "FAIL"; res["error"] = str(e2).splitlines()[0][:200]
                return res
        else:
            res["status"] = "FAIL"; res["error"] = str(e).splitlines()[0][:200]
            return res

    if not opened:
        res["status"] = "NOT_FOUND"
        res["error"] = "ไม่พบคำขอจากการค้นหา (ตรวจเลขคำขอ/บัญชีให้ตรงกัน)"
        log(f"     ✗ ไม่พบคำขอ {req_no}")
        return res

    name_eng = _extract_alien_eng_name(page)
    if not name_eng:
        # ดึงชื่อไม่ได้ → ใช้ req_no เป็นชื่อ (ยังคงดาวน์โหลดต่อ ไม่ skip)
        log(f"     ⚠ {req_no}: ดึงชื่อคนต่างด้าวไม่ได้ — ใช้เลขคำขอเป็นชื่อไฟล์แทน")
    res["name_eng"] = name_eng
    name_safe = (_receipt_safe_name(name_eng) or "").strip(" -_.").strip() or _receipt_safe_name(req_no)

    # ดึงชื่อบริษัท (ไทย) จากแท็บ 'คำขออนุญาต' → ใช้ตั้งชื่อ subfolder
    employer_th = _extract_employer_name_th(page)
    res["employer_th"] = employer_th
    subdir = _company_folder_name(employer_th) or "_no_company"
    res["subdir"] = subdir
    req_docs_dir = docs_dir / subdir
    req_docs_dir.mkdir(parents=True, exist_ok=True)
    if employer_th:
        log(f"     🏢 บริษัท: {employer_th} → {subdir}/")
    else:
        log(f"     ⚠ ดึงชื่อบริษัทไม่ได้ → เก็บใน {subdir}/")

    def _download_all() -> None:
        for dk in doc_keys:
            r = _download_response_doc_named(page, req_docs_dir, name_safe, RESULT_DOC_TYPES[dk], log=log, req_no=req_no)
            res["docs"][dk] = {"status": r["status"], "file": r["file"], "error": r["error"]}

    _download_all()

    # ถ้าดาวน์โหลดไม่สำเร็จเลย และตรวจพบ session หมด (เช่น modal Session Timeout
    # โผล่กลางทาง) → login ใหม่ เปิด detail ใหม่ แล้วดาวน์โหลดอีกครั้ง
    if all(res["docs"][dk]["status"] not in ("SUCCESS", "SKIP_EXISTS") for dk in doc_keys) \
            and _is_logged_out(page):
        try:
            log("     ⚠ session หมดอายุระหว่างดาวน์โหลด — login ใหม่แล้วลองอีกครั้ง")
            login(page, login_cfg)
            if _attempt():
                name_eng = _extract_alien_eng_name(page)
                if name_eng:
                    res["name_eng"] = name_eng
                    name_safe = (_receipt_safe_name(name_eng) or "").strip(" -_.").strip() or _receipt_safe_name(req_no)
                    employer_th2 = _extract_employer_name_th(page)
                    if employer_th2:
                        res["employer_th"] = employer_th2
                        subdir = _company_folder_name(employer_th2) or "_no_company"
                        res["subdir"] = subdir
                        req_docs_dir = docs_dir / subdir
                        req_docs_dir.mkdir(parents=True, exist_ok=True)
                    _download_all()
        except Exception as _e:
            log(f"     ✗ ฟื้น session ไม่สำเร็จ: {_e}")

    statuses = [res["docs"][dk]["status"] for dk in doc_keys]
    if all(s in ("SUCCESS", "SKIP_EXISTS") for s in statuses):
        res["status"] = "SUCCESS"
    elif any(s == "SUCCESS" for s in statuses):
        res["status"] = "PARTIAL"
    else:
        res["status"] = "FAIL"
    res["error"] = " | ".join(
        f"{RESULT_DOC_TYPES[dk]['label']}: {res['docs'][dk]['error']}"
        for dk in doc_keys if res["docs"][dk].get("error")
    )
    return res


def run_result_docs_by_ref(
    cfg: dict,
    ref_excel: Path,
    login_excel: Path,
    out_path: Path,
    row_range: str | None = None,
    doc_keys: list[str] | None = None,
    log=print,
    progress=None,
    is_cancelled=None,
) -> tuple[int, Path]:
    """ดาวน์โหลดใบแจ้งผล/ใบรับคำขอ โดย 'ค้นหาตามเลขคำขอ' จาก Ref_number.xlsx
    - Ref_number.xlsx: คอลัมน์ Ref_number (เลขคำขอ) + user (Username)
    - จัดกลุ่มตาม user → login บัญชีที่ตรงใน UsernameLogin.xlsx → ค้นหาทีละเลข → โหลด
    - resume ได้: ข้ามไฟล์ PDF ที่มีอยู่แล้ว (ในตัว _download_response_doc_named)
    """
    doc_keys = [d for d in (doc_keys or RESULT_DOC_TYPES_DEFAULT) if d in RESULT_DOC_TYPES]
    if not doc_keys:
        raise ValueError("ต้องเลือกประเภทเอกสารอย่างน้อย 1 อย่าง")
    out_path = _timestamped_path(out_path)
    records = _read_ref_numbers(ref_excel)
    accounts = _read_login_accounts(login_excel)
    total = len(records)
    indices = _parse_row_range(row_range, total)
    selected = [records[i - 1] for i in indices]
    log(f"[1/3] อ่าน {ref_excel.name} ({total} เลขคำขอ) → จะทำ {len(selected)} รายการ: {row_range or 'ทั้งหมด'}")
    log(f"      บัญชี login: {login_excel.name} ({len(accounts)} บัญชี)")
    log(f"      เอกสารที่จะดาวน์โหลด: {', '.join(RESULT_DOC_TYPES[d]['label'] for d in doc_keys)}")

    docs_dir = out_path.parent / "result_docs"
    docs_dir.mkdir(parents=True, exist_ok=True)

    # จัดกลุ่มตาม username (คงลำดับเดิม)
    groups: dict[str, list[dict[str, Any]]] = {}
    for rec in selected:
        groups.setdefault(rec["username"], []).append(rec)

    results: list[dict[str, Any]] = []
    done_count = 0
    success = 0
    if progress:
        try: progress(0, len(selected))
        except Exception: pass

    headless = bool(cfg.get("headless", False))
    with sync_playwright() as pw:
        browser = _launch_chromium(pw, cfg, ["--ignore-certificate-errors", "--start-maximized"])
        ctx = browser.new_context(
            locale="th-TH", ignore_https_errors=True,
            viewport={"width": 1920, "height": 1080}, accept_downloads=True,
        )
        page = ctx.new_page()
        try:
            for gi, (username, recs) in enumerate(groups.items(), start=1):
                if is_cancelled and is_cancelled():
                    log("[!] ผู้ใช้ยกเลิก — หยุด"); break

                acct = accounts.get((username or "").strip().lower())
                if not acct:
                    log(f"[กลุ่ม {gi}/{len(groups)}] ✗ ไม่พบบัญชีของ '{username}' ใน {login_excel.name} — ข้าม {len(recs)} รายการ")
                    for rec in recs:
                        results.append({**rec, "name_eng": "", "employer_th": "", "subdir": "", "status_text": "",
                                        "docs": {dk: {"status": "", "file": "", "error": ""} for dk in doc_keys},
                                        "status": "NO_ACCOUNT",
                                        "error": f"ไม่พบ username '{username}' ใน {login_excel.name}"})
                        done_count += 1
                    if progress:
                        try: progress(done_count, len(selected))
                        except Exception: pass
                    _save_result_docs_report(results, out_path, doc_keys, log=lambda *_: None)
                    continue

                login_cfg = {
                    "username": acct["username"],
                    "password": acct["password"],
                    "user_type": acct["type"],
                    "method": acct.get("method") or cfg.get("method", "E-Workpermit"),
                }
                log(f"[กลุ่ม {gi}/{len(groups)}] เข้าสู่ระบบ: {acct['username']} ({acct['type']}) — {len(recs)} รายการ")
                try:
                    if gi > 1:
                        _logout_safely(page)
                        page.wait_for_timeout(800)
                    login(page, login_cfg)
                except Exception as e:
                    log(f"      ✗ login ไม่สำเร็จ: {e} — ข้ามกลุ่มนี้")
                    for rec in recs:
                        results.append({**rec, "name_eng": "", "employer_th": "", "subdir": "", "status_text": "",
                                        "docs": {dk: {"status": "", "file": "", "error": ""} for dk in doc_keys},
                                        "status": "LOGIN_FAIL", "error": str(e)[:200]})
                        done_count += 1
                    _save_result_docs_report(results, out_path, doc_keys, log=lambda *_: None)
                    continue

                RELOGIN_EVERY = 80  # re-login เชิงรุกทุก N รายการ กัน session timeout
                MAX_RETRY_PASSES = 3  # จำนวนรอบ retry คำขอที่ยังไม่สำเร็จ (หลังรอบแรก)

                # เก็บรายการคำขอของบัญชีนี้ครั้งเดียว (filter สถานะ AP + SS) เพื่อ map
                # เลขคำขอ → พารามิเตอร์ (user_id/form_type/group_id) สำหรับ navigate ตรง
                # เร็วกว่าค้นหาทีละเลข + ลด round-trip ที่ทำให้ดาวน์โหลดค้างในงานยาว
                row_by_req: dict[str, dict] = {}
                try:
                    collect_status_ids = list(cfg.get("result_status_ids") or ["AP", "SS"])
                    goto_tracking(page)
                    apply_wa_filter(page, "", status_ids=collect_status_ids)
                    collected = collect_all_wa_rows(page, log=log)
                    for cr in collected:
                        rn = str(cr.get("reqNo", "") or "").strip()
                        if rn:
                            row_by_req[rn] = cr
                    log(f"      เก็บรายการคำขอได้ {len(row_by_req)} เลข (สำหรับ navigate ตรง)")
                except Exception as _e:
                    log(f"      (เก็บรายการคำขอล่วงหน้าไม่สำเร็จ: {_e} — จะ fallback ไปค้นหาทีละเลข)")

                retry_pending: list[tuple[int, dict, dict | None]] = []
                for ri, rec in enumerate(recs, start=1):
                    if is_cancelled and is_cancelled():
                        log("[!] ผู้ใช้ยกเลิก — หยุด"); break
                    prebuilt = row_by_req.get(str(rec.get("req_no", "") or "").strip())
                    res = _process_one_result_by_ref(
                        page, rec, login_cfg, docs_dir, doc_keys, log=log,
                        prebuilt_row=prebuilt,
                    )
                    results.append(res)
                    res_idx = len(results) - 1
                    done_count += 1
                    if res.get("status") != "SUCCESS":
                        # ยังไม่สำเร็จครบ (FAIL/PARTIAL/NO_NAME/NOT_FOUND) → เก็บไว้ retry
                        retry_pending.append((res_idx, rec, prebuilt))
                    if progress:
                        try: progress(done_count, len(selected))
                        except Exception: pass
                    _save_result_docs_report(results, out_path, doc_keys, log=lambda *_: None)
                    # re-login เชิงรุกกัน session timeout หลังดาวน์โหลดไปสักระยะ
                    if RELOGIN_EVERY and ri % RELOGIN_EVERY == 0 and ri < len(recs):
                        log(f"      🔄 ครบ {ri} รายการ — login ใหม่เชิงรุกกัน session timeout...")
                        try:
                            login(page, login_cfg)
                        except Exception as _e:
                            log(f"      (re-login ไม่สำเร็จ: {_e} — ใช้กลไกฟื้นอัตโนมัติตามปกติ)")

                # รอบ retry: ลองคำขอที่ยังไม่สำเร็จซ้ำได้ถึง MAX_RETRY_PASSES รอบ
                # (login ใหม่ก่อนแต่ละรอบกัน session หมด) — resume ในตัวจะข้ามไฟล์ที่โหลดแล้ว
                for rpass in range(1, MAX_RETRY_PASSES + 1):
                    if not retry_pending or (is_cancelled and is_cancelled()):
                        break
                    log(f"      🔁 retry รอบ {rpass}/{MAX_RETRY_PASSES}: ลองคำขอที่ยังไม่สำเร็จ {len(retry_pending)} รายการ")
                    try:
                        login(page, login_cfg)
                    except Exception as _e:
                        log(f"      (login retry รอบ {rpass} ไม่สำเร็จ: {_e})")
                    still_pending: list[tuple[int, dict, dict | None]] = []
                    for res_idx, rec, prebuilt in retry_pending:
                        if is_cancelled and is_cancelled():
                            log("[!] ผู้ใช้ยกเลิก — หยุด"); break
                        res2 = _process_one_result_by_ref(
                            page, rec, login_cfg, docs_dir, doc_keys, log=log,
                            prebuilt_row=prebuilt,
                        )
                        results[res_idx] = res2
                        if res2.get("status") != "SUCCESS":
                            still_pending.append((res_idx, rec, prebuilt))
                        _save_result_docs_report(results, out_path, doc_keys, log=lambda *_: None)
                    retry_pending = still_pending
                if retry_pending:
                    log(f"      ⚠ หลัง retry {MAX_RETRY_PASSES} รอบ ยังเหลือ {len(retry_pending)} รายการที่ไม่สำเร็จ — บันทึกใน report")
        finally:
            ctx.close(); browser.close()

    _save_result_docs_report(results, out_path, doc_keys, log=log)
    success = sum(1 for r in results if r.get("status") in ("SUCCESS", "PARTIAL"))
    fail = sum(1 for r in results if r.get("status") not in ("SUCCESS", "PARTIAL"))
    log(f"[3/3] สรุป: สำเร็จ {success} / {len(results)} (ล้มเหลว/ไม่พบ {fail})")
    return success, out_path


def run_result_docs(
    cfg: dict,
    login_excel: Path,
    out_path: Path,
    request_type: str = "",
    doc_keys: list[str] | None = None,
    log=print,
    progress=None,
    is_cancelled=None,
) -> tuple[int, Path]:
    """ดาวน์โหลด 'ใบแจ้งผลใบอนุญาตทำงาน' / 'ใบรับคำขอใบอนุญาตทำงาน' จากแท็บเอกสารตอบรับ
    - ไม่ใช้ RequestData — ดึงคำขอจาก e-Tracking (filter สถานะรอนัดหมาย + รายการคำขอที่เลือก)
    - login จาก UsernameLogin.xlsx (วนทุกบัญชี)
    - ตั้งชื่อไฟล์ {ชื่อคนต่างด้าว(Eng)}_{ชื่อเอกสารบนเว็บ}.pdf
    """
    doc_keys = [d for d in (doc_keys or RESULT_DOC_TYPES_DEFAULT) if d in RESULT_DOC_TYPES]
    if not doc_keys:
        raise ValueError("ต้องเลือกประเภทเอกสารอย่างน้อย 1 อย่าง")
    out_path = _timestamped_path(out_path)
    accounts = _read_login_accounts(login_excel)
    docs_dir = out_path.parent / "result_docs"
    docs_dir.mkdir(parents=True, exist_ok=True)
    # สถานะ AP = รอนัดหมาย (รวมทั้ง 'รอทำการนัดหมาย' และ 'นัดหมายแล้ว')
    # สถานะที่ดึง (default = AP+SS): AP = รอนัดหมาย, SS = ดำเนินการเสร็จสิ้น
    # — override ได้ผ่าน cfg['result_status_ids'] หรือ CLI --result-status-ids
    status_ids = list(cfg.get("result_status_ids") or ["AP", "SS"])

    # checkpoint ระดับคำขอ — รันใหม่จะข้ามคำขอที่เสร็จแล้วทันที (ไม่เปิด detail ซ้ำ)
    prog_path = _progress_path(out_path)
    done_recs = _load_keyed_progress(prog_path, log=log)
    done_keys = set(done_recs.keys())

    log(f"[1/3] บัญชี login: {login_excel.name} ({len(accounts)} บัญชี)")
    log(f"      รายการคำขอ: {request_type or 'ทั้งหมด'} | สถานะ: {','.join(status_ids)}")
    log(f"      เอกสารที่จะดาวน์โหลด: {', '.join(RESULT_DOC_TYPES[d]['label'] for d in doc_keys)}")

    # นำผลที่เคยทำเสร็จแล้วใส่รายงานด้วย (ให้รายงานครบ)
    results: list[dict[str, Any]] = [done_recs[k] for k in done_recs]
    success = sum(1 for r in results if r.get("status") == "SUCCESS")
    done_count = len(results)
    total_known = 0

    if progress:
        try: progress(0, 1)
        except Exception: pass

    headless = bool(cfg.get("headless", False))
    with sync_playwright() as pw:
        browser = _launch_chromium(pw, cfg, ["--ignore-certificate-errors", "--start-maximized"])
        ctx = browser.new_context(
            locale="th-TH", ignore_https_errors=True,
            viewport={"width": 1920, "height": 1080}, accept_downloads=True,
        )
        page = ctx.new_page()
        try:
            for gi, (ukey, acct) in enumerate(accounts.items(), start=1):
                if is_cancelled and is_cancelled():
                    log("[!] ผู้ใช้ยกเลิก — หยุด"); break
                login_cfg = {
                    "username": acct["username"],
                    "password": acct["password"],
                    "user_type": acct["type"],
                    "method": acct.get("method") or cfg.get("method", "E-Workpermit"),
                }
                log(f"[บัญชี {gi}/{len(accounts)}] เข้าสู่ระบบ: {acct['username']} ({acct['type']})")
                try:
                    if gi > 1:
                        _logout_safely(page)
                        page.wait_for_timeout(800)
                    login(page, login_cfg)
                    goto_tracking(page)
                    apply_wa_filter(page, request_type or "", status_ids=status_ids)
                except Exception as e:
                    log(f"      ✗ login/เปิด tracking ไม่สำเร็จ: {e} — ข้ามบัญชีนี้")
                    continue

                try:
                    rows = collect_all_wa_rows(page, log=log)
                except Exception as e:
                    log(f"      ✗ เก็บรายการคำขอไม่สำเร็จ: {e} — ข้ามบัญชีนี้")
                    continue
                log(f"      พบ {len(rows)} คำขอ")
                total_known += len(rows)
                RELOGIN_EVERY = 80  # re-login เชิงรุกทุก N รายการ กัน session timeout
                MAX_RETRY_PASSES = 3  # จำนวนรอบ retry คำขอที่ยังไม่สำเร็จ (หลังรอบแรก)
                ri_done = 0  # นับเฉพาะรายการที่ทำจริง (ไม่นับที่ resume/skip)

                def _process_row(row: dict, seq_label: str) -> dict[str, Any]:
                    """เปิด detail + ดาวน์โหลดใบแจ้งผล/ใบรับคำขอของ 1 แถว → คืน rec
                    ถ้าดึงชื่อคนต่างด้าวไม่ได้ → คืน status=NO_NAME โดยไม่ดาวน์โหลด
                    (กันได้ไฟล์ชื่อผิด/หาย — ปล่อยให้รอบสองมาทำใหม่)
                    """
                    req_no = row.get("reqNo", "")
                    rec: dict[str, Any] = {
                        "seq": seq_label, "req_no": req_no,
                        "username": acct["username"], "name_eng": "",
                        "employer_th": "", "subdir": "",
                        "status_text": row.get("statusText", ""),
                        "docs": {dk: {"status": "", "file": "", "error": ""} for dk in doc_keys},
                        "status": "", "error": "",
                    }
                    log(f"   • [{seq_label}] คำขอ {req_no} ({row.get('statusText','')})")

                    # เปิด detail โดย navigate ตรงผ่าน URL (เร็วกว่าคลิกตาราง+go_back)
                    # ถ้าหลุด session → login ใหม่ด้วยบัญชีเดิม แล้วลองอีกครั้ง
                    opened = _open_detail_direct(page, row)
                    if not opened and _is_logged_out(page):
                        log("     ⚠ session หมดอายุ — login ใหม่แล้วลองอีกครั้ง")
                        try:
                            login(page, login_cfg)
                            opened = _open_detail_direct(page, row)
                        except Exception as _e:
                            log(f"     ✗ login ใหม่ไม่สำเร็จ: {_e}")
                    if not opened:
                        rec["status"] = "FAIL"; rec["error"] = "เปิดหน้า detail ไม่สำเร็จ"
                        return rec

                    name_eng = _extract_alien_eng_name(page)
                    if not name_eng:
                        # ดึงชื่อไม่ได้ → ใช้ req_no เป็นชื่อ (ยังคงดาวน์โหลดต่อ ไม่ skip)
                        log(f"     ⚠ {req_no}: ดึงชื่อคนต่างด้าวไม่ได้ — ใช้เลขคำขอเป็นชื่อไฟล์แทน")
                    rec["name_eng"] = name_eng
                    name_safe = (_receipt_safe_name(name_eng) or _receipt_safe_name(req_no)).strip(" -_.").strip()

                    # ดึงชื่อบริษัท (ไทย) → ใช้ตั้ง subfolder
                    employer_th = _extract_employer_name_th(page)
                    rec["employer_th"] = employer_th
                    subdir = _company_folder_name(employer_th) or "_no_company"
                    rec["subdir"] = subdir
                    req_docs_dir = docs_dir / subdir
                    req_docs_dir.mkdir(parents=True, exist_ok=True)
                    if employer_th:
                        log(f"     🏢 บริษัท: {employer_th} → {subdir}/")
                    else:
                        log(f"     ⚠ ดึงชื่อบริษัทไม่ได้ → เก็บใน {subdir}/")

                    for dk in doc_keys:
                        r = _download_response_doc_named(page, req_docs_dir, name_safe, RESULT_DOC_TYPES[dk], log=log, req_no=req_no)
                        rec["docs"][dk] = {"status": r["status"], "file": r["file"], "error": r["error"]}

                    # session หมดกลางทาง (เช่น modal Session Timeout) → login ใหม่ + เปิด detail
                    # ใหม่ + ดาวน์โหลดอีกครั้ง
                    if all(rec["docs"][dk]["status"] not in ("SUCCESS", "SKIP_EXISTS") for dk in doc_keys) \
                            and _is_logged_out(page):
                        try:
                            log("     ⚠ session หมดอายุระหว่างดาวน์โหลด — login ใหม่แล้วลองอีกครั้ง")
                            login(page, login_cfg)
                            if _open_detail_direct(page, row):
                                name_eng2 = _extract_alien_eng_name(page)
                                if name_eng2:
                                    rec["name_eng"] = name_eng2
                                    name_safe = (_receipt_safe_name(name_eng2) or _receipt_safe_name(req_no)).strip(" -_.").strip()
                                    employer_th2 = _extract_employer_name_th(page)
                                    if employer_th2:
                                        rec["employer_th"] = employer_th2
                                        subdir = _company_folder_name(employer_th2) or "_no_company"
                                        rec["subdir"] = subdir
                                        req_docs_dir = docs_dir / subdir
                                        req_docs_dir.mkdir(parents=True, exist_ok=True)
                                    for dk in doc_keys:
                                        r = _download_response_doc_named(page, req_docs_dir, name_safe, RESULT_DOC_TYPES[dk], log=log, req_no=req_no)
                                        rec["docs"][dk] = {"status": r["status"], "file": r["file"], "error": r["error"]}
                        except Exception as _e:
                            log(f"     ✗ ฟื้น session ไม่สำเร็จ: {_e}")

                    statuses = [rec["docs"][dk]["status"] for dk in doc_keys]
                    if all(s in ("SUCCESS", "SKIP_EXISTS") for s in statuses):
                        rec["status"] = "SUCCESS"
                    elif any(s == "SUCCESS" for s in statuses):
                        rec["status"] = "PARTIAL"
                    else:
                        rec["status"] = "FAIL"
                    rec["error"] = " | ".join(
                        f"{RESULT_DOC_TYPES[dk]['label']}: {rec['docs'][dk]['error']}"
                        for dk in doc_keys if rec["docs"][dk].get("error")
                    )
                    return rec

                retry_pending: list[tuple[int, dict, str]] = []  # (res_idx, row, ckey)
                for ri, row in enumerate(rows, start=1):
                    if is_cancelled and is_cancelled():
                        log("[!] ผู้ใช้ยกเลิก — หยุด"); break
                    req_no = row.get("reqNo", "")
                    ckey = f"{acct['username']}|{req_no}"
                    # ข้ามคำขอที่เคยทำเสร็จแล้ว (resume) — ไม่ต้องเปิด detail ซ้ำ
                    if ckey in done_keys:
                        log(f"   • [{gi}.{ri}] คำขอ {req_no} — ทำเสร็จแล้ว ข้าม (resume)")
                        if progress:
                            try: progress(done_count, max(total_known, done_count))
                            except Exception: pass
                        continue

                    rec = _process_row(row, f"{gi}.{ri}")
                    results.append(rec)
                    res_idx = len(results) - 1
                    done_count += 1
                    # บันทึก checkpoint เฉพาะคำขอที่เสร็จครบ (ครั้งหน้าจะข้าม) —
                    # ถ้ายังไม่ครบ (PARTIAL/FAIL/NO_NAME) จะเก็บไว้ retry และไม่บันทึก checkpoint
                    if rec["status"] == "SUCCESS":
                        done_keys.add(ckey)
                        _append_progress(prog_path, {**rec, "ckey": ckey})
                    else:
                        retry_pending.append((res_idx, row, ckey))
                    if progress:
                        try: progress(done_count, max(total_known, done_count))
                        except Exception: pass

                    # หมายเหตุ: ไม่ต้อง go_back — รายการถัดไป navigate เข้า detail ตรงผ่าน URL

                    # re-login เชิงรุกกัน session timeout หลังดาวน์โหลดไปสักระยะ
                    ri_done += 1
                    if RELOGIN_EVERY and ri_done % RELOGIN_EVERY == 0 and ri < len(rows):
                        log(f"      🔄 ครบ {ri_done} รายการ — login ใหม่เชิงรุกกัน session timeout...")
                        try:
                            login(page, login_cfg)
                        except Exception as _e:
                            log(f"      (re-login ไม่สำเร็จ: {_e} — ใช้กลไกฟื้นอัตโนมัติตามปกติ)")

                # รอบ retry: ลองคำขอที่ยังไม่สำเร็จซ้ำได้ถึง MAX_RETRY_PASSES รอบ
                # (login ใหม่ก่อนแต่ละรอบกัน session หมด) — resume ในตัวจะข้ามไฟล์ที่โหลดแล้ว
                for rpass in range(1, MAX_RETRY_PASSES + 1):
                    if not retry_pending or (is_cancelled and is_cancelled()):
                        break
                    log(f"      🔁 retry รอบ {rpass}/{MAX_RETRY_PASSES}: ลองคำขอที่ยังไม่สำเร็จ {len(retry_pending)} รายการ")
                    try:
                        login(page, login_cfg)
                    except Exception as _e:
                        log(f"      (login retry รอบ {rpass} ไม่สำเร็จ: {_e})")
                    still_pending: list[tuple[int, dict, str]] = []
                    for res_idx, row, ckey in retry_pending:
                        if is_cancelled and is_cancelled():
                            log("[!] ผู้ใช้ยกเลิก — หยุด"); break
                        rec2 = _process_row(row, results[res_idx].get("seq", ""))
                        results[res_idx] = rec2
                        if rec2["status"] == "SUCCESS":
                            done_keys.add(ckey)
                            _append_progress(prog_path, {**rec2, "ckey": ckey})
                        else:
                            still_pending.append((res_idx, row, ckey))
                        _save_result_docs_report(results, out_path, doc_keys, log=lambda *_: None)
                    retry_pending = still_pending
                if retry_pending:
                    log(f"      ⚠ หลัง retry {MAX_RETRY_PASSES} รอบ ยังเหลือ {len(retry_pending)} รายการที่ไม่สำเร็จ — บันทึกใน report")

                # เซฟ report เป็นระยะ
                _save_result_docs_report(results, out_path, doc_keys, log=lambda *_: None)
        finally:
            ctx.close(); browser.close()

    _save_result_docs_report(results, out_path, doc_keys, log=log)
    success = sum(1 for r in results if r.get("status") in ("SUCCESS", "PARTIAL"))
    fail = sum(1 for r in results if r.get("status") not in ("SUCCESS", "PARTIAL"))
    log(f"[3/3] สรุป: สำเร็จ {success} / {len(results)} (ล้มเหลว/บางส่วน {fail})")
    return success, out_path


# ─────────────────────────────────────────────────────────────────
# โหมด INFORM_ENTER_EXIT — แจ้งการจ้างคนต่างด้าวเข้าทำงาน (แบบ บต.52)
#   อ่าน FormRequestEmployment.xlsx (ประเภทนายจ้าง, รหัสนายจ้าง, เลขใบอนุญาตทำงาน, วันที่จ้าง, Files)
#   - login บัญชีจาก UsernameLogin.xlsx (ตรงกับฟอร์มที่ผู้ใช้ login)
#   - เปิด URL form INFORM_ENTER_EXIT → ค้นหานายจ้าง → เลือกสาขา → ติ๊กเลขใบอนุญาตทำงาน
#   - แนบเอกสารผู้รับมอบอำนาจ → เลือกวันที่จ้าง → ติ๊กยืนยัน → ถัดไป → จบ
#   safety: default จะหยุดก่อนกด "ยืนยันส่งจริง" (ข้อ 7+8)
#           ต้องส่ง commit=True เพื่อทำข้อ 7+8 ต่อ
# ─────────────────────────────────────────────────────────────────
INFORM_ENTER_EXIT_URL = (
    "https://eworkpermit.doe.go.th/WorkPermit"
    "?user_type=emp&ft=INFORM_ENTER_EXIT&ut=emp&uti=3"
)


def _read_inform_data(path: Path) -> list[dict[str, Any]]:
    """อ่าน FormRequestEmployment.xlsx → [{emp_type, emp_id, work_permit, hire_date, files, row_index}]
    คอลัมน์ (header แถวแรก):
      ประเภทนายจ้าง, รหัสนายจ้าง/เลขบัตรประจำตัวประชาชน/เลขที่นิติบุคคล,
      เลขใบอนุญาตทำงาน, วันที่เริ่มงาน, Files
    """
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    hdr = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    def col(*names: str) -> int:
        for nm in names:
            for i, h in enumerate(hdr):
                if h == nm:
                    return i
        for nm in names:
            for i, h in enumerate(hdr):
                if nm in h:
                    return i
        return -1

    i_etype = col("ประเภทนายจ้าง", "EmployerType", "Employer Type")
    i_eid = col("รหัสนายจ้าง", "เลขนิติบุคคล", "EmployerID", "Tax")
    i_wp = col("เลขใบอนุญาตทำงาน", "WorkPermit", "Permit")
    i_date = col("วันที่เริ่มงาน", "วันที่จ้าง", "HireDate", "StartDate")
    i_files = col("Files", "File", "เอกสาร")
    i_login = col("ผู้ใช้งานระบบ", "ผู้ใช้งาน", "ผู้ใช้", "บัญชีล็อกอิน", "บัญชี",
                  "เลขบัตรผู้กระทำการแทน", "Username", "UserLogin", "Login", "User")
    i_users = col("Users", "User", "Username", "ผู้ใช้งาน", "บัญชี")

    def fmt_date(v: Any) -> str:
        if v is None or v == "":
            return ""
        if isinstance(v, datetime):
            return v.strftime("%d/%m/%Y")
        if isinstance(v, date):
            return v.strftime("%d/%m/%Y")
        return str(v).strip()

    rows: list[dict[str, Any]] = []
    for ridx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not any(v not in (None, "") for v in row):
            continue

        def g(i: int) -> Any:
            return row[i] if 0 <= i < len(row) else None

        emp_type = (str(g(i_etype)).strip() if g(i_etype) is not None else "")
        emp_id = (str(g(i_eid)).strip() if g(i_eid) is not None else "")
        login_raw = g(i_login)
        if isinstance(login_raw, float) and login_raw.is_integer():
            login_raw = int(login_raw)
        login_user = str(login_raw).strip() if login_raw is not None else ""
        wp_raw = g(i_wp)
        # work-permit อาจมาเป็นเลข int → แปลงให้คงเป็น digit อย่างเดียว
        if isinstance(wp_raw, float) and wp_raw.is_integer():
            wp_raw = int(wp_raw)
        wp = str(wp_raw).strip() if wp_raw is not None else ""
        # เลขใบอนุญาตทำงานเป็น 13 หลัก — เติม 0 นำหน้าที่หายไปจาก Excel (เซลล์ตัวเลข)
        if wp.isdigit() and 0 < len(wp) < 13:
            wp = wp.zfill(13)
        files_raw = (str(g(i_files)).strip() if g(i_files) is not None else "")
        files = [s.strip() for s in re.split(r"[;|]", files_raw) if s.strip()]
        # คอลัมน์ Users = บัญชี login ที่ใช้ยื่นแทนนายจ้างรายนี้
        u_raw = g(i_users)
        if isinstance(u_raw, float) and u_raw.is_integer():
            u_raw = int(u_raw)
        username = str(u_raw).strip() if u_raw is not None else ""
        rec = {
            "emp_type": emp_type,
            "emp_id": emp_id,
            "login": login_user,
            "work_permit": wp,
            "hire_date": fmt_date(g(i_date)),
            "files": files,
            "username": username,
            "row_index": ridx,
        }
        if emp_id and wp:
            rows.append(rec)
    return rows


def _group_inform_records(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """รวม record ที่เป็นนายจ้างเดียวกัน + บัญชี login เดียวกัน (username + emp_type + emp_id)
    เข้าเป็นกลุ่มเดียว → 1 กลุ่ม = 1 การยื่น (ติ๊กแรงงานได้หลายคนในครั้งเดียว)
    คืนค่า [{username, emp_type, emp_id, files, workers:[{work_permit, hire_date, row_index}]}]
    เรียงตามลำดับที่พบครั้งแรกใน Excel
    """
    groups: dict[tuple[str, str, str], dict[str, Any]] = {}
    order: list[tuple[str, str, str]] = []
    for r in records:
        key = (r.get("username", ""), r.get("emp_type", ""), r.get("emp_id", ""))
        g = groups.get(key)
        if g is None:
            g = {
                "username": r.get("username", ""),
                "login": r.get("login", ""),
                "emp_type": r.get("emp_type", ""),
                "emp_id": r.get("emp_id", ""),
                "files": [],
                "workers": [],
                "row_indices": [],
            }
            groups[key] = g
            order.append(key)
        g["workers"].append({
            "work_permit": r.get("work_permit", ""),
            "hire_date": r.get("hire_date", ""),
            "row_index": r.get("row_index"),
        })
        if r.get("row_index") is not None:
            g["row_indices"].append(r["row_index"])
        for f in r.get("files", []):
            if f not in g["files"]:
                g["files"].append(f)
    return [groups[k] for k in order]



def _save_inform_report(rows: list[dict[str, Any]], out_path: Path, log=print) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "ผลแจ้งเข้านายจ้าง"
    headers = [
        "ลำดับ", "บัญชี (Users)", "ประเภทนายจ้าง", "รหัสนายจ้าง", "เลขใบอนุญาตทำงาน",
        "วันที่จ้าง", "เลขอ้างอิงคำขอ", "สถานะ", "หมายเหตุ", "Screenshot",
    ]
    ws.append(headers)
    head_fill = PatternFill("solid", fgColor="305496")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    text_cols = {"บัญชี (Users)", "รหัสนายจ้าง", "เลขใบอนุญาตทำงาน", "เลขอ้างอิงคำขอ"}
    link_font = Font(color="0563C1", underline="single")
    for r in rows:
        ws.append([
            r.get("seq", ""), r.get("username", ""), r.get("emp_type", ""), r.get("emp_id", ""),
            r.get("work_permit", ""), r.get("hire_date", ""),
            r.get("ref_no", ""), r.get("status", ""), r.get("note", ""),
            r.get("screenshot", ""),
        ])
        ridx = ws.max_row
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=ridx, column=c_idx)
            if h in text_cols and cell.value not in (None, ""):
                cell.value = str(cell.value); cell.number_format = "@"
            if h == "Screenshot" and cell.value:
                fname = str(cell.value)
                cell.value = f'=HYPERLINK("inform_screenshots/{fname}","{fname}")'
                cell.font = link_font

    widths = [8, 18, 36, 22, 22, 14, 22, 14, 50, 36]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    log(f"บันทึกรายงาน: {out_path}")


def _open_inform_form(page: Page, log=print) -> bool:
    """เปิดหน้าฟอร์ม INFORM_ENTER_EXIT (แบบ บต.52) ผ่านเมนูบริการ
    flow: หน้าหลัก → คลิก 'เมนูบริการ' (.lang_menu_service)
          → คลิก '#MT_13_1_INFORM' (การแจ้งการจ้างคนต่างด้าวเข้าทำงานกับนายจ้าง บต.52)
    เปิดผ่านเมนูจำเป็นเพราะ select 'ประเภทนายจ้าง' จะโหลด options เมื่อเข้าผ่าน flow ปกติเท่านั้น
    """
    try:
        page.goto("https://eworkpermit.doe.go.th/", wait_until="domcontentloaded", timeout=30_000)
        page.wait_for_timeout(2000)
    except Exception as e:
        log(f"      ✗ เปิดหน้าหลักไม่สำเร็จ: {e}")
        return False

    # 1) คลิก 'เมนูบริการ'
    clicked = page.evaluate(r"""() => {
        const a = document.querySelector('a.lang_menu_service');
        if (a) { a.click(); return true; }
        const fb = Array.from(document.querySelectorAll('a,button'))
            .find(x => /เมนูบริการ/.test((x.textContent||'').trim()));
        if (fb) { fb.click(); return true; }
        return false;
    }""")
    if not clicked:
        log("      ✗ ไม่พบลิงก์ 'เมนูบริการ'")
        return False
    page.wait_for_timeout(1500)

    # 2) เปิด tab INFORM_ENTER_EXIT (กดปุ่มหมวดด้านซ้าย) แล้วคลิก #MT_13_1_INFORM
    ok = page.evaluate(r"""() => {
        // เปิดหมวด tab_INFORM_ENTER_EXIT ก่อน
        try {
            if (typeof openCity === 'function') {
                openCity('tab_INFORM_ENTER_EXIT', new Event('click'));
            }
        } catch(e) {}
        // คลิกการ์ด/ลิงก์ #MT_13_1_INFORM (บต.52)
        const target = document.querySelector('#MT_13_1_INFORM');
        if (target) { target.click(); return true; }
        return false;
    }""")
    if not ok:
        log("      ✗ ไม่พบเมนู #MT_13_1_INFORM (บต.52)")
        return False
    # บางครั้ง onclick เรียก setFormTypeRenew เฉยๆ ไม่ได้นำทางทันที — รอ + ลองคลิกซ้ำ
    page.wait_for_timeout(2000)
    if "INFORM_ENTER_EXIT" not in (page.url or ""):
        # ลองคลิกอีกครั้ง (บางครั้งคลิกครั้งแรกแค่ highlight)
        try:
            page.evaluate("() => { const t = document.querySelector('#MT_13_1_INFORM'); if (t) t.click(); }")
            page.wait_for_timeout(2000)
        except Exception:
            pass
    # ถ้ายังไม่ไป → fallback ใช้ URL ตรง (อาจไม่ได้ผล แต่ลองดู)
    if "INFORM_ENTER_EXIT" not in (page.url or ""):
        log(f"      ⚠ ยังไม่เข้าฟอร์ม (URL={page.url}) — ลอง URL ตรงเป็น fallback")
        try:
            page.goto(INFORM_ENTER_EXIT_URL, wait_until="domcontentloaded", timeout=30_000)
        except Exception:
            pass
        page.wait_for_timeout(2500)

    # ยืนยันว่าฟอร์มพร้อม + ปุ่ม 'ค้นหานายจ้าง' พร้อม
    try:
        page.wait_for_function(
            r"""() => {
              const url = location.href || '';
              const inForm = /INFORM_ENTER_EXIT/i.test(url);
              const hasBtn = Array.from(document.querySelectorAll('.bus_type_modal'))
                .some(b => (b.textContent||'').includes('ค้นหานายจ้าง') && b.id !== 'addEmpBtn');
              return inForm && hasBtn;
            }""",
            timeout=20_000,
        )
        return True
    except PWTimeoutError:
        log(f"      ✗ เปิดฟอร์มไม่สำเร็จ (URL={page.url})")
        return False


def _inform_search_employer(page: Page, emp_type: str, emp_id: str, log=print) -> bool:
    """กด 'ค้นหานายจ้าง' → เลือกประเภท + กรอกเลข + ค้นหา + บันทึก"""
    page.evaluate(r"""() => {
        const btns = Array.from(document.querySelectorAll('.bus_type_modal'));
        const target = btns.find(b => (b.textContent||'').includes('ค้นหานายจ้าง') && b.id !== 'addEmpBtn');
        if (target) target.click();
    }""")
    page.wait_for_timeout(2000)
    # รอให้ select มี options โหลดเสร็จ
    try:
        page.wait_for_function(
            r"""() => {
              const s = document.querySelector('#slt_employer_type');
              return s && s.options && s.options.length > 1;
            }""",
            timeout=15_000,
        )
    except PWTimeoutError:
        log("      ✗ select ประเภทนายจ้าง โหลดช้า/ไม่ขึ้น")
        return False

    matched_value = page.evaluate(r"""(want) => {
        const s = document.querySelector('#slt_employer_type');
        if (!s) return '';
        const norm = t => (t||'').replace(/\s+/g,' ').trim().toLowerCase();
        const w = norm(want);
        const opts = Array.from(s.options);
        let opt = opts.find(o => norm(o.textContent) === w);
        if (!opt) opt = opts.find(o => w && norm(o.textContent).includes(w));
        if (!opt) return '';
        s.value = opt.value;
        s.dispatchEvent(new Event('change', {bubbles:true}));
        if (window.jQuery) try { jQuery(s).trigger('change'); } catch(e){}
        return opt.value + '||' + (opt.textContent||'').trim();
    }""", emp_type)
    if not matched_value:
        log(f"      ✗ ไม่พบประเภทนายจ้าง '{emp_type}' ใน dropdown")
        return False
    log(f"      เลือกประเภทนายจ้าง: {matched_value.split('||',1)[-1]}")
    page.wait_for_timeout(800)

    # กรอกเลข + ค้นหา
    page.evaluate(r"""(v) => {
        const t = document.querySelector('#employer_id_search');
        if (t) {
            t.value = v;
            t.dispatchEvent(new Event('input', {bubbles:true}));
            t.dispatchEvent(new Event('change', {bubbles:true}));
        }
    }""", emp_id)
    page.evaluate(r"""() => {
        const b = document.querySelector('.searchEmployer');
        if (b) b.click();
    }""")
    page.wait_for_timeout(2500)
    # รอให้ผลค้นหาแสดง (มี รหัสนายจ้าง / RE...) — ถ้าไม่มี ให้ฟ้อง
    try:
        page.wait_for_function(
            r"""() => {
              const txt = document.body.innerText || '';
              return /RE\d{10,}/.test(txt);
            }""",
            timeout=12_000,
        )
    except PWTimeoutError:
        log("      ✗ ค้นหานายจ้างไม่พบผล")
        return False

    # กดบันทึก (#btn_search_emp_submit)
    page.evaluate(r"""() => {
        const b = document.querySelector('#btn_search_emp_submit');
        if (b) b.click();
    }""")
    # หลังบันทึกอาจมี popup 'ค้นหานายจ้างสำเร็จ' เด้งขึ้น — พยายามปิดแบบ best-effort
    # (ไม่ block: ถ้าปิดไม่ได้ก็เดินหน้าขั้นต่อไป)
    page.wait_for_timeout(1500)
    for _ in range(4):
        _inform_dismiss_popup(page, log=log)
        page.wait_for_timeout(600)
        still = page.evaluate(r"""() => {
            const m = document.querySelector('.swal2-popup, .modal.show');
            return !!(m && m.offsetParent !== null);
        }""")
        if not still:
            break
    page.wait_for_timeout(500)
    return True


def _inform_dismiss_popup(page: Page, log=print) -> bool:
    """ปิด modal/popup ที่เด้งขึ้น (success/แจ้งเตือน) — รองรับ SweetAlert + bootstrap modal
    กดปุ่ม 'ปิด/ตกลง/OK' ตัวที่มองเห็นได้จริงบนหน้า (ไม่ผูกกับ scope ใด scope หนึ่ง
    เพราะหน้านี้มี dialog template ซ่อนอยู่หลายตัว)
    """
    closed = page.evaluate(r"""() => {
        let did = false;
        const isVis = el => {
            if (!el) return false;
            if (el.offsetParent === null) return false;
            const r = el.getBoundingClientRect();
            return r.width > 0 && r.height > 0;
        };
        // 1) SweetAlert2
        const sw = document.querySelector('.swal2-confirm, .swal2-close');
        if (isVis(sw)) { sw.click(); did = true; }
        // 2) ปุ่ม/ลิงก์ที่มองเห็นได้จริง ข้อความ ปิด/ตกลง/OK/รับทราบ ทั้งหน้า
        if (!did) {
            const cands = Array.from(document.querySelectorAll('button, a.btn, .btn, a'))
                .filter(isVis)
                .filter(x => /^(ปิด|ตกลง|OK|รับทราบ|ปิดหน้าต่าง|ตกลง|Close)$/i.test((x.textContent||'').trim()));
            // กดตัวสุดท้าย (popup มักถูก append ท้ายสุด)
            const b = cands[cands.length - 1];
            if (b) { b.click(); did = true; }
        }
        // 3) ปุ่ม close (×) ของ modal ที่แสดงอยู่
        if (!did) {
            const xs = Array.from(document.querySelectorAll('.modal.show .close, .modal.show [data-dismiss="modal"], .swal2-close, .close')).filter(isVis);
            if (xs.length) { xs[xs.length - 1].click(); did = true; }
        }
        return did;
    }""")
    if closed:
        log("      ✓ ปิด popup หลังบันทึก")
    return bool(closed)


def _inform_consent_first(page: Page, log=print) -> bool:
    """ติ๊ก checkbox ยืนยันรับรองในขั้นที่ 1 (ก่อน 'ถัดไป')"""
    ok = page.evaluate(r"""() => {
        const cbs = Array.from(document.querySelectorAll('input[type=checkbox]'));
        const textFor = cb => {
            // เก็บข้อความจาก label[for], label ที่ครอบ, และ ancestor สูงสุด 5 ชั้น
            let parts = [];
            const lbl = cb.closest('label') || (cb.id ? document.querySelector(`label[for="${cb.id}"]`) : null);
            if (lbl) parts.push(lbl.textContent || '');
            let p = cb.parentElement, lvl = 0;
            while (p && lvl < 5) { parts.push(p.textContent || ''); p = p.parentElement; lvl++; }
            return parts.join(' ').replace(/\s+/g,' ');
        };
        for (const cb of cbs) {
            if (cb.offsetParent === null) continue;  // ต้องมองเห็นได้
            const txt = textFor(cb);
            if (/ขอรับรองว่า/.test(txt) ||
                /มีความประสงค์ในการยื่นคำขอ/.test(txt) ||
                /ได้รับความยินยอมจากบุคคลดังกล่าว/.test(txt)) {
                if (!cb.checked) {
                    cb.click();  // ให้ handler เดิมทำงาน
                    if (!cb.checked) {  // ถ้ายังไม่ติ๊ก บังคับ set + dispatch
                        cb.checked = true;
                        cb.dispatchEvent(new Event('change', {bubbles:true}));
                        cb.dispatchEvent(new Event('click', {bubbles:true}));
                    }
                }
                return true;
            }
        }
        return false;
    }""")
    if not ok:
        log("      ⚠ ไม่พบ checkbox รับรอง (ขั้นที่ 1) — ลองเดินหน้าต่อ")
    return bool(ok)


def _inform_click_next_step1(page: Page, log=print) -> bool:
    """กดปุ่ม 'ถัดไป' ของขั้นที่ 1 (#gonextSubmit) แล้วรอเข้าหน้าถัดไป"""
    cur_url = page.url
    clicked = page.evaluate(r"""() => {
        const b = document.querySelector('#gonextSubmit:not(.d-none)') || document.querySelector('#gonextSubmit');
        if (b) { b.click(); return true; }
        return false;
    }""")
    if not clicked:
        log("      ✗ ไม่พบปุ่ม 'ถัดไป' (#gonextSubmit)")
        return False
    page.wait_for_timeout(2500)
    # อาจขึ้น modal alert ก่อน — กดยืนยัน
    try:
        page.evaluate(r"""() => {
            const btn = Array.from(document.querySelectorAll('.swal2-confirm,.btn-primary'))
                .find(b => /ยืนยัน|ตกลง|ถัดไป/.test((b.textContent||'')));
            if (btn) btn.click();
        }""")
    except Exception:
        pass
    page.wait_for_timeout(1500)
    return True


def _inform_collect_branches(page: Page) -> list[dict[str, str]]:
    """อ่านรายการสาขาจาก dropdown '#workplace-options' (สถานที่ทำงาน/สาขา) → [{value, text}]"""
    return page.evaluate(r"""() => {
        let s = document.querySelector('#workplace-options');
        if (!s) {
            // fallback: หา select ที่ option มีคำว่า 'สำนักงาน'/'สาขา'/'ที่อยู่ที่ทำงาน'
            const sels = Array.from(document.querySelectorAll('select'));
            s = sels.find(x => Array.from(x.options).some(o => /สำนักงาน|สาขา|ที่อยู่ที่ทำงาน/.test(o.textContent||'')));
        }
        if (!s) return [];
        return Array.from(s.options).map(o => ({
            value: o.value, text: (o.textContent||'').replace(/\s+/g,' ').trim(),
        })).filter(o => o.value && !/^กรุณาเลือก/.test(o.text));
    }""")


def _inform_select_branch(page: Page, value: str, log=print) -> bool:
    """เลือกสาขาตาม value ใน '#workplace-options' แล้วกดปุ่ม 'เลือกสาขา' (#searchWorkplaces)"""
    ok = page.evaluate(r"""(v) => {
        let s = document.querySelector('#workplace-options');
        if (!s) {
            const sels = Array.from(document.querySelectorAll('select'));
            s = sels.find(x => Array.from(x.options).some(o => o.value === v));
        }
        if (!s) return false;
        s.value = v;
        s.dispatchEvent(new Event('change', {bubbles:true}));
        if (window.jQuery) try { jQuery(s).trigger('change'); } catch(e){}
        const b = document.querySelector('#searchWorkplaces') ||
                  Array.from(document.querySelectorAll('button,a.btn,.btn')).find(x => /เลือกสาขา/.test((x.textContent||'').trim()));
        if (b) b.click();
        return true;
    }""", value)
    if not ok:
        log(f"      ✗ เลือกสาขาไม่สำเร็จ (value={value})")
    page.wait_for_timeout(2500)
    return bool(ok)


def _inform_check_workpermits(page: Page, wp_numbers: list[str], log=print) -> tuple[int, list[str]]:
    """หาแถวที่ 'เลขที่ใบอนุญาตทำงาน' ตรงกับใน list แล้ว tick checkbox
    return (matched_count, unmatched_list)
    """
    res = page.evaluate(r"""(wants) => {
        const wantSet = new Set(wants.map(s => String(s).replace(/\s+/g,'')));
        const found = new Set();
        // ไล่ทุกตารางในหน้า เลือก row ที่ cell ใดมี 'เลขที่ใบอนุญาตทำงาน' ตรง
        const rows = Array.from(document.querySelectorAll('table tr'));
        for (const tr of rows) {
            const cells = Array.from(tr.querySelectorAll('td'));
            for (const td of cells) {
                const txt = (td.textContent||'').replace(/\s+/g,'');
                if (wantSet.has(txt)) {
                    const cb = tr.querySelector('input[type=checkbox]');
                    if (cb) {
                        if (!cb.checked) {
                            cb.click();  // ให้ handler เดิมทำงาน (อย่า set checked=true ก่อน click)
                            if (!cb.checked) {  // ถ้ายังไม่ติด บังคับ set + dispatch
                                cb.checked = true;
                                cb.dispatchEvent(new Event('change', {bubbles:true}));
                                cb.dispatchEvent(new Event('click', {bubbles:true}));
                            }
                        }
                        found.add(txt);
                    }
                    break;
                }
            }
        }
        return {found: Array.from(found), wants: Array.from(wantSet)};
    }""", wp_numbers)
    matched = res.get("found", [])
    unmatched = [w for w in wp_numbers if w.replace(" ", "") not in matched]
    return len(matched), unmatched


def _inform_set_hire_date(page: Page, wp: str, hire_date: str, log=print) -> bool:
    """หาแถวที่เลขใบอนุญาตทำงาน = wp → กดลิงก์ 'เพิ่มวันที่จ้าง' (a.alien-addDate)
    → กรอก hire_date (รูปแบบ dd/mm/yyyy) ใน #work_duration_date_edit → กดบันทึก → รอ panel ปิด
    ทำทีละคน (panel ต้องปิดก่อนจึงทำคนถัดไปได้)
    """
    want = str(wp).replace(" ", "")
    # 0) ให้แน่ใจว่าไม่มี panel วันที่ค้างเปิดอยู่ก่อน
    try:
        if page.locator("#work_duration_date_edit").is_visible():
            page.evaluate(r"""() => {
                const b = document.querySelector('#btnCloseAddDateAlien');
                if (b) b.click();
            }""")
            page.wait_for_timeout(600)
    except Exception:
        pass
    # 1) คลิกลิงก์ 'เพิ่มวันที่จ้าง' ของแถวที่ wp ตรง (เฉพาะแถวที่มีลิงก์ — ข้าม template)
    clicked = page.evaluate(r"""(want) => {
        const norm = s => String(s||'').replace(/\s+/g,'');
        const rows = Array.from(document.querySelectorAll('table tbody tr'));
        for (const tr of rows) {
            const cells = Array.from(tr.querySelectorAll('td'));
            const hit = cells.some(td => norm(td.textContent) === want);
            if (!hit) continue;
            const a = tr.querySelector('a.alien-addDate');
            if (a && a.offsetParent !== null) { a.click(); return true; }
            // เจอแถวแต่ไม่มีลิงก์ (template/ตั้งไปแล้ว) → วนหาแถวอื่นต่อ
        }
        return false;
    }""", want)
    if not clicked:
        log(f"      ⚠ ไม่พบลิงก์ 'เพิ่มวันที่จ้าง' ของ wp={wp} (อาจตั้งไปแล้ว)")
        return False
    # 2) รอ panel เปิด
    try:
        page.wait_for_selector("#work_duration_date_edit", state="visible", timeout=8000)
    except Exception:
        log(f"      ⚠ panel วันที่จ้างไม่เปิด (wp={wp})")
        return False
    page.wait_for_timeout(500)
    # 3) กรอกวันที่ลงใน #work_duration_date_edit (placeholder วว/ดด/ปปปป = dd/mm/yyyy)
    ok = page.evaluate(r"""(v) => {
        const t = document.getElementById('work_duration_date_edit');
        if (!t) return false;
        t.removeAttribute('readonly');
        t.value = v;
        t.dispatchEvent(new Event('input', {bubbles:true}));
        t.dispatchEvent(new Event('change', {bubbles:true}));
        t.dispatchEvent(new Event('blur', {bubbles:true}));
        if (window.jQuery) { try { jQuery(t).trigger('change'); } catch(e){} }
        return t.value;
    }""", hire_date)
    if not ok:
        log(f"      ⚠ ไม่พบช่องวันที่จ้าง (#work_duration_date_edit)")
        return False
    page.wait_for_timeout(500)
    # 4) กดบันทึกภายใน off-canvas เดียวกับช่องวันที่
    saved = page.evaluate(r"""() => {
        const isVis = el => el && el.offsetParent !== null;
        const inp = document.getElementById('work_duration_date_edit');
        const oc = inp ? inp.closest('.off-canvas, .offcanvas, [class*=canvas]') : null;
        const scope = oc || document;
        const b = Array.from(scope.querySelectorAll('button,a.btn')).filter(isVis)
            .find(x => /บันทึก/.test((x.textContent||'').trim()));
        if (b) { b.click(); return true; }
        return false;
    }""")
    if not saved:
        log(f"      ⚠ ไม่พบปุ่มบันทึกใน panel วันที่จ้าง (wp={wp})")
        return False
    # 5) รอ panel ปิด (ยืนยันว่าบันทึกแล้ว) ก่อนทำคนถัดไป
    try:
        page.wait_for_selector("#work_duration_date_edit", state="hidden", timeout=8000)
    except Exception:
        # เผื่อ panel ไม่ปิดเอง — กดปิด
        page.evaluate(r"""() => { const b=document.querySelector('#btnCloseAddDateAlien'); if(b) b.click(); }""")
        page.wait_for_timeout(600)
    page.wait_for_timeout(800)
    log(f"      ✓ ตั้งวันที่จ้าง {hire_date} (wp={wp})")
    return True




def _inform_attach_authorization(page: Page, file_paths: list[Path], log=print) -> bool:
    """หา input file ในส่วน 'เอกสาร : ผู้รับมอบอำนาจ' → set_input_files
    ช่องผู้รับมอบอำนาจมี id ขึ้นต้น 'btnChkProx' (ปุ่มเลือกไฟล์คือ 'btnSelectFileProx...')
    """
    if not file_paths:
        log("      ⚠ Excel ไม่ได้ระบุไฟล์แนบ — ข้าม")
        return False
    # หา id ของ input ช่องผู้รับมอบอำนาจ
    target_id = page.evaluate(r"""() => {
        const inputs = Array.from(document.querySelectorAll('input[type=file]'));
        // 1) id ขึ้นต้น btnChkProx (ช่องผู้รับมอบอำนาจ)
        let t = inputs.find(i => /^btnChkProx/i.test(i.id));
        if (t) return t.id;
        // 2) input ที่ ancestor (สูงสุด 10 ชั้น) มีคำว่า 'ผู้รับมอบอำนาจ'
        for (const i of inputs) {
            let p = i, lvl = 0;
            while (p && lvl < 10) {
                if (/ผู้รับมอบอำนาจ/.test(p.textContent||'')) return i.id || ('__idx_' + inputs.indexOf(i));
                p = p.parentElement; lvl++;
            }
        }
        return '';
    }""")
    if not target_id:
        log("      ✗ ไม่พบช่องแนบ 'ผู้รับมอบอำนาจ'")
        return False
    try:
        if target_id.startswith("__idx_"):
            idx = int(target_id.split("_")[-1])
            target = page.locator("input[type=file]").nth(idx)
        else:
            target = page.locator(f"#{target_id}")
        target.set_input_files([str(p) for p in file_paths])
        page.wait_for_timeout(2000)
        log(f"      ✓ แนบไฟล์ (ผู้รับมอบอำนาจ): {', '.join(p.name for p in file_paths)}")
        return True
    except Exception as e:
        log(f"      ✗ แนบไฟล์ไม่สำเร็จ: {e}")
        return False


def _inform_consent_summary(page: Page, log=print) -> bool:
    """ติ๊ก checkbox ยืนยันในหน้า 'สรุปคำขอ' (ข้อ 7) — มีคำว่า 'ตรวจสอบข้อมูล' หรือ 'ใบอนุญาตทำงาน'"""
    ok = page.evaluate(r"""() => {
        const cbs = Array.from(document.querySelectorAll('input[type=checkbox]'));
        for (const cb of cbs) {
            const lbl = cb.closest('label') || (cb.id ? document.querySelector(`label[for="${cb.id}"]`) : null);
            const txt = ((lbl && lbl.textContent) || cb.parentElement?.textContent || '').replace(/\s+/g,' ');
            if (/ข้าพเจ้าได้ตรวจสอบข้อมูล/.test(txt) ||
                /รับทราบว่าข้อมูลนี้จะถูกพิมพ์ในใบอนุญาต/.test(txt)) {
                if (!cb.checked) {
                    cb.click();
                    if (!cb.checked) {
                        cb.checked = true;
                        cb.dispatchEvent(new Event('change', {bubbles:true}));
                        cb.dispatchEvent(new Event('click', {bubbles:true}));
                    }
                }
                return true;
            }
        }
        return false;
    }""")
    return bool(ok)


def _inform_extract_ref_no(page: Page) -> str:
    """ดึงเลขอ้างอิงคำขอจากหน้า 'เสร็จสิ้น' (ข้อ 9)
    ดึงเฉพาะข้อความใน #step4 ก่อน (กันไปจับ 'เลขที่นิติบุคคล' ของนายจ้าง)"""
    try:
        return page.evaluate(r"""() => {
            const s4 = document.querySelector('#step4');
            // ใช้ข้อความใน step4 เป็นหลัก; ถ้าไม่มีค่อย fallback ทั้งหน้า
            const scope = (s4 && s4.offsetParent !== null) ? (s4.innerText||'') : (document.body.innerText||'');
            const txt = scope.replace(/\u00a0/g,' ');
            // 1) ต้องมีคำว่า คำขอ/อ้างอิง/รับเรื่อง จริง ๆ (ไม่เอา 'เลขที่นิติบุคคล')
            let m = txt.match(/(?:เลขที่คำขอ|เลขคำขอ|หมายเลขคำขอ|เลขที่อ้างอิง|เลขอ้างอิง|เลขที่รับเรื่อง|เลขรับเรื่อง|Ref(?:erence)?\s*No\.?)[\s:：]*([0-9][0-9\-\/]{8,20})/i);
            if (m) return m[1].replace(/[^\d]/g,'');
            // 2) เลขขึ้นต้น 69 ตามด้วย 12 หลัก (รูปแบบเลขคำขอที่พบ)
            m = txt.match(/\b69\d{12}\b/);
            if (m) return m[0];
            return '';
        }""") or ""


    except Exception:
        return ""


def _inform_process_one(
    page: Page,
    group: dict[str, Any],
    files_root: Path,
    screenshot_dir: Path,
    commit: bool,
    log=print,
) -> dict[str, Any]:
    """ทำ flow ข้อ 1-9 สำหรับ 1 นายจ้าง (ติ๊กแรงงานได้หลายคนในการยื่นครั้งเดียว)
    group = {emp_type, emp_id, files, workers:[{work_permit, hire_date, row_index}]}
    commit=False → หยุดก่อนข้อ 7 (ติ๊กยืนยัน + ส่ง) — เก็บ screenshot สรุปคำขอแทน
    """
    emp_type = group.get("emp_type", "")
    emp_id = group.get("emp_id", "")
    workers = group.get("workers", [])
    wp_targets = [w["work_permit"] for w in workers if w.get("work_permit")]
    wp_tag = wp_targets[0] if len(wp_targets) == 1 else f"{len(wp_targets)}คน"
    res: dict[str, Any] = {
        "emp_type": emp_type, "emp_id": emp_id, "workers": workers,
        "ref_no": "", "status": "", "note": "", "screenshot": "",
    }
    try:
        # ข้อ 1: เปิดฟอร์ม
        if not _open_inform_form(page, log=log):
            res["status"] = "FAIL"; res["note"] = "เปิดฟอร์มไม่สำเร็จ"
            return res

        # ข้อ 2: ค้นหานายจ้าง + บันทึก
        if not _inform_search_employer(page, emp_type, emp_id, log=log):
            res["status"] = "FAIL"; res["note"] = "ค้นหานายจ้างไม่สำเร็จ"
            return res
        log("      ✓ ค้นหานายจ้าง + บันทึก")

        # ข้อ 3: ติ๊กยืนยันรับรอง + ถัดไป
        _inform_consent_first(page, log=log)
        page.wait_for_timeout(800)
        if not _inform_click_next_step1(page, log=log):
            res["status"] = "FAIL"; res["note"] = "กดถัดไป (ขั้น 1) ไม่สำเร็จ"
            return res
        log("      ✓ ผ่านขั้นที่ 1 → เข้าสู่ขั้นที่ 2 (สาขา/รายชื่อ)")

        # ข้อ 4: ไล่ทุกสาขา → tick ทุก wp ที่ตรงกับ Excel (หลายคนในครั้งเดียว)
        branches = _inform_collect_branches(page)
        log(f"      สาขาในระบบ: {len(branches)} สาขา — จะไล่ทีละสาขา (เป้าหมาย {len(wp_targets)} คน)")
        matched_all: set[str] = set()
        for bi, br in enumerate(branches, start=1):
            log(f"        [สาขา {bi}/{len(branches)}] {br['text'][:80]}")
            if not _inform_select_branch(page, br["value"], log=log):
                continue
            n, unmatched = _inform_check_workpermits(page, wp_targets, log=log)
            if n:
                # บันทึกว่าตัวไหนถูกติ๊กในสาขานี้ (ตัวที่ไม่อยู่ใน unmatched)
                for wp in wp_targets:
                    if wp.replace(" ", "") not in [u.replace(" ", "") for u in unmatched]:
                        matched_all.add(wp.replace(" ", ""))
                log(f"          ✓ ติ๊ก {n} รายการในสาขานี้")
        matched_total = len(matched_all)
        if matched_total == 0:
            res["status"] = "FAIL"
            res["note"] = f"ไม่พบเลขใบอนุญาต {', '.join(wp_targets)} ในสาขาใดๆ"
            return res
        missing = [w for w in wp_targets if w.replace(" ", "") not in matched_all]
        if missing:
            log(f"      ⚠ ไม่พบ {len(missing)} คน: {', '.join(missing)} — ดำเนินต่อกับที่เจอ {matched_total} คน")
        log(f"      ✓ ติ๊กแรงงานครบ {matched_total}/{len(wp_targets)} คน → กดถัดไป (#beginNext2)")

        # กด 'ถัดไป' ของหน้าเลือกสาขา/รายชื่อ (#beginNext2) → ไปหน้าแนบเอกสาร
        page.evaluate(r"""() => {
            const b = document.querySelector('#beginNext2');
            if (b) b.click();
        }""")
        page.wait_for_timeout(2500)
        _inform_dismiss_popup(page, log=log)
        page.wait_for_timeout(1200)

        # ข้อ 5: แนบไฟล์ผู้รับมอบอำนาจ + ถัดไป
        files: list[Path] = []
        for f in group.get("files", []):
            p = (files_root / f) if not Path(f).is_absolute() else Path(f)
            if p.exists():
                files.append(p)
            else:
                log(f"      ⚠ ไม่พบไฟล์: {p}")
        _inform_attach_authorization(page, files, log=log)

        # กด 'ถัดไป' (#beginNext3) สู่หน้าสรุปคำขอ (ที่มีรายชื่อแรงงาน + วันที่จ้าง)
        page.wait_for_timeout(1000)
        page.evaluate(r"""() => {
            const b = document.querySelector('#beginNext3');
            if (b) b.click();
        }""")
        page.wait_for_timeout(3000)
        _inform_dismiss_popup(page, log=log)
        page.wait_for_timeout(1500)

        # ข้อ 6: เลือกวันที่จ้าง + บันทึก รายคน (บนหน้าสรุปคำขอ)
        for w in workers:
            wp = w.get("work_permit", "")
            hd = w.get("hire_date", "")
            if wp.replace(" ", "") not in matched_all:
                continue
            if hd:
                ok = _inform_set_hire_date(page, wp, hd, log=log)
                if not ok:
                    log(f"      ⚠ ตั้งวันที่จ้างไม่สำเร็จ: wp={wp}")
            else:
                log(f"      ⚠ Excel ไม่ระบุวันที่จ้าง: wp={wp}")

        # screenshot สถานะก่อนกดยืนยัน (ข้อ 7)
        screenshot_dir.mkdir(parents=True, exist_ok=True)
        shot = screenshot_dir / f"{emp_id or 'x'}_{wp_tag}_review.png"
        try:
            page.screenshot(path=str(shot), full_page=True)
            res["screenshot"] = shot.name
        except Exception:
            pass

        if not commit:
            res["status"] = "DRY_RUN_OK"
            res["note"] = f"หยุดก่อนข้อ 7 (ทดสอบ) — ติ๊ก {matched_total}/{len(wp_targets)} คน — ดู screenshot"
            return res

        # ข้อ 7: ติ๊กยืนยันสรุป (#checkRealityEvery) + กดส่งด้วย #beginNext4 เท่านั้น
        if not _inform_consent_summary(page, log=log):
            log("      ⚠ ติ๊กยืนยันสรุปครั้งแรกไม่สำเร็จ — ลองอีกครั้ง")
            page.wait_for_timeout(800)
            _inform_consent_summary(page, log=log)
        page.wait_for_timeout(800)
        # ตรวจว่า #beginNext4 พร้อมกด (consent ติดแล้ว) — ถ้าไม่ ยกเลิกการส่ง (ยังไม่ยื่น) เพื่อ retry
        clicked7 = page.evaluate(r"""() => {
            const b = document.querySelector('#beginNext4');
            if (!b) return 'notfound';
            if (b.offsetParent === null) return 'hidden';
            if (b.disabled) return 'disabled';
            b.click(); return 'ok';
        }""")
        if clicked7 != 'ok':
            log(f"      ✗ กดส่ง (#beginNext4) ไม่ได้: {clicked7} — ยกเลิก (ยังไม่ยื่นจริง) เพื่อให้ลองใหม่ได้")
            shot = screenshot_dir / f"{emp_id or 'x'}_{wp_tag}_step7fail.png"
            try:
                page.screenshot(path=str(shot), full_page=True); res["screenshot"] = shot.name
            except Exception:
                pass
            res["status"] = "FAIL"
            res["note"] = f"กดส่งไม่ได้ ({clicked7}) — consent อาจไม่ติด ยังไม่ได้ยื่น"
            return res
        log("      ✓ กดส่งคำขอ (#beginNext4)")

        # ข้อ 8: จัดการ popup ยืนยัน (SweetAlert2) + รอเข้าสู่ขั้นที่ 4 (เสร็จสิ้น) จริง
        advanced = False
        for _round in range(12):
            page.wait_for_timeout(1200)
            state = page.evaluate(r"""() => {
                const s4 = document.querySelector('#step4');
                const onStep4 = !!(s4 && s4.offsetParent !== null);
                const doneTxt = /ยื่นคำขอเรียบร้อย|ยื่นคำขอสำเร็จ|ดำเนินการเรียบร้อย|เลขที่คำขอ|เลขอ้างอิง|บันทึกข้อมูลเรียบร้อย/.test(document.body.innerText||'');
                const sw = document.querySelector('.swal2-popup');
                const swVisible = !!(sw && sw.offsetParent !== null);
                let swText = '', swType = 'confirm';
                if (swVisible) {
                    swText = (((document.querySelector('.swal2-title')||{}).innerText||'') + ' ' +
                              ((document.querySelector('.swal2-html-container')||{}).innerText||'')).trim();
                    if (document.querySelector('.swal2-icon.swal2-error, .swal2-icon.swal2-warning')) swType = 'error';
                }
                return {onStep4, doneTxt, swVisible, swText, swType};
            }""")
            if state.get("onStep4") or state.get("doneTxt"):
                advanced = True
                break
            if state.get("swVisible"):
                swt = state.get("swText", "")
                if swt:
                    log(f"      • popup: {swt[:120]}")
                if state.get("swType") == "error":
                    # ป๊อปแจ้ง error/validation — ปิดแล้วถือว่ายังยื่นไม่ผ่าน
                    page.evaluate(r"""() => { const b=document.querySelector('.swal2-confirm'); if(b) b.click(); }""")
                    res["note"] = f"ระบบแจ้งเตือน: {swt[:160]}"
                    break
                # ป๊อปยืนยัน — กดยืนยัน
                page.evaluate(r"""() => { const b=document.querySelector('.swal2-confirm'); if(b) b.click(); }""")
                log("      ✓ กดยืนยันใน popup (.swal2-confirm)")
                continue

        if not advanced:
            log("      ✗ ไม่เข้าสู่ขั้นที่ 4 หลังกดส่ง — ยังไม่ได้ยื่น (ปลอดภัย ลองใหม่ได้)")
            shot = screenshot_dir / f"{emp_id or 'x'}_{wp_tag}_step7fail.png"
            try:
                page.screenshot(path=str(shot), full_page=True); res["screenshot"] = shot.name
            except Exception:
                pass
            try:
                (screenshot_dir / f"{emp_id or 'x'}_{wp_tag}_step7fail.html").write_text(
                    page.content(), encoding="utf-8")
            except Exception:
                pass
            res["status"] = "FAIL"
            if not res.get("note"):
                res["note"] = "กดส่งแล้วแต่ไม่เข้าสู่ขั้นที่ 4 — ยังไม่ได้ยื่น"
            return res

        log("      ✓ เข้าสู่ขั้นที่ 4 (เสร็จสิ้น)")
        page.wait_for_timeout(1800)

        # ข้อ 9: เก็บหลักฐานให้ครบก่อน (one-shot!) — screenshot + HTML ก่อนแล้วค่อยดึงเลขอ้างอิง
        shot2 = screenshot_dir / f"{emp_id or 'x'}_{wp_tag}_done.png"
        try:
            page.screenshot(path=str(shot2), full_page=True)
            res["screenshot"] = shot2.name
        except Exception:
            pass
        try:
            (screenshot_dir / f"{emp_id or 'x'}_{wp_tag}_done.html").write_text(
                page.content(), encoding="utf-8")
        except Exception:
            pass
        ref_no = ""
        for _ in range(5):
            ref_no = _inform_extract_ref_no(page)
            if ref_no:
                break
            page.wait_for_timeout(1000)
        res["ref_no"] = ref_no
        res["status"] = "SUCCESS" if ref_no else "PARTIAL"
        res["note"] = (f"ยื่น {matched_total} คน (เลขอ้างอิง {ref_no})"
                       if ref_no else f"ยื่น {matched_total} คนแล้ว (ถึงขั้นที่ 4) แต่ดึงเลขอ้างอิงไม่ได้ — ดู screenshot/HTML")


    except Exception as e:
        res["status"] = "ERROR"; res["note"] = str(e)[:200]
        try:
            shot = screenshot_dir / f"{emp_id or 'x'}_{wp_tag}_error.png"
            page.screenshot(path=str(shot), full_page=True)
            res["screenshot"] = shot.name
        except Exception:
            pass
    return res



def run_inform_employer(
    cfg: dict,
    inform_excel: Path,
    login_excel: Path,
    out_path: Path,
    row_range: str | None = None,
    commit: bool = False,
    log=print,
    progress=None,
    is_cancelled=None,
) -> tuple[int, Path]:
    """โหมด INFORM_ENTER_EXIT — แจ้งการจ้างคนต่างด้าวเข้าทำงาน (แบบ บต.52)
    - อ่าน FormRequestEmployment.xlsx → รวมเป็นกลุ่มต่อนายจ้าง → ทำ flow ข้อ 1-9
    - commit=False → หยุดก่อนข้อ 7 (ทดสอบ) / commit=True → ส่งจริง
    - login จาก UsernameLogin.xlsx โดยจับคู่คอลัมน์ 'Users' ของแต่ละนายจ้าง
      กับ Username ใน UsernameLogin (ถ้าไม่ระบุ → ใช้บัญชีแรก)
    """
    out_path = _timestamped_path(out_path)
    records = _read_inform_data(inform_excel)
    accounts = _read_login_accounts(login_excel)
    if not accounts:
        raise ValueError(f"ไม่มีบัญชีใน {login_excel.name}")
    total = len(records)
    indices = _parse_row_range(row_range, total)
    selected = [records[i - 1] for i in indices]
    # รวมแถวที่เป็นนายจ้าง+บัญชีเดียวกันเข้าเป็นกลุ่มเดียว → 1 กลุ่ม = 1 การยื่น (หลายคนต่อครั้ง)
    groups = _group_inform_records(selected)
    log(f"[1/3] อ่าน {inform_excel.name} ({total} แถว) → เลือก {len(selected)} แถว: {row_range or 'ทั้งหมด'}")
    log(f"      จัดกลุ่มเป็น {len(groups)} นายจ้าง (ยื่นครั้งเดียวต่อนายจ้าง — ติ๊กแรงงานได้หลายคน)")
    log(f"      บัญชี login: {login_excel.name} ({len(accounts)} บัญชี)")
    log(f"      โหมด: {'COMMIT (ส่งจริง)' if commit else 'DRY-RUN (หยุดก่อนข้อ 7)'}")

    files_root = inform_excel.parent
    screenshot_dir = out_path.parent / "inform_screenshots"
    fallback_acct = next(iter(accounts.values()))

    results: list[dict[str, Any]] = []
    success = 0
    if progress:
        try: progress(0, len(groups))
        except Exception: pass

    headless = bool(cfg.get("headless", False))
    with sync_playwright() as pw:
        browser = _launch_chromium(pw, cfg, ["--ignore-certificate-errors", "--start-maximized"])
        ctx = browser.new_context(
            locale="th-TH", ignore_https_errors=True,
            viewport={"width": 1920, "height": 1080}, accept_downloads=True,
        )
        page = ctx.new_page()
        current_user = None  # username ที่ login อยู่ขณะนี้ (re-login เมื่อเปลี่ยน)
        try:
            _current_user = [None]
            for k, group in enumerate(groups, start=1):
                if is_cancelled and is_cancelled():
                    log("[!] ผู้ใช้ยกเลิก — หยุด"); break
                wps = [w.get("work_permit", "") for w in group.get("workers", [])]
                # จับคู่บัญชีล็อกอินให้ตรงกับ group (จากคอลัมน์ผู้ใช้/Login ใน Excel)
                _key = str(group.get("login", "")).strip().lower()
                _acct = accounts.get(_key)
                if not _acct:
                    _acct = next(iter(accounts.values()))
                    if group.get("login"):
                        log(f"      ⚠ ไม่พบบัญชี '{group.get('login')}' ใน {login_excel.name} — ใช้บัญชีแรกแทน")
                # เข้าสู่ระบบใหม่เมื่อเปลี่ยนบัญชี
                if _acct["username"] != _current_user[0]:
                    login(page, {
                        "username": _acct["username"],
                        "password": _acct["password"],
                        "user_type": _acct["type"],
                        "method": _acct.get("method") or cfg.get("method", "E-Service"),
                    })
                    page.wait_for_timeout(1500)
                    _current_user[0] = _acct["username"]
                    log(f"      🔐 เข้าสู่ระบบ: {_acct['username']} ({_acct['type']})")
                g_user = (group.get("username") or "").strip()

                # จับคู่บัญชี login ให้ตรงกับคอลัมน์ Users ของนายจ้างรายนี้
                acct = accounts.get(g_user.lower()) if g_user else None
                if not acct and not g_user:
                    acct = fallback_acct
                    log(f"[2/3] ({k}/{len(groups)}) นายจ้าง {group['emp_id']} | "
                        f"⚠ ไม่ได้ระบุ Users → ใช้บัญชีแรก {acct['username']}")
                if not acct:
                    log(f"[2/3] ({k}/{len(groups)}) นายจ้าง {group['emp_id']} | "
                        f"✗ ไม่พบบัญชี '{g_user}' ใน {login_excel.name} — ข้าม")
                    for w in group.get("workers", []):
                        results.append({
                            "seq": w.get("row_index", ""),
                            "username": group.get("username", ""),
                            "emp_type": group.get("emp_type", ""),
                            "emp_id": group.get("emp_id", ""),
                            "work_permit": w.get("work_permit", ""),
                            "hire_date": w.get("hire_date", ""),
                            "ref_no": "", "status": "NO_ACCOUNT",
                            "note": f"ไม่พบ Users '{g_user}' ใน {login_excel.name}",
                            "screenshot": "",
                        })
                    _save_inform_report(results, out_path, log=lambda *_: None)
                    if progress:
                        try: progress(k, len(groups))
                        except Exception: pass
                    continue

                # login (ครั้งแรก หรือเมื่อสลับบัญชี)
                if current_user != acct["username"].lower():
                    login_cfg = {
                        "username": acct["username"],
                        "password": acct["password"],
                        "user_type": acct["type"],
                        "method": acct.get("method") or cfg.get("method", "E-Service"),
                    }
                    log(f"      เข้าสู่ระบบ: {acct['username']} ({acct['type']}, {login_cfg['method']})")
                    try:
                        if current_user is not None:
                            _logout_safely(page)
                            page.wait_for_timeout(800)
                        login(page, login_cfg)
                        page.wait_for_timeout(1500)
                        current_user = acct["username"].lower()
                    except Exception as e:
                        log(f"      ✗ login ไม่สำเร็จ: {e} — ข้ามกลุ่มนี้")
                        for w in group.get("workers", []):
                            results.append({
                                "seq": w.get("row_index", ""),
                                "username": group.get("username", ""),
                                "emp_type": group.get("emp_type", ""),
                                "emp_id": group.get("emp_id", ""),
                                "work_permit": w.get("work_permit", ""),
                                "hire_date": w.get("hire_date", ""),
                                "ref_no": "", "status": "LOGIN_FAIL",
                                "note": str(e)[:150], "screenshot": "",
                            })
                        _save_inform_report(results, out_path, log=lambda *_: None)
                        if progress:
                            try: progress(k, len(groups))
                            except Exception: pass
                        continue

                log(f"[2/3] ({k}/{len(groups)}) นายจ้าง {group['emp_id']} | บัญชี {acct['username']} | "
                    f"{len(wps)} คน: {', '.join(wps)}")
                gres = _inform_process_one(
                    page, group, files_root, screenshot_dir, commit=commit, log=log,
                )
                st = gres.get("status", "")
                log(f"      → สถานะ: {st} | เลขอ้างอิง: {gres.get('ref_no','-')}")
                if st in ("SUCCESS", "DRY_RUN_OK"):
                    success += 1
                # แตกผลเป็นรายคน (1 แถวต่อ 1 work-permit) สำหรับ report
                for w in group.get("workers", []):
                    results.append({
                        "seq": w.get("row_index", ""),
                        "username": group.get("username", ""),
                        "emp_type": group.get("emp_type", ""),
                        "emp_id": group.get("emp_id", ""),
                        "work_permit": w.get("work_permit", ""),
                        "hire_date": w.get("hire_date", ""),
                        "ref_no": gres.get("ref_no", ""),
                        "status": gres.get("status", ""),
                        "note": gres.get("note", ""),
                        "screenshot": gres.get("screenshot", ""),
                    })
                if progress:
                    try: progress(k, len(groups))
                    except Exception: pass
                # บันทึก report เป็นระยะ
                _save_inform_report(results, out_path, log=lambda *_: None)
        finally:
            ctx.close(); browser.close()


    _save_inform_report(results, out_path, log=log)
    fail = sum(1 for r in results if r.get("status") not in ("SUCCESS", "DRY_RUN_OK"))
    log(f"[3/3] สรุป: สำเร็จ {success} / {len(groups)} นายจ้าง "
        f"({len(results)} รายการแรงงาน, ล้มเหลว {fail})")
    return success, out_path



def _ensure_on_tracking(page: Page, cfg: dict, request_type: str, status_ids: list[str], log=print) -> bool:
    """ยืนยันว่าอยู่หน้า e-Tracking + filter พร้อม (ฟื้น session ถ้าหลุด)"""
    try:
        url = page.url or ""
    except Exception:
        url = ""
    if "/Permit/Tracking" in url and not _is_logged_out(page):
        return True
    try:
        if _is_logged_out(page):
            login(page, cfg)
        goto_tracking(page)
        apply_wa_filter(page, request_type or "", status_ids=status_ids)
        return "/Permit/Tracking" in (page.url or "")
    except Exception as e:
        log(f"      ✗ ฟื้นหน้า tracking ไม่สำเร็จ: {e}")
        return False


def run_aliens_scrape(
    cfg: dict,
    out_path: Path,
    sub_tabs: list[str] | None = None,
    log=print,
    progress=None,
    is_cancelled=None,
) -> tuple[int, Path]:
    """รัน scraping แบบ 'ข้อมูลคนต่างด้าว' จากหน้าจัดการบัญชี
    sub_tabs: ชื่อ sub-tab ที่จะดึง (default = ทั้งหมด)
    """
    sub_tabs = sub_tabs or ALIEN_SUB_TABS
    out_path = _timestamped_path(out_path)
    log(f"[1/3] เริ่มดึงข้อมูลคนต่างด้าว — sub-tabs: {', '.join(sub_tabs)}")
    log(f"      ไฟล์รายงาน: {out_path.name}")
    headless = bool(cfg.get("headless", False))
    with sync_playwright() as pw:
        browser = _launch_chromium(pw, cfg)
        ctx = browser.new_context(locale="th-TH", ignore_https_errors=True)
        page = ctx.new_page()
        try:
            log("[2/3] เข้าสู่ระบบ...")
            login(page, cfg)
            page.wait_for_timeout(1500)
            data = scrape_aliens(page, sub_tabs, log=log)
            total = sum(len(v) for v in data.values())
            log(f"[3/3] บันทึกไฟล์ Excel: {out_path} (รวม {total} แถว / {len(data)} sheet)")
            save_excel_aliens(data, out_path)
            return total, out_path
        finally:
            ctx.close(); browser.close()


# ─────────────────────────────────────────────────────────────────
# โหมด บต.30 — ยื่นต่ออายุใบอนุญาตทำงานของคนต่างด้าวตาม MoU (MT_59_MOU_RENEWAL)
#   1) Login จาก UsernameLogin.xlsx
#   2) เมนูบริการ → การยื่นขอต่ออายุใบอนุญาตทำงาน → ...ตาม MoU (แบบ บต.30)
#   3) กรอกฟอร์ม "ค้นหาข้อมูลคนต่างด้าว" ตาม from_bt30.xlsx แล้วกดบันทึก
# ─────────────────────────────────────────────────────────────────
# URL ฟอร์ม บต.30 (จาก setFormTypeRenew ของเมนู MT_59_MOU_RENEWAL)
BT30_FORM_URL = (
    "https://eworkpermit.doe.go.th/WorkPermit?user_type=alien&ft=RENEW_REQ&ut=alien&uti=3"
)


def _read_bt30_excel(path: Path) -> list[dict[str, Any]]:
    """อ่าน from_bt30.xlsx → list of dict (ข้อมูลครบทั้งขั้นตอน 1 และ 2)
    คอลัมน์ (header แถวแรก, ลำดับคงที่ตาม template):
      ขั้น 1 : No., คำนำหน้า, ชื่อ, สัญชาติ, เพศ, วันเกิด
      2.2    : เลขที่, หมู่ที่/อาคาร, ซอย, ถนน, จังหวัด, เขต/อำเภอ, แขวง/ตำบล
      2.3    : ประเภทเอกสาร, เลขที่เอกสาร, สถานที่ออกให้, ออกให้วันที่, ใช้ได้ถึงวันที่
      วีซ่า  : เลขที่การตรวจลงตรา, ตรวจลงตราประเภท, ออกให้ที่, ออกให้วันที่, ใช้ได้ถึงวันที่
      2.4    : ได้รับอนุญาตจากพนักงานเจ้าหน้าที่ตรวจคนเข้าเมือง, วันที่เดินทางมาถึงราชอาณาจักร, อยู่ได้ถึงวันที่
      หน้า2/2: สถานที่ทำงาน/สาขา, ประเภทกิจการ,
               เลขที่/ออกให้โดย/วันที่ออกเอกสาร/วันที่เอกสารหมดอายุ (*เอกสารแสดงการอนุญาตหรือการรับรอง*)
    วันที่ทุกช่องแปลงเป็น dd/mm/yyyy (ค.ศ.) ตามที่ datepicker ของเว็บรองรับ
    (header วันที่ซ้ำกัน จึง anchor ตำแหน่งจากคอลัมน์ชื่อไม่ซ้ำที่อยู่ก่อนหน้า)
    """
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    hdr = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    def col(*names: str) -> int:
        for nm in names:
            for i, h in enumerate(hdr):
                if h == nm:
                    return i
        for nm in names:
            for i, h in enumerate(hdr):
                if nm and nm in h:
                    return i
        return -1

    # --- ขั้นตอน 1: ข้อมูลคนต่างด้าว ---
    i_no = col("No.", "No", "ลำดับ", "ลําดับ")
    i_prefix = col("คำนำหน้า", "คํานําหน้า", "Prefix", "Title")
    i_name = col("ชื่อ", "ชื่อ-สกุล", "Name")
    i_nat = col("สัญชาติ", "Nationality")
    i_sex = col("เพศ", "Sex", "Gender")
    i_birth = col("วันเกิด", "วันเดือนปีเกิด", "BirthDate", "DOB", "Birth")
    # บัญชี login ที่ใช้ยื่นของแถวนี้ (รองรับยื่นหลาย Username) — เว้นว่าง = ใช้บัญชีหลัก
    i_username = col("Username", "username", "ชื่อผู้ใช้")

    # --- ขั้นตอน 2.2: ที่อยู่ที่ติดต่อได้ (header ไม่ซ้ำ จับด้วยชื่อได้) ---
    i_addr_no = col("เลขที่")            # exact → ' เลขที่' (ไม่ชนกับ 'เลขที่เอกสาร')
    i_addr_moo = col("หมู่ที่/อาคาร", "หมู่ที่", "อาคาร")
    i_addr_soi = col("ซอย")
    i_addr_road = col("ถนน")
    i_addr_prov = col("จังหวัด")
    i_addr_dist = col("เขต/อำเภอ")
    i_addr_sub = col("แขวง/ตำบล")

    # --- ขั้นตอน 2.3: เอกสารแสดงการได้รับอนุญาต (วันที่ใช้ header ซ้ำ → anchor ตำแหน่ง) ---
    i_doc_type = col("ประเภทเอกสาร")     # เอกสารแสดงการได้รับอนุญาต (หนังสือเดินทาง ฯลฯ)
    i_doc_no = col("เลขที่เอกสาร")
    i_doc_place = col("สถานที่ออกให้")
    # ออกให้วันที่ / ใช้ได้ถึงวันที่ ของเอกสาร = ถัดจาก 'สถานที่ออกให้'
    i_doc_issue = i_doc_place + 1 if i_doc_place >= 0 else -1
    i_doc_expire = i_doc_place + 2 if i_doc_place >= 0 else -1

    # --- การตรวจลงตรา (วีซ่า) ---
    i_visa_no = col("เลขที่การตรวจลงตรา")
    i_visa_type = col("ตรวจลงตราประเภท")
    i_visa_place = col("ออกให้ที่")
    i_visa_issue = i_visa_place + 1 if i_visa_place >= 0 else -1
    i_visa_expire = i_visa_place + 2 if i_visa_place >= 0 else -1

    # --- ขั้นตอน 2.4: ได้รับอนุญาตจากพนักงานเจ้าหน้าที่ตรวจคนเข้าเมือง ---
    i_imm = col("ได้รับอนุญาตจากพนักงานเจ้าหน้าที่ตรวจคนเข้าเมือง", "ตรวจคนเข้าเมือง")
    i_arrival = col("วันที่เดินทางมาถึงราชอาณาจักร", "เดินทางมาถึง")
    i_stay_until = i_arrival + 1 if i_arrival >= 0 else -1

    # --- หน้า 2/2 (Step 3): 'โดยจะมาทำงาน' + 'เอกสารแสดงการอนุญาตหรือการรับรอง' ---
    # คอลัมน์เอกสารอนุญาตมี header ซ้ำคำกับ 2.2/2.3 จึงจับด้วย keyword + marker เฉพาะหัวข้อนี้
    def col_auth(keyword: str) -> int:
        mark = "เอกสารแสดงการอนุญาตหรือการรับรอง"
        for j, h in enumerate(hdr):
            if keyword in h and mark in h:
                return j
        return -1

    i_work_place = col("สถานที่ทำงาน/สาขา", "สถานที่ทำงาน")
    i_work_biz = col("ประเภทกิจการ")
    i_auth_no = col_auth("เลขที่")
    i_auth_by = col_auth("ออกให้โดย")
    i_auth_issue = col_auth("วันที่ออกเอกสาร")
    i_auth_expire = col_auth("วันที่เอกสารหมดอายุ")

    # --- ขั้นตอน 3 (แนบเอกสาร): คอลัมน์ path ไฟล์เอกสารแต่ละประเภท ---
    # header ยาวและไม่ซ้ำ → จับด้วย keyword เฉพาะ (contains)
    i_doc_passport = col("สำเนาหนังสือเดินทาง")                  # 3.1
    i_doc_entry = col("หลักฐานการอนุญาตให้เข้ามาในราชอาณาจักร")   # 3.2
    i_doc_contract = col("สำเนาสัญญาจ้าง")                        # 3.3
    i_doc_medical = col("เวชกรรม")                                # 3.4 (ใบรับรองแพทย์)
    i_doc_photo = col("รูปถ่าย")                                  # 3.5
    i_doc_bt46 = col("บต.46")                                     # 3.6
    i_doc_poa_agent = col("ดำเนินการแทน")                         # 3.7 หนังสือมอบอำนาจ
    i_doc_poa_stamp = col("อากรแสตมป์")                           # 3.8 ใบมอบอำนาจติดอากร
    i_doc_id_grantor = col("ประชาชนของผู้มอบอำนาจ")               # 3.9 บัตร ปชช ผู้มอบอำนาจ
    i_doc_id_grantee = col("ประชาชนของผู้รับมอบอำนาจ")            # 4.0 บัตร ปชช ผู้รับมอบอำนาจ
    i_doc_workpermit = col("ใบอนุญาตทำงาน")                       # 4.1
    # 4.2 เอกสารอื่นๆที่เกี่ยวข้อง 1..5 (แนบได้หลายไฟล์ — ทีละไฟล์)
    i_doc_others = [c for c in (col(f"เอกสารอื่นๆที่เกี่ยวข้อง {n}") for n in range(1, 9)) if c >= 0]

    # --- คอลัมน์ผลลัพธ์ที่ระบบเคยเขียนกลับ (สำหรับ resume/skip คนที่ส่งคำขอแล้ว) ---
    i_done_status = col("สถานะส่งคำขอ")
    i_done_req = col("เลขที่คำขอ")

    base_dir = path.parent

    def fmt_date(v: Any) -> str:
        if v is None or v == "":
            return ""
        if isinstance(v, (datetime, date)):
            return v.strftime("%d/%m/%Y")
        return str(v).strip()

    rows: list[dict[str, Any]] = []
    for ridx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=1):
        if not any(v not in (None, "") for v in row):
            continue

        def g(i: int) -> str:
            return str(row[i]).strip() if 0 <= i < len(row) and row[i] is not None else ""

        def gd(i: int) -> str:
            return fmt_date(row[i] if 0 <= i < len(row) else None)

        def gp(i: int) -> str:
            """อ่าน path ไฟล์เอกสาร → absolute path (อิงโฟลเดอร์ของไฟล์ Excel)"""
            s = g(i)
            if not s:
                return ""
            p = Path(s.replace("\\", "/").strip())
            if not p.is_absolute():
                p = base_dir / p
            return str(p)

        rec = {
            # ขั้นตอน 1
            "seq": g(i_no) or str(ridx),
            "prefix": g(i_prefix),
            "name": g(i_name),
            "nationality": g(i_nat),
            "sex": g(i_sex),
            "birthdate": gd(i_birth),
            # 2.2 ที่อยู่ที่ติดต่อได้
            "addr_no": g(i_addr_no),
            "addr_moo": g(i_addr_moo),
            "addr_soi": g(i_addr_soi),
            "addr_road": g(i_addr_road),
            "addr_prov": g(i_addr_prov),
            "addr_dist": g(i_addr_dist),
            "addr_subdist": g(i_addr_sub),
            # 2.3 เอกสารแสดงการได้รับอนุญาต
            "stay_doc_type": g(i_doc_type),
            "doc_no": g(i_doc_no),
            "doc_place": g(i_doc_place),
            "doc_issue": gd(i_doc_issue),
            "doc_expire": gd(i_doc_expire),
            # การตรวจลงตรา (วีซ่า)
            "visa_no": g(i_visa_no),
            "visa_type": g(i_visa_type),
            "visa_place": g(i_visa_place),
            "visa_issue": gd(i_visa_issue),
            "visa_expire": gd(i_visa_expire),
            # 2.4 ตรวจคนเข้าเมือง
            "imm_office": g(i_imm),
            "arrival_date": gd(i_arrival),
            "stay_until": gd(i_stay_until),
            # หน้า 2/2: โดยจะมาทำงาน
            "work_place": g(i_work_place),
            "work_biz": g(i_work_biz),
            # หน้า 2/2: เอกสารแสดงการอนุญาตหรือการรับรอง
            "auth_no": g(i_auth_no),
            "auth_by": g(i_auth_by),
            "auth_issue": gd(i_auth_issue),
            "auth_expire": gd(i_auth_expire),
            # ขั้นตอน 3: แนบเอกสาร (path ไฟล์)
            "doc_passport": gp(i_doc_passport),
            "doc_entry": gp(i_doc_entry),
            "doc_contract": gp(i_doc_contract),
            "doc_medical": gp(i_doc_medical),
            "doc_photo": gp(i_doc_photo),
            "doc_bt46": gp(i_doc_bt46),
            "doc_poa_agent": gp(i_doc_poa_agent),
            "doc_poa_stamp": gp(i_doc_poa_stamp),
            "doc_id_grantor": gp(i_doc_id_grantor),
            "doc_id_grantee": gp(i_doc_id_grantee),
            "doc_workpermit": gp(i_doc_workpermit),
            "doc_others": [gp(j) for j in i_doc_others if gp(j)],
            # ผลลัพธ์เดิมที่ระบบเคยเขียนกลับ (resume/skip)
            "done_submit_status": g(i_done_status),
            "done_request_no": g(i_done_req),
            # บัญชี login รายแถว (รองรับยื่นหลาย Username) — ว่าง = ใช้บัญชีหลัก
            "username": g(i_username),
            "row_index": ridx,
        }
        if rec["name"] or rec["birthdate"]:
            rows.append(rec)
    return rows


def _select_option_by_text(page: Page, sel_id: str, want: str) -> bool:
    """เลือก option ของ <select id=sel_id> ที่ข้อความตรง/ใกล้เคียงกับ want
    ลำดับการจับคู่: ตรงเป๊ะ → contains → ขึ้นต้นด้วย (ข้าม 'กรุณาเลือก')
    """
    return bool(page.evaluate(
        r"""(args) => {
          const { id, want } = args;
          const s = document.getElementById(id);
          if (!s) return false;
          const norm = x => String(x || '').replace(/\s+/g, '').toLowerCase();
          const w = norm(want);
          if (!w) return false;
          const opts = Array.from(s.options).filter(o => norm(o.textContent) !== 'กรุณาเลือก'.toLowerCase());
          let opt = opts.find(o => norm(o.textContent) === w);
          if (!opt) opt = opts.find(o => norm(o.textContent).includes(w));
          if (!opt) opt = opts.find(o => w.includes(norm(o.textContent)) && norm(o.textContent).length >= 2);
          if (!opt) return false;
          s.value = opt.value;
          s.dispatchEvent(new Event('change', { bubbles: true }));
          if (window.jQuery) { try { jQuery(s).trigger('change'); } catch (e) {} }
          return true;
        }""",
        {"id": sel_id, "want": want},
    ))


def _select2_pick(page: Page, sel_id: str, want: str, log=print) -> bool:
    """เลือก option ของ <select id=sel_id> ที่ถูกครอบด้วย select2 — แบบ user จริง
    1) set ค่า native select + ยิง change ให้ครบ (input/change + jQuery + select2:select)
    2) ถ้า select2 UI ไม่สะท้อนค่า → คลิกเปิด dropdown แล้วคลิก option ที่ตรงข้อความ
    คืน True เมื่อค่าถูกเลือก (select2 UI สะท้อนค่า หรือ native ตรงเมื่อไม่มี select2)
    """
    norm = lambda x: "".join(str(x or "").split()).lower()
    w = norm(want)
    if not w:
        return False

    # (A) ตั้งค่า native + ยิง event ครบชุด (ครอบคลุม select2 ที่ผูกผ่าน jQuery)
    set_ok = bool(page.evaluate(
        r"""(args) => {
          const { id, want } = args;
          const s = document.getElementById(id);
          if (!s) return false;
          const norm = x => String(x || '').replace(/\s+/g, '').toLowerCase();
          const w = norm(want);
          const opts = Array.from(s.options).filter(o => norm(o.textContent) !== 'กรุณาเลือก'.toLowerCase());
          let opt = opts.find(o => norm(o.textContent) === w)
                 || opts.find(o => norm(o.textContent).includes(w))
                 || opts.find(o => w.includes(norm(o.textContent)) && norm(o.textContent).length >= 2);
          if (!opt) return false;
          s.value = opt.value;
          s.dispatchEvent(new Event('input', { bubbles: true }));
          s.dispatchEvent(new Event('change', { bubbles: true }));
          if (window.jQuery) {
            try {
              const $s = jQuery(s);
              $s.val(opt.value);
              $s.trigger('change');
              $s.trigger({ type: 'select2:select', params: { data: { id: opt.value, text: opt.textContent } } });
            } catch (e) {}
          }
          return true;
        }""",
        {"id": sel_id, "want": want},
    ))

    def select2_reflects() -> bool:
        return bool(page.evaluate(
            r"""(args) => {
              const { id, want } = args;
              const s = document.getElementById(id);
              if (!s) return false;
              const norm = x => String(x || '').replace(/\s+/g, '').toLowerCase();
              const w = norm(want);
              let rendered = null;
              const sib = s.nextElementSibling;
              if (sib && sib.classList && sib.classList.contains('select2')) {
                rendered = sib.querySelector('.select2-selection__rendered');
              }
              if (!rendered) {
                const cont = document.querySelector('.select2-selection__rendered');
                rendered = cont || null;
              }
              if (!rendered) return true; // ไม่มี select2 → native พอ
              return norm(rendered.textContent).includes(w);
            }""",
            {"id": sel_id, "want": want},
        ))

    if select2_reflects():
        return set_ok

    # (C) คลิกเปิด select2 dropdown แล้วเลือก option ที่ตรง (แบบ user)
    try:
        opened = bool(page.evaluate(
            r"""(id) => {
              const s = document.getElementById(id);
              if (!s) return false;
              let container = (s.nextElementSibling && s.nextElementSibling.classList && s.nextElementSibling.classList.contains('select2'))
                ? s.nextElementSibling : null;
              if (!container && s.parentElement) {
                container = s.parentElement.querySelector('.select2');
              }
              const sel = container && container.querySelector('.select2-selection');
              if (!sel) return false;
              sel.scrollIntoView({ block: 'center' });
              sel.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
              sel.click();
              return true;
            }""",
            sel_id,
        ))
        if opened:
            page.wait_for_timeout(300)
            picked = bool(page.evaluate(
                r"""(want) => {
                  const norm = x => String(x || '').replace(/\s+/g, '').toLowerCase();
                  const w = norm(want);
                  const opts = Array.from(document.querySelectorAll('.select2-results__option'));
                  let opt = opts.find(o => norm(o.textContent) === w)
                         || opts.find(o => norm(o.textContent).includes(w))
                         || opts.find(o => w.includes(norm(o.textContent)) && norm(o.textContent).length >= 2);
                  if (!opt) return false;
                  opt.scrollIntoView({ block: 'center' });
                  opt.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
                  opt.click();
                  return true;
                }""",
                want,
            ))
            page.wait_for_timeout(300)
            if picked:
                return True
    except Exception as e:
        log(f"      ⚠ select2 pick error: {str(e)[:80]}")

    return set_ok


def _open_bt30_form(page: Page, log=print) -> bool:
    """เปิดฟอร์ม บต.30 (MT_59_MOU_RENEWAL) ผ่านเมนูบริการ
    flow: หน้าหลัก → 'เมนูบริการ' → openCity('tab_RENEW_REQ') → คลิก #MT_59_MOU_RENEWAL
    (onclick = setFormTypeRenew(...) ซึ่งตั้งค่า form แล้วนำทางไป /WorkPermit?...ft=RENEW_REQ)
    """
    try:
        page.goto("https://eworkpermit.doe.go.th/", wait_until="domcontentloaded", timeout=30_000)
        page.wait_for_timeout(2000)
    except Exception as e:
        log(f"      ✗ เปิดหน้าหลักไม่สำเร็จ: {e}")
        return False

    # 1) คลิก 'เมนูบริการ'
    clicked = page.evaluate(r"""() => {
        const a = document.querySelector('a.lang_menu_service')
          || Array.from(document.querySelectorAll('a,button'))
               .find(x => /เมนูบริการ/.test((x.textContent || '').trim()));
        if (a) { a.click(); return true; }
        return false;
    }""")
    if not clicked:
        log("      ✗ ไม่พบลิงก์ 'เมนูบริการ'")
        return False
    page.wait_for_timeout(1200)

    # 2) เปิดหมวด 'การยื่นขอต่ออายุใบอนุญาตทำงาน' (tab_RENEW_REQ) แล้วคลิก MT_59_MOU_RENEWAL
    page.evaluate(r"""() => {
        try { if (typeof openCity === 'function') openCity('tab_RENEW_REQ', new Event('click')); } catch (e) {}
    }""")
    page.wait_for_timeout(900)
    ok = page.evaluate(r"""() => {
        const t = document.querySelector('#MT_59_MOU_RENEWAL');
        if (t) { t.click(); return true; }
        return false;
    }""")
    if not ok:
        log("      ✗ ไม่พบเมนู #MT_59_MOU_RENEWAL (แบบ บต.30)")
        return False

    # 3) รอเข้าหน้าฟอร์ม + ปุ่ม 'ค้นหาข้อมูลคนต่างด้าว' พร้อม
    try:
        page.wait_for_url("**/WorkPermit**", timeout=30_000)
    except PWTimeoutError:
        log(f"      ⚠ ยังไม่เข้าฟอร์ม (URL={page.url}) — ลองรอ element ต่อ")
    page.wait_for_timeout(2500)
    try:
        page.wait_for_function(
            r"""() => Array.from(document.querySelectorAll('button,a'))
                  .some(e => /search_alien_modal\.show/.test(e.getAttribute('onclick') || ''))""",
            timeout=20_000,
        )
        return True
    except PWTimeoutError:
        log(f"      ✗ เปิดฟอร์ม บต.30 ไม่สำเร็จ (URL={page.url})")
        return False


def _bt30_field_errors(page: Page) -> str:
    """อ่านข้อความ validation (label.error) ที่แสดงอยู่ในฟอร์ม"""
    try:
        return page.evaluate(r"""() => {
            return Array.from(document.querySelectorAll('label.error, .error'))
              .filter(e => e.offsetParent !== null && (e.textContent || '').trim())
              .map(e => e.textContent.trim()).slice(0, 6).join(' | ');
        }""") or ""
    except Exception:
        return ""


def _bt30_fill_search_one(
    page: Page,
    rec: dict[str, Any],
    screenshot_dir: Path,
    log=print,
) -> dict[str, Any]:
    """เปิด modal 'ค้นหาข้อมูลคนต่างด้าว' → กรอกข้อมูล 1 คน → กดบันทึก → เก็บผล
    คืน {status, note, screenshot}
    """
    seq = rec.get("seq", "?")
    name = rec.get("name", "")
    base = _safe_filename(f"{seq}_{name}")
    res: dict[str, Any] = {"status": "", "note": "", "screenshot": ""}

    try:
        # เปิด modal ค้นหาข้อมูลคนต่างด้าว
        opened = page.evaluate(r"""() => {
            const b = Array.from(document.querySelectorAll('button,a'))
              .find(e => /search_alien_modal\.show/.test(e.getAttribute('onclick') || ''));
            if (b) { b.click(); return true; }
            return false;
        }""")
        if not opened:
            res["status"] = "FAIL"
            res["note"] = "ไม่พบปุ่ม 'ค้นหาข้อมูลคนต่างด้าว'"
            return res
        page.wait_for_selector("#btn_search_alien_submit", state="visible", timeout=10_000)
        page.wait_for_timeout(600)

        # คำนำหน้า (select)
        if rec.get("prefix") and not _select_option_by_text(page, "alien_prefix", rec["prefix"]):
            log(f"      ⚠ เลือกคำนำหน้า '{rec['prefix']}' ไม่ได้")
        # ชื่อ (text) — ใส่ชื่อเต็มตาม Excel
        page.evaluate(
            r"""(v) => {
              const t = document.getElementById('other_name');
              if (t) {
                t.value = v;
                t.dispatchEvent(new Event('input', { bubbles: true }));
                t.dispatchEvent(new Event('change', { bubbles: true }));
              }
            }""", name,
        )
        # สัญชาติ (select)
        if rec.get("nationality") and not _select_option_by_text(page, "nationality_al", rec["nationality"]):
            log(f"      ⚠ เลือกสัญชาติ '{rec['nationality']}' ไม่ได้")
        # เพศ (select)
        if rec.get("sex") and not _select_option_by_text(page, "sexCheck", rec["sex"]):
            log(f"      ⚠ เลือกเพศ '{rec['sex']}' ไม่ได้")
        # วันเกิด (datepicker text, dd/mm/yyyy ค.ศ.)
        if rec.get("birthdate"):
            page.evaluate(
                r"""(v) => {
                  const t = document.getElementById('birthDateCheck');
                  if (t) {
                    t.removeAttribute('readonly');
                    t.value = v;
                    t.dispatchEvent(new Event('input', { bubbles: true }));
                    t.dispatchEvent(new Event('change', { bubbles: true }));
                    t.dispatchEvent(new Event('blur', { bubbles: true }));
                    if (window.jQuery) { try { jQuery(t).trigger('change'); } catch (e) {} }
                  }
                }""", rec["birthdate"],
            )
        page.wait_for_timeout(400)
        try:
            page.screenshot(path=str(screenshot_dir / f"{base}_filled.png"), full_page=True)
            res["screenshot"] = f"{base}_filled.png"
        except Exception:
            pass

        # กดบันทึก (#btn_search_alien_submit)
        page.evaluate(r"""() => { const b = document.getElementById('btn_search_alien_submit'); if (b) b.click(); }""")
        page.wait_for_timeout(2800)

        alert = _capture_register_alert(page)
        errs = _bt30_field_errors(page)
        modal_open = page.evaluate(
            r"""() => { const b = document.getElementById('btn_search_alien_submit'); return !!(b && b.offsetParent !== null); }"""
        )

        # บันทึก screenshot หลังกด
        try:
            page.screenshot(path=str(screenshot_dir / f"{base}_aftersave.png"), full_page=True)
            res["screenshot"] = f"{base}_aftersave.png"
        except Exception:
            pass

        if alert and any(k in alert for k in ("เสร็จสมบูรณ์", "เรียบร้อยแล้ว", "เรียบร้อย", "สำเร็จแล้ว")):
            # ระบบเพิ่มข้อมูลคนต่างด้าวที่ต้องการดำเนินการแทนเรียบร้อย (= บันทึกค้นหาสำเร็จ)
            res["status"] = "SUCCESS"
            res["note"] = alert[:400]
            _close_register_alert(page)  # กดปุ่ม 'ปิด' — ไม่กด 'ยินยอม' (อยู่นอกขอบเขต Step 1)
        elif alert and any(k in alert for k in ("ไม่พบ", "ไม่ถูกต้อง", "ผิดพลาด", "ไม่สำเร็จ", "ไม่สามารถ", "ซ้ำ", "กรอกข้อมูล")):
            res["status"] = "ALERT"
            res["note"] = alert[:400]
            _close_register_alert(page)
        elif alert:
            res["status"] = "REVIEW"
            res["note"] = alert[:400]
            _close_register_alert(page)
        elif errs:
            res["status"] = "VALIDATE"
            res["note"] = f"ฟอร์มแจ้งเตือน: {errs}"[:300]
        elif not modal_open:
            res["status"] = "SUCCESS"
            res["note"] = "บันทึกค้นหา (modal ปิด)"
        else:
            res["status"] = "REVIEW"
            res["note"] = "modal ยังเปิดอยู่หลังกดบันทึก — ตรวจสอบ screenshot"
        return res
    except Exception as e:
        try:
            page.screenshot(path=str(screenshot_dir / f"{base}_error.png"), full_page=True)
            res["screenshot"] = f"{base}_error.png"
        except Exception:
            pass
        res["status"] = "ERROR"
        res["note"] = str(e)[:300]
        return res


# ====================== บต.30 ขั้นตอนที่ 2 (กรอกรายละเอียดคำขอ) ======================

def _bt30_fill_text(page: Page, field_id: str, value: str) -> None:
    """กรอก <input id=field_id> ด้วย value (ปลด readonly + dispatch input/change/keyup)"""
    if value is None or value == "":
        return
    page.evaluate(
        r"""(args) => {
          const t = document.getElementById(args.id);
          if (!t) return;
          t.removeAttribute('readonly');
          t.value = args.v;
          t.dispatchEvent(new Event('input', { bubbles: true }));
          t.dispatchEvent(new Event('change', { bubbles: true }));
          t.dispatchEvent(new Event('keyup', { bubbles: true }));
        }""",
        {"id": field_id, "v": str(value)},
    )


def _bt30_fill_date(page: Page, field_id: str, value: str) -> None:
    """กรอก datepicker (dd/mm/yyyy ค.ศ.) — ปลด readonly + set value + dispatch + jQuery trigger"""
    if not value:
        return
    page.evaluate(
        r"""(args) => {
          const t = document.getElementById(args.id);
          if (!t) return;
          t.removeAttribute('readonly');
          t.value = args.v;
          t.dispatchEvent(new Event('input', { bubbles: true }));
          t.dispatchEvent(new Event('change', { bubbles: true }));
          t.dispatchEvent(new Event('blur', { bubbles: true }));
          if (window.jQuery) { try { jQuery(t).trigger('change'); jQuery(t).trigger('blur'); } catch (e) {} }
        }""",
        {"id": field_id, "v": str(value)},
    )


def _bt30_wait_select_option(page: Page, sel_id: str, want: str, timeout: int = 9000) -> bool:
    """รอจน <select id=sel_id> มี option ที่ตรง/ใกล้เคียง want (สำหรับ dropdown cascading จังหวัด→อำเภอ→ตำบล)"""
    if not want:
        return False
    try:
        page.wait_for_function(
            r"""(args) => {
              const s = document.getElementById(args.id);
              if (!s) return false;
              const norm = x => String(x || '').replace(/\s+/g,'').toLowerCase();
              const w = norm(args.want);
              return Array.from(s.options).some(o => {
                const t = norm(o.textContent);
                return t && t !== 'กรุณาเลือก'.toLowerCase() && (t === w || t.includes(w) || w.includes(t));
              });
            }""",
            arg={"id": sel_id, "want": want},
            timeout=timeout,
        )
        return True
    except PWTimeoutError:
        return False


def _bt30_dismiss_news(page: Page, log=print) -> None:
    """ปิด popup ข่าวสาร/ประชาสัมพันธ์ บนหน้ารายละเอียด (gentle — ไม่กด Escape เพื่อไม่ปิด modal ที่กำลังจะเปิด)"""
    try:
        page.evaluate(r"""() => {
          document.querySelectorAll('.modal').forEach(m => {
            const txt = (m.textContent || '');
            if (/ข่าวสาร|ประชาสัมพันธ์/.test(txt)) {
              const x = m.querySelector('.close, [data-dismiss="modal"], button');
              if (x) { try { x.click(); } catch (e) {} }
              try { if (window.jQuery) jQuery(m).modal('hide'); } catch (e) {}
              m.classList.remove('show'); m.style.display = 'none';
            }
          });
          document.querySelectorAll('.modal-backdrop').forEach(b => b.remove());
          document.body.classList.remove('modal-open');
          document.body.style.removeProperty('overflow');
          document.body.style.removeProperty('padding-right');
        }""")
    except Exception:
        pass


def _bt30_consent_next(page: Page, log=print) -> bool:
    """2.1 ติ๊ก checkbox รับรอง (#check_truth) → กด 'ถัดไป' (#gonextSubmit) → รอเข้าหน้ารายละเอียด"""
    checked = page.evaluate(r"""() => {
        const matchTxt = el => {
            let parts = [];
            const lbl = el.closest('label') || (el.id ? document.querySelector(`label[for="${el.id}"]`) : null);
            if (lbl) parts.push(lbl.textContent || '');
            let p = el.parentElement, lvl = 0;
            while (p && lvl < 5) { parts.push(p.textContent || ''); p = p.parentElement; lvl++; }
            return parts.join(' ');
        };
        let cb = document.getElementById('check_truth');
        if (!cb || cb.offsetParent === null) {
            cb = Array.from(document.querySelectorAll('input[type=checkbox]'))
                .find(c => c.offsetParent !== null &&
                    /ขอรับรองว่า|มีความประสงค์ในการยื่นคำขอ|ได้รับความยินยอม/.test(matchTxt(c)));
        }
        if (!cb) return false;
        if (!cb.checked) {
            cb.click();
            if (!cb.checked) {
                cb.checked = true;
                cb.dispatchEvent(new Event('change', { bubbles: true }));
                cb.dispatchEvent(new Event('click', { bubbles: true }));
            }
        }
        return true;
    }""")
    if not checked:
        log("      ⚠ ไม่พบ checkbox รับรอง (2.1) — ลองเดินหน้าต่อ")
    page.wait_for_timeout(500)

    clicked = page.evaluate(r"""() => {
        const b = document.querySelector('#gonextSubmit:not(.d-none)') || document.getElementById('gonextSubmit');
        if (b) { b.click(); return true; }
        const b2 = Array.from(document.querySelectorAll('button,a'))
          .find(x => /openSubmitPage/.test(x.getAttribute('onclick') || ''));
        if (b2) { b2.click(); return true; }
        return false;
    }""")
    if not clicked:
        log("      ✗ ไม่พบปุ่ม 'ถัดไป' (2.1 / #gonextSubmit)")
        return False
    page.wait_for_timeout(1500)
    # ถ้ามี SweetAlert ยืนยัน
    try:
        page.evaluate(r"""() => {
            const b = document.querySelector('.swal2-confirm');
            if (b && b.offsetParent !== null) b.click();
        }""")
    except Exception:
        pass
    try:
        page.wait_for_url("**/RenewMOU/FormRenewMOU**", timeout=25_000)
    except PWTimeoutError:
        log(f"      ⚠ ยังไม่เข้าหน้ารายละเอียด (URL={page.url})")
    page.wait_for_timeout(2500)
    _bt30_dismiss_news(page, log=log)
    return True


def _bt30_fill_address(page: Page, rec: dict[str, Any], log=print) -> dict[str, Any]:
    """2.2 เปิด modal 'ที่อยู่ที่ติดต่อได้' → กรอกที่อยู่ → บันทึก (#btn_current_address_save)"""
    res = {"ok": False, "note": ""}
    notes: list[str] = []
    opened = page.evaluate(r"""() => {
        const b = Array.from(document.querySelectorAll('button,a'))
          .find(x => /ModalEditAlienComponentCurrenAddress/.test(x.getAttribute('onclick') || ''));
        if (b) { b.click(); return true; }
        try {
            if (typeof ModalEditAlienComponentCurrenAddress !== 'undefined') {
                ModalEditAlienComponentCurrenAddress.init({ event: new Event('click') });
                return true;
            }
        } catch (e) {}
        return false;
    }""")
    if not opened:
        res["note"] = "ไม่พบปุ่มแก้ไขที่อยู่ที่ติดต่อได้"
        return res
    try:
        page.wait_for_selector("#btn_current_address_save", state="visible", timeout=12_000)
    except PWTimeoutError:
        res["note"] = "modal ที่อยู่ไม่แสดง"
        return res
    page.wait_for_timeout(700)

    _bt30_fill_text(page, "tbx_current_address_desc", rec.get("addr_no", ""))
    _bt30_fill_text(page, "tbx_current_address_village_building", rec.get("addr_moo", ""))
    _bt30_fill_text(page, "tbx_current_address_soi", rec.get("addr_soi", ""))
    _bt30_fill_text(page, "tbx_current_address_road", rec.get("addr_road", ""))

    # cascading: จังหวัด → เขต/อำเภอ → แขวง/ตำบล (ต้องรอ option โหลดก่อนเลือกตัวถัดไป)
    if rec.get("addr_prov"):
        if _select_option_by_text(page, "ddl_current_address_prov_id", rec["addr_prov"]):
            page.wait_for_timeout(1200)
        else:
            notes.append(f"จังหวัด '{rec['addr_prov']}'?")
    if rec.get("addr_dist"):
        _bt30_wait_select_option(page, "ddl_current_address_dist_id", rec["addr_dist"])
        if _select_option_by_text(page, "ddl_current_address_dist_id", rec["addr_dist"]):
            page.wait_for_timeout(1200)
        else:
            notes.append(f"อำเภอ '{rec['addr_dist']}'?")
    if rec.get("addr_subdist"):
        _bt30_wait_select_option(page, "ddl_current_address_subdist_id", rec["addr_subdist"])
        if not _select_option_by_text(page, "ddl_current_address_subdist_id", rec["addr_subdist"]):
            notes.append(f"ตำบล '{rec['addr_subdist']}'?")
        page.wait_for_timeout(800)

    page.evaluate(r"""() => { const b = document.getElementById('btn_current_address_save'); if (b) b.click(); }""")
    page.wait_for_timeout(1800)
    alert = _capture_register_alert(page)
    if alert:
        if any(k in alert for k in ("ไม่", "ผิดพลาด", "กรุณา")):
            notes.append(alert[:150])
        _close_register_alert(page)
    still = page.evaluate(
        r"""() => { const b = document.getElementById('btn_current_address_save'); return !!(b && b.offsetParent !== null); }"""
    )
    res["ok"] = not still
    res["note"] = ("; ".join(notes)) or ("บันทึกที่อยู่" if res["ok"] else "modal ยังเปิดหลังบันทึก")
    return res


def _bt30_fill_staypermit(page: Page, rec: dict[str, Any], log=print) -> dict[str, Any]:
    """2.3+2.4 เปิด modal 'ข้อมูลเพิ่มเติม' (#staypermitModal) → กรอกเอกสาร/วีซ่า/ตม. → บันทึก"""
    res = {"ok": False, "note": ""}
    notes: list[str] = []
    page.evaluate(r"""() => {
        try {
            if (typeof ModalEditAlienComponentStayPermit !== 'undefined')
                ModalEditAlienComponentStayPermit.init({ event: new Event('click'), form_type_id: 'MT_59', alien_emp_relate_id: '' });
        } catch (e) {}
    }""")
    page.wait_for_timeout(2500)
    # init เติมข้อมูลแต่ไม่ show modal → บังคับแสดง
    page.evaluate(r"""() => {
        if (window.jQuery) { try { jQuery('#staypermitModal').modal('show'); } catch (e) {} }
        const m = document.getElementById('staypermitModal');
        if (m) { m.classList.add('show'); m.style.display = 'block'; }
    }""")
    try:
        page.wait_for_selector("#btn_staypermit_save", state="visible", timeout=10_000)
    except PWTimeoutError:
        res["note"] = "modal 'ข้อมูลเพิ่มเติม' ไม่แสดง"
        return res
    page.wait_for_timeout(700)

    # ประเภทเอกสารแสดงการได้รับอนุญาต (Excel ไม่มีคอลัมน์ → default หนังสือเดินทาง ตามข้อมูล passport)
    doc_type = rec.get("stay_doc_type") or "หนังสือเดินทาง"
    if not _select_option_by_text(page, "stay-permission-type-selector", doc_type):
        notes.append(f"ประเภทเอกสาร '{doc_type}'?")
    page.wait_for_timeout(700)

    # เลขที่เอกสาร / สถานที่ออกให้ / ประเทศ (จากสัญชาติ) / วันที่
    _bt30_fill_text(page, "travelDocIdCheck", rec.get("doc_no", ""))
    _bt30_fill_text(page, "placeIssuanceDocTravel", rec.get("doc_place", ""))
    if rec.get("nationality"):
        _select_option_by_text(page, "countryTravelDocSelect", rec["nationality"])
    _bt30_fill_date(page, "dateOfStartTravelDoc", rec.get("doc_issue", ""))
    _bt30_fill_date(page, "dateOfTravelDocExpiry", rec.get("doc_expire", ""))

    # การตรวจลงตรา (วีซ่า)
    _bt30_fill_text(page, "surveillanceId", rec.get("visa_no", ""))
    if rec.get("visa_type") and not _select_option_by_text(page, "typeOfSurveillance", rec["visa_type"]):
        notes.append("ประเภทวีซ่า?")
    _bt30_fill_text(page, "surveillancePlace", rec.get("visa_place", ""))
    _bt30_fill_date(page, "surveillanceDate", rec.get("visa_issue", ""))
    _bt30_fill_date(page, "surveillanceExpire", rec.get("visa_expire", ""))

    # 2.4 ตรวจคนเข้าเมือง
    if rec.get("imm_office") and not _select_option_by_text(page, "arrivalInTheKingdomApproveTypeNew", rec["imm_office"]):
        notes.append(f"ด่าน/ตม. '{rec['imm_office']}'?")
    _bt30_fill_date(page, "arrivalInTheKingdomDate", rec.get("arrival_date", ""))
    _bt30_fill_date(page, "arrivalInTheKingdomExpire", rec.get("stay_until", ""))
    page.wait_for_timeout(500)

    page.evaluate(r"""() => { const b = document.getElementById('btn_staypermit_save'); if (b) b.click(); }""")
    page.wait_for_timeout(2200)
    alert = _capture_register_alert(page)
    if alert:
        if any(k in alert for k in ("ไม่", "ผิดพลาด", "กรุณา")):
            notes.append(alert[:150])
        _close_register_alert(page)
    still = page.evaluate(
        r"""() => { const b = document.getElementById('btn_staypermit_save'); return !!(b && b.offsetParent !== null); }"""
    )
    res["ok"] = not still
    res["note"] = ("; ".join(notes)) or ("บันทึกข้อมูลเพิ่มเติม" if res["ok"] else "modal ยังเปิดหลังบันทึก")
    return res


def _bt30_step2_next(page: Page, log=print) -> dict[str, Any]:
    """2.5 กดปุ่ม 'ถัดไป' (#validateRenewcheck) บนหน้ารายละเอียด (หยุดที่ขอบเขตนี้ — ไม่ส่งคำขอจริง)"""
    res = {"ok": False, "note": ""}
    _bt30_dismiss_news(page, log=log)
    # ติ๊ก checkbox ยืนยันข้อมูล (จำเป็นต่อการกดถัดไป) ถ้ามีและมองเห็น
    page.evaluate(r"""() => {
        const cb = document.getElementById('check_truth');
        if (cb && cb.offsetParent !== null && !cb.checked) {
            cb.click();
            if (!cb.checked) { cb.checked = true; cb.dispatchEvent(new Event('change', { bubbles: true })); }
        }
    }""")
    page.wait_for_timeout(400)
    clicked = page.evaluate(r"""() => {
        const b = document.getElementById('validateRenewcheck');
        if (b) { b.click(); return true; }
        const b2 = Array.from(document.querySelectorAll('button,a'))
          .find(x => /ถัดไป/.test((x.textContent || '').trim()) && /btn-next|action-button/.test(x.className || ''));
        if (b2) { b2.click(); return true; }
        return false;
    }""")
    if not clicked:
        res["note"] = "ไม่พบปุ่ม 'ถัดไป' (2.5 / #validateRenewcheck)"
        return res
    page.wait_for_timeout(2500)
    alert = _capture_register_alert(page)
    res["ok"] = True
    res["note"] = (alert[:200] if alert else "กดถัดไป (2.5) แล้ว")
    return res


def _bt30_select2_by_prefix(page: Page, id_prefix: str, want: str, fuzzy: bool = True) -> dict[str, Any]:
    """เลือก option ของ <select> (select2) ที่ id ขึ้นต้นด้วย id_prefix (id เป็นแบบไดนามิก)
    ลำดับการจับคู่: ค่าปัจจุบันตรงอยู่แล้ว → ตรงเป๊ะ → contains/reverse-contains
    ถ้า fuzzy=True เพิ่มการเทียบ 'ป้ายนำหน้า' ก่อน ':' (เช่น สำนักงาน/สาขา) + ตัวเลขในข้อความ
    + คำที่ซ้ำกัน และถ้าค่าปัจจุบันมีป้ายนำหน้าตรงกับที่ต้องการอยู่แล้ว จะคงค่าเดิม
    (ระบบ preselect ไว้ถูกต้อง แม้ข้อความใน Excel จะพิมพ์ต่างเล็กน้อย)
    คืน {ok, kept?, text?, reason?}
    """
    return page.evaluate(
        r"""(args) => {
          const { prefix, want, fuzzy } = args;
          const s = document.querySelector(`select[id^="${prefix}"]`);
          if (!s) return { ok:false, reason:'no-select' };
          const norm = x => String(x||'').replace(/\s+/g,'').toLowerCase();
          const w = norm(want);
          if (!w) return { ok:false, reason:'empty-want' };
          const isPh = t => /กรุณาเลือก/.test(t||'');
          const opts = Array.from(s.options).filter(o => !isPh(o.textContent||''));
          const cur = s.options[s.selectedIndex] || null;
          const apply = (opt) => {
            try {
              s.value = opt.value;
              if (window.jQuery) { jQuery(s).val(opt.value).trigger('change'); }
              s.dispatchEvent(new Event('change', { bubbles:true }));
            } catch(e) {}
            return { ok:true, text:(opt.textContent||'').trim() };
          };
          if (cur && !isPh(cur.textContent||'') && norm(cur.textContent) === w)
            return { ok:true, kept:true, text:(cur.textContent||'').trim() };
          let opt = opts.find(o => norm(o.textContent) === w);
          if (opt) return apply(opt);
          opt = opts.find(o => norm(o.textContent).includes(w)
                            || (w.includes(norm(o.textContent)) && norm(o.textContent).length >= 3));
          if (opt) return apply(opt);
          if (!fuzzy) {
            if (cur && !isPh(cur.textContent||'') && s.value)
              return { ok:true, kept:true, text:(cur.textContent||'').trim() };
            return { ok:false, reason:'no-match' };
          }
          const lead = x => norm(String(x||'').split(':')[0]);
          const digs = x => (String(x||'').match(/\d+/g) || []);
          const words = x => (norm(x).match(/[ก-๙a-z0-9]{2,}/g) || []);
          const wLead = lead(want), wDig = digs(want), wWords = words(want);
          if (cur && !isPh(cur.textContent||'') && s.value && wLead && lead(cur.textContent) === wLead)
            return { ok:true, kept:true, text:(cur.textContent||'').trim() };
          const score = (o) => {
            const t = o.textContent || '';
            let sc = 0;
            if (wLead && lead(t) === wLead) sc += 100;
            const od = digs(t);
            sc += od.filter(d => wDig.includes(d)).length * 10;
            const ow = words(t);
            sc += ow.filter(x => wWords.includes(x)).length;
            return sc;
          };
          let best=null, bestSc=-1;
          for (const o of opts) { const sc = score(o); if (sc > bestSc) { bestSc = sc; best = o; } }
          if (best && bestSc >= 100) return apply(best);
          if (cur && !isPh(cur.textContent||'') && s.value)
            return { ok:true, kept:true, text:(cur.textContent||'').trim() };
          if (best && bestSc > 0) return apply(best);
          return { ok:false, reason:'no-fuzzy' };
        }""",
        {"prefix": id_prefix, "want": want, "fuzzy": fuzzy},
    )


def _bt30_fill_page2(page: Page, rec: dict[str, Any], log=print) -> dict[str, Any]:
    """หน้า 2/2 ของ 'กรอกข้อมูลคำขอ':
    2.2.1 หัวข้อ 'โดยจะมาทำงาน' → สถานที่ทำงาน/สาขา + ประเภทกิจการ (select2 id ไดนามิก)
    2.2.2 หัวข้อ 'เอกสารแสดงการอนุญาตหรือการรับรอง' → เลขที่, ออกให้โดย, วันที่ออกเอกสาร, วันที่เอกสารหมดอายุ
    คืน {ok, note}
    """
    res = {"ok": False, "note": ""}
    notes: list[str] = []
    try:
        _bt30_dismiss_news(page, log=log)
        page.wait_for_timeout(500)

        # 2.2.1 สถานที่ทำงาน/สาขา (select2)
        want_place = rec.get("work_place", "")
        if want_place:
            r = _bt30_select2_by_prefix(page, "workplace_full_address_emp_", want_place, fuzzy=True)
            tag = "คงค่าเดิม" if r.get("kept") else ("OK" if r.get("ok") else "X")
            notes.append(f"สถานที่ทำงาน:{tag}")
            page.wait_for_timeout(1000)  # เผื่อ cascade โหลดประเภทกิจการใหม่

        # ประเภทกิจการ (ขึ้นกับสถานที่ทำงาน)
        want_biz = rec.get("work_biz", "")
        if want_biz:
            rb = _bt30_select2_by_prefix(page, "bus_type_name_th_emp_", want_biz, fuzzy=False)
            tag = "คงค่าเดิม" if rb.get("kept") else ("OK" if rb.get("ok") else "X")
            notes.append(f"ประเภทกิจการ:{tag}")

        # 2.2.2 เอกสารแสดงการอนุญาตหรือการรับรอง (text + datepicker)
        if rec.get("auth_no"):
            _bt30_fill_text(page, "alien_cert_no", rec["auth_no"])
        if rec.get("auth_by"):
            _bt30_fill_text(page, "alien_cert_issue_at", rec["auth_by"])
        if rec.get("auth_issue"):
            _bt30_fill_date(page, "alien_cert_issue_dt", rec["auth_issue"])
        if rec.get("auth_expire"):
            _bt30_fill_date(page, "alien_cert_expired_dt", rec["auth_expire"])
        notes.append("เอกสารอนุญาต:กรอกแล้ว")
        page.wait_for_timeout(400)

        res["ok"] = True
        res["note"] = " | ".join(notes)
        return res
    except Exception as e:
        res["note"] = (" | ".join(notes) + " | " + str(e))[:300]
        return res


def _bt30_page2_next(page: Page, log=print) -> dict[str, Any]:
    """หน้า 2/2: กดปุ่ม 'ถัดไป' (#NextStepTwoPageOne) ไปขั้น 'แนบเอกสาร'
    (หยุดที่ขอบเขตนี้ — ยังไม่แนบไฟล์/ไม่ส่งคำขอ/ไม่ชำระเงิน)
    """
    res = {"ok": False, "note": ""}
    _bt30_dismiss_news(page, log=log)
    clicked = page.evaluate(r"""() => {
        const b = document.getElementById('NextStepTwoPageOne');
        if (b) { b.click(); return true; }
        const b2 = Array.from(document.querySelectorAll('button,a'))
          .find(x => /ถัดไป/.test((x.textContent || '').trim())
                  && /NextStep|btn-next|action-button/.test((x.className || '') + (x.id || '')));
        if (b2) { b2.click(); return true; }
        return false;
    }""")
    if not clicked:
        res["note"] = "ไม่พบปุ่ม 'ถัดไป' (#NextStepTwoPageOne)"
        return res
    page.wait_for_timeout(2800)
    alert = _capture_register_alert(page)
    res["ok"] = True
    res["note"] = (alert[:200] if alert else "กดถัดไป (หน้า 2/2) แล้ว")
    return res


# ---- ขั้นตอน 3: แนบเอกสาร (attach documents) ----------------------------------
# JS: หา file input จริงของแต่ละประเภทเอกสาร (id=file_name_NN, onchange=selectFileupload)
# โดยจับคู่จากข้อความของแถว (row label) กับ keyword ของแต่ละช่อง (กันชน group-alt 34/35)
_BT30_RESOLVE_DOC_JS = r"""(slots) => {
  const txt = e => (e.textContent || '').replace(/\s+/g, ' ').trim();
  const norm = x => String(x || '').replace(/\s+/g, '').toLowerCase();
  const inputs = Array.from(document.querySelectorAll('input[type=file]'))
    .filter(e => /selectFileupload/.test(e.getAttribute('onchange') || ''))
    .map(e => {
      let row = e, best = '';
      for (let up = 0; up < 8 && row; up++) {
        row = row.parentElement;
        if (row) {
          const t = txt(row);
          if (t.length > 10 && /เอกสาร|สำเนา|ใบ|รูปถ่าย|หนังสือ|บัตร|อนุญาต|บต\.|มอบอำนาจ|อากร|เวชกรรม/.test(t)) {
            best = t; break;
          }
        }
      }
      return { id: e.id, rowText: best || txt(e.parentElement || e) };
    });
  const used = new Set();
  const out = {};
  for (const [key, kws] of slots) {
    let found = '';
    for (const inp of inputs) {
      if (used.has(inp.id)) continue;
      const rt = norm(inp.rowText);
      if (kws.some(k => rt.includes(norm(k)))) { found = inp.id; used.add(inp.id); break; }
    }
    out[key] = found;
  }
  return out;
}"""

# ลำดับช่องเอกสาร (key ใน rec, keyword จับแถว, ป้ายแสดงผล)
_BT30_DOC_SLOTS: list[tuple[str, list[str], str]] = [
    ("doc_passport", ["สำเนาหนังสือเดินทาง"], "3.1 สำเนาหนังสือเดินทาง"),
    ("doc_entry", ["หลักฐานการอนุญาตให้เข้ามาในราชอาณาจักร"], "3.2 หลักฐานการอนุญาตเข้าราชอาณาจักร"),
    ("doc_contract", ["สำเนาสัญญาจ้าง"], "3.3 สำเนาสัญญาจ้าง"),
    ("doc_medical", ["เวชกรรม", "ใบรับรองแพทย์"], "3.4 ใบรับรองแพทย์"),
    ("doc_photo", ["รูปถ่าย"], "3.5 รูปถ่าย 3x4"),
    ("doc_bt46", ["บต.46"], "3.6 บต.46"),
    ("doc_poa_agent", ["ดำเนินการแทน"], "3.7 หนังสือมอบอำนาจผู้ดำเนินการแทน"),
    ("doc_poa_stamp", ["อากรแสตมป์"], "3.8 ใบมอบอำนาจติดอากรแสตมป์"),
    ("doc_id_grantor", ["ประชาชนของผู้มอบอำนาจ"], "3.9 บัตร ปชช.ผู้มอบอำนาจ"),
    ("doc_id_grantee", ["ประชาชนของผู้รับมอบอำนาจ"], "4.0 บัตร ปชช.ผู้รับมอบอำนาจ"),
    ("doc_workpermit", ["ใบอนุญาตทำงาน"], "4.1 ใบอนุญาตทำงาน"),
]

# ── ขนาดไฟล์แนบสูงสุดที่ระบบ e-WorkPermit ยอมรับ (ต่อ 1 ไฟล์) ──
_BT30_MAX_DOC_MB: float = 4.0


def _bt30_collect_doc_paths(rec: dict[str, Any]) -> list[tuple[str, str]]:
    """รวม (label, path) ของไฟล์เอกสารทั้งหมดที่จะแนบของคนต่างด้าว 1 แถว
    (3.1–4.1 ตาม _BT30_DOC_SLOTS + 4.2 เอกสารอื่นๆ) — เฉพาะช่องที่กรอก path ไว้
    """
    items: list[tuple[str, str]] = []
    for key, _kws, label in _BT30_DOC_SLOTS:
        p = (rec.get(key) or "").strip()
        if p:
            items.append((label, p))
    for i, p in enumerate((rec.get("doc_others") or []), start=1):
        if p:
            items.append((f"4.2 เอกสารอื่นๆ #{i}", p))
    return items


def _bt30_oversized_docs(
    rec: dict[str, Any], limit_mb: float = _BT30_MAX_DOC_MB
) -> list[dict[str, Any]]:
    """ตรวจไฟล์เอกสารของ 1 แถวที่ 'มีอยู่จริง แต่ขนาดเกิน' limit (default 4 MB)
    คืนรายการ {label, name, mb, path} (ไฟล์ที่หาไม่พบจะข้าม — ให้ขั้นแนบรายงาน 'ไม่พบไฟล์' เอง)
    """
    limit = int(limit_mb * 1024 * 1024)
    out: list[dict[str, Any]] = []
    for label, path in _bt30_collect_doc_paths(rec):
        try:
            p = Path(path)
            if p.is_file():
                sz = p.stat().st_size
                if sz > limit:
                    out.append({"label": label, "name": p.name,
                                "mb": round(sz / 1024 / 1024, 2), "path": str(p)})
        except OSError:
            pass
    return out


def _bt30_preflight_doc_sizes(
    records: list[dict[str, Any]], log=print, limit_mb: float = _BT30_MAX_DOC_MB
) -> list[dict[str, Any]]:
    """สแกนขนาดไฟล์เอกสารของทุกแถว 'ก่อน' เริ่มทำงานจริง — รายงานไฟล์ที่เกิน limit
    คืนรายการปัญหา [{row, seq, name, label, file, mb}] (ว่าง = ผ่านหมด)
    เรียกตอนต้น run_bt30 เพื่อให้ผู้ใช้แก้ไฟล์ก่อนยื่น (ระบบจำกัด 4 MB/ไฟล์)
    """
    problems: list[dict[str, Any]] = []
    for rec in records:
        for ov in _bt30_oversized_docs(rec, limit_mb):
            problems.append({
                "row": rec.get("row_index"), "seq": rec.get("seq"),
                "name": rec.get("name", ""), "label": ov["label"],
                "file": ov["name"], "mb": ov["mb"],
            })
    if problems:
        log(f"  ⚠ ตรวจขนาดไฟล์แนบ: พบ {len(problems)} ไฟล์เกิน {limit_mb:.0f} MB "
            f"(ระบบจำกัดไม่เกิน {limit_mb:.0f} MB/ไฟล์ — ต้องย่อก่อนยื่น):")
        for p in problems:
            log(f"      • แถว {p['row']} {p['name']} | {p['label']}: "
                f"{p['file']} = {p['mb']} MB")
    else:
        log(f"  ✓ ตรวจขนาดไฟล์แนบ: ทุกไฟล์ ≤ {limit_mb:.0f} MB")
    return problems


def _bt30_handle_crop_modal(page: Page, log=print) -> bool:
    """ช่องรูปภาพที่ต้องครอป (เช่น 3.5 รูปถ่าย 3x4, doc id=42) เมื่อแนบไฟล์ภาพจะเด้งโมดอล
    #crop-modal (Cropper.js) ขึ้นมา ต้องตั้งกรอบให้ครอบเต็มรูป (อัตราส่วน 3:4) แล้วกด 'บันทึก'
    (#save-crop-button) มิฉะนั้นโมดอลจะค้างและบล็อกการบันทึกเอกสารถัดไป (เช่น 4.2)
    คืน True ถ้าพบและจัดการโมดอลครอปแล้ว, False ถ้าไม่มีโมดอลครอป
    """
    # 1) รอดูว่ามีโมดอลครอปเด้งขึ้นไหม (FileReader อ่านภาพ → $('#crop-modal').modal('show'))
    appeared = False
    for _ in range(12):  # รอสูงสุด ~6 วินาที
        try:
            shown = bool(page.evaluate(r"""() => {
                const m = document.getElementById('crop-modal');
                return !!(m && (m.classList.contains('show')
                    || (getComputedStyle(m).display !== 'none' && m.offsetParent !== null)));
            }"""))
        except Exception:
            shown = False
        if shown:
            appeared = True
            break
        page.wait_for_timeout(500)
    if not appeared:
        return False
    # 2) รอ Cropper พร้อม (global cropper + รูปโหลดเสร็จ)
    try:
        page.wait_for_function(r"""() => {
            const img = document.getElementById('crop-image');
            return !!(window.cropper && img && img.complete && img.naturalWidth > 0);
        }""", timeout=8000)
    except PWTimeoutError:
        pass
    page.wait_for_timeout(500)
    # 3) ตั้งกรอบครอปให้ครอบพื้นที่ 3:4 ใหญ่สุด กึ่งกลางรูป (= ครอปภาพให้พอดีทั้งรูป)
    try:
        page.evaluate(r"""() => {
            const cr = window.cropper;
            if (!cr || typeof cr.getCanvasData !== 'function') return;
            const c = cr.getCanvasData();
            const ar = 3 / 4;            // กว้าง:สูง = 3:4
            let w = c.width, h = w / ar; // h = w * 4/3
            if (h > c.height) { h = c.height; w = h * ar; }
            cr.setCropBoxData({
                left: c.left + (c.width - w) / 2,
                top: c.top + (c.height - h) / 2,
                width: w,
                height: h,
            });
        }""")
    except Exception:
        pass
    page.wait_for_timeout(300)
    # 4) กดบันทึก (#save-crop-button → getCroppedCanvas().toBlob → modal hide + resolve)
    try:
        page.evaluate(r"""() => { const b = document.getElementById('save-crop-button'); if (b) b.click(); }""")
    except Exception:
        pass
    # 5) รอโมดอลครอปปิด (toBlob เป็น async)
    closed = False
    for _ in range(20):  # รอสูงสุด ~10 วินาที
        page.wait_for_timeout(500)
        try:
            closed = bool(page.evaluate(r"""() => {
                const m = document.getElementById('crop-modal');
                return !m || !(m.classList.contains('show'));
            }"""))
        except Exception:
            closed = False
        if closed:
            break
    page.wait_for_timeout(600)
    log("        \u2713 3.5 ครอปรูปถ่าย (3:4) แล้วบันทึก")
    return True


def _bt30_upload_doc(page: Page, input_id: str, abs_path: str, log=print) -> dict[str, Any]:
    """แนบไฟล์ลงช่องเอกสาร 1 ช่อง (set_input_files บน input ซ่อน → AJAX อัปโหลด)
    ยืนยันสำเร็จเมื่อปุ่มลบ closefilerequest("NN",...) ปรากฏ
    หมายเหตุ: ช่องรูปภาพ (เช่น 3.5 รูปถ่าย) จะเด้งโมดอลครอป → จัดการผ่าน _bt30_handle_crop_modal
    """
    res = {"ok": False, "note": ""}
    p = Path(abs_path)
    if not abs_path or not p.exists():
        res["note"] = f"ไม่พบไฟล์ ({p.name if abs_path else 'ว่าง'})"
        return res
    try:
        sz = p.stat().st_size
        if sz > _BT30_MAX_DOC_MB * 1024 * 1024:
            res["note"] = (f"ไฟล์เกิน {_BT30_MAX_DOC_MB:.0f}MB "
                           f"({sz / 1024 / 1024:.2f}MB) — ย่อก่อนยื่น: {p.name}")
            return res
    except OSError:
        pass
    num = input_id.replace("file_name_", "")
    needle = f'closefilerequest("{num}"'
    try:
        page.set_input_files(f"#{input_id}", str(p))
    except Exception as e:
        res["note"] = f"แนบไฟล์ล้มเหลว: {e}"[:120]
        return res
    # ถ้าเป็นไฟล์ภาพ อาจเด้งโมดอลครอป (3.5 รูปถ่าย 3x4) → ตั้งกรอบ 3:4 เต็มรูปแล้วกดบันทึก
    cropped = False
    if p.suffix.lower() in (".jpg", ".jpeg", ".png"):
        cropped = _bt30_handle_crop_modal(page, log=log)
    done = False
    for _ in range(40):  # รอสูงสุด ~24 วินาที
        page.wait_for_timeout(600)
        try:
            done = bool(page.evaluate(
                r"""(needle) => Array.from(document.querySelectorAll('[onclick]'))
                      .some(e => (e.getAttribute('onclick') || '').includes(needle))""",
                needle,
            ))
        except Exception:
            done = False
        if done:
            break
    if not done and cropped:
        # ยืนยันอีกทาง: creatfileCrop เซ็ตไฟล์เข้า #file_name_NN แล้วหรือยัง
        try:
            done = bool(page.evaluate(
                "(id) => { const e = document.getElementById('file_name_' + id); return !!(e && e.files && e.files.length > 0); }",
                num,
            ))
        except Exception:
            pass
    res["ok"] = done
    res["note"] = (f"{p.name}" if done else f"อัปโหลดไม่ยืนยัน ({p.name})")
    return res


def _bt30_attach_others(page: Page, paths: list[str], log=print) -> dict[str, Any]:
    """4.2 เอกสารอื่นๆที่เกี่ยวข้อง — แนบทีละไฟล์ผ่านโมดอล 'เพิ่มเอกสาร' (#AddFileOrther)
    ต่อ 1 ไฟล์: เปิดโมดอล → แนบไฟล์ที่ #file_other_N (onchange=AddNewFileOrther) →
    รอ reader.onload เซ็ต #file_name_select_N → กรอกชื่อเอกสาร (#text_file_name_00/_N) →
    ยืนยัน (SaveAddFileOrther) → โมดอลปิด
    ยืนยันสำเร็จเมื่อรายการ '#ShowtxtBox_File .checklength' เพิ่มขึ้น
    """
    res = {"ok": False, "note": "", "n_ok": 0}
    valid = [p for p in paths if p and Path(p).exists()]
    n_ok = 0
    for idx, ap in enumerate(paths, 1):
        p = Path(ap)
        if not ap or not p.exists():
            log(f"        ✗ 4.2 ไฟล์ที่ {idx}: ไม่พบไฟล์ ({p.name if ap else 'ว่าง'})")
            continue
        try:
            if p.stat().st_size > _BT30_MAX_DOC_MB * 1024 * 1024:
                log(f"        ✗ 4.2 ไฟล์ที่ {idx}: เกิน {_BT30_MAX_DOC_MB:.0f}MB "
                    f"({p.stat().st_size / 1024 / 1024:.2f}MB) — ย่อก่อนยื่น ({p.name})")
                continue
        except OSError:
            pass
        try:
            n0 = int(page.evaluate(
                "() => document.querySelectorAll('.checklength').length"))
            opened = page.evaluate(r"""() => {
                const b = Array.from(document.querySelectorAll('button,a,label'))
                  .find(e => /เพิ่มเอกสาร/.test((e.textContent || '').trim())
                        || /OncOpenModalAddFile\(/.test(e.getAttribute('onclick') || ''));
                if (b) { b.click(); return true; }
                return false;
            }""")
            if not opened:
                log("        ✗ 4.2 ไม่พบปุ่ม 'เพิ่มเอกสาร'")
                break
            try:
                page.wait_for_selector("#AddFileOrther.show", timeout=8000)
            except PWTimeoutError:
                page.wait_for_timeout(1500)

            # หา input ที่กำลังรอแนบ (onchange=AddNewFileOrther) → ได้เลขช่อง N
            try:
                page.wait_for_function(r"""() =>
                  Array.from(document.querySelectorAll('#AddFileOrther input[type=file]'))
                    .some(e => /AddNewFileOrther/.test(e.getAttribute('onchange') || ''))
                """, timeout=6000)
            except PWTimeoutError:
                pass
            fid = page.evaluate(r"""() => {
                const list = Array.from(document.querySelectorAll('#AddFileOrther input[type=file]'))
                  .filter(e => /AddNewFileOrther/.test(e.getAttribute('onchange') || ''));
                return list.length ? list[list.length - 1].id : '';
            }""")
            if not fid:
                log(f"        ✗ 4.2 ไฟล์ที่ {idx}: ไม่พบช่องแนบไฟล์ในโมดอล")
                page.evaluate("() => { if (window.jQuery) { try { jQuery('#AddFileOrther').modal('hide'); } catch (e) {} } }")
                page.wait_for_timeout(600)
                continue
            num = fid.replace("file_other_", "")
            text_id = "text_file_name_00" if num == "0" else f"text_file_name_{num}"

            # แนบไฟล์ที่ช่องเฉพาะ (ทริกเกอร์ onchange=AddNewFileOrther)
            page.set_input_files(f"#{fid}", str(p))

            # รอ reader.onload เซ็ต #file_name_select_N (= อัปโหลดไฟล์เข้า hidden_file_up เสร็จ)
            got_sel = False
            for _ in range(20):  # สูงสุด ~10 วินาที
                v = page.evaluate(
                    "(id) => { const e = document.getElementById(id); return e ? (e.value || '') : ''; }",
                    f"file_name_select_{num}")
                if v:
                    got_sel = True
                    break
                page.wait_for_timeout(500)
            if not got_sel:
                log(f"        ✗ 4.2 ไฟล์ที่ {idx}: อัปโหลดไม่เสร็จ (file_name_select ว่าง) ({p.name})")
                page.evaluate("() => { if (window.jQuery) { try { jQuery('#AddFileOrther').modal('hide'); } catch (e) {} } }")
                page.wait_for_timeout(600)
                continue

            # กรอกชื่อเอกสาร (ใช้ id ที่ถูกต้อง — onload เซ็ตชื่อไฟล์ไว้แล้ว, เขียนทับเป็นชื่อสะอาด)
            page.evaluate(r"""(args) => {
                const t = document.getElementById(args.id);
                if (t) {
                  t.value = args.name;
                  t.dispatchEvent(new Event('input', { bubbles: true }));
                  t.dispatchEvent(new Event('change', { bubbles: true }));
                }
            }""", {"id": text_id, "name": p.stem})
            page.wait_for_timeout(300)

            # ยืนยัน — เรียก SaveAddFileOrther() ตรงๆ (หน้านี้มีปุ่ม #buttonSaveFileSelect ซ้ำ 2 ปุ่ม
            # การ getElementById().click() จะโดนปุ่มแรกที่ไม่ผูก onclick → ฟังก์ชันไม่ทำงาน)
            # ฟังก์ชันอ่าน $("#file_other_0")/$("#text_file_name_00") ซึ่งมีชุดเดียว จึงเรียกตรงได้ปลอดภัย
            page.evaluate(r"""() => {
                try {
                    if (typeof window.SaveAddFileOrther === 'function') { window.SaveAddFileOrther(); return 'called'; }
                } catch (e) { return 'ERR:' + e; }
                const b = document.getElementById('buttonSaveFileSelect');
                if (b) { b.click(); return 'clicked'; }
                return 'none';
            }""")
            page.wait_for_timeout(500)

            # ยืนยันสำเร็จเมื่อจำนวน .checklength (getElementDoc) เพิ่มขึ้น
            ok_added = False
            for _ in range(16):  # รอรายการเพิ่มขึ้นสูงสุด ~8 วินาที
                try:
                    n1 = int(page.evaluate(
                        "() => document.querySelectorAll('.checklength').length"))
                except Exception:
                    n1 = n0
                if n1 > n0:
                    ok_added = True
                    break
                page.wait_for_timeout(500)

            # ปิดโมดอลถ้ายังเปิดอยู่
            page.evaluate(
                "() => { if (window.jQuery) { try { jQuery('#AddFileOrther').modal('hide'); } catch (e) {} } }")
            page.wait_for_timeout(900)

            if ok_added:
                n_ok += 1
                log(f"        ✓ 4.2 ไฟล์ที่ {idx}: {p.name}")
            else:
                log(f"        ✗ 4.2 ไฟล์ที่ {idx}: ไม่ยืนยัน ({p.name})")
        except Exception as e:
            log(f"        ✗ 4.2 ไฟล์ที่ {idx} ผิดพลาด: {str(e)[:80]}")
            try:
                page.evaluate(
                    "() => { if (window.jQuery) { try { jQuery('#AddFileOrther').modal('hide'); } catch (e) {} } }")
            except Exception:
                pass
            page.wait_for_timeout(600)
    res["n_ok"] = n_ok
    res["ok"] = bool(valid) and n_ok == len(valid)
    res["note"] = f"{n_ok}/{len(valid)} ไฟล์"
    return res



def _bt30_fill_attachments(page: Page, rec: dict[str, Any], log=print) -> dict[str, Any]:
    """แนบเอกสารครบทุกประเภทตาม path ใน Excel (3.1–4.1 + 4.2 เอกสารอื่นๆ)
    คืน {ok, note}
    """
    res = {"ok": False, "note": ""}
    notes: list[str] = []
    try:
        _bt30_dismiss_news(page, log=log)
        page.wait_for_timeout(600)

        mapping = page.evaluate(
            _BT30_RESOLVE_DOC_JS, [[key, kws] for key, kws, _ in _BT30_DOC_SLOTS])

        n_ok = 0
        n_try = 0
        for key, _kws, label in _BT30_DOC_SLOTS:
            path = (rec.get(key) or "").strip()
            if not path:
                continue
            n_try += 1
            input_id = mapping.get(key) or ""
            if not input_id:
                notes.append(f"{label}:ไม่พบช่อง")
                log(f"        ✗ {label} — ไม่พบช่องอัปโหลดบนหน้า")
                continue
            r = _bt30_upload_doc(page, input_id, path, log=log)
            if r["ok"]:
                n_ok += 1
            notes.append(f"{label}:{'OK' if r['ok'] else 'X'}")
            log(f"        {'✓' if r['ok'] else '✗'} {label} — {r['note']}")
            page.wait_for_timeout(400)

        # 4.2 เอกสารอื่นๆที่เกี่ยวข้อง (แนบทีละไฟล์)
        others = rec.get("doc_others") or []
        others = [o for o in others if o]
        if others:
            ro = _bt30_attach_others(page, others, log=log)
            n_ok += ro.get("n_ok", 0)
            n_try += len([o for o in others if Path(o).exists()])
            notes.append(f"4.2 เอกสารอื่นๆ:{ro['note']}")
            log(f"      {'✓' if ro['ok'] else '⚠'} 4.2 เอกสารอื่นๆที่เกี่ยวข้อง — {ro['note']}")

        res["ok"] = n_try > 0 and n_ok == n_try
        res["note"] = f"แนบ {n_ok}/{n_try} | " + " ".join(notes)
        return res
    except Exception as e:
        res["note"] = (" ".join(notes) + " | " + str(e))[:400]
        return res


def _bt30_attach_next(page: Page, log=print) -> dict[str, Any]:
    """แนบเอกสาร: กดปุ่ม 'ถัดไป' (#nextstepcheckfile) เพื่อไปขั้นถัดไป
    (หยุดที่ขอบเขตนี้ — ไม่ส่งคำขอ/ไม่ยืนยันตัวตน/ไม่ชำระเงิน)
    """
    res = {"ok": False, "note": ""}
    _bt30_dismiss_news(page, log=log)
    clicked = page.evaluate(r"""() => {
        const b = document.getElementById('nextstepcheckfile');
        if (b) { b.click(); return true; }
        const b2 = Array.from(document.querySelectorAll('button,a'))
          .find(x => /ถัดไป/.test((x.textContent || '').trim())
                  && /next|step/i.test((x.id || '') + (x.className || '')));
        if (b2) { b2.click(); return true; }
        return false;
    }""")
    if not clicked:
        res["note"] = "ไม่พบปุ่ม 'ถัดไป' (#nextstepcheckfile)"
        return res
    page.wait_for_timeout(2800)
    alert = _capture_register_alert(page)
    res["ok"] = True
    res["note"] = (alert[:200] if alert else "กดถัดไป (แนบเอกสาร) แล้ว")
    return res


def _bt30_click_next_generic(
    page: Page, prefer_id: str | None, label: str, log=print, wait_ms: int = 3000
) -> dict[str, Any]:
    """กดปุ่ม 'ถัดไป' แบบทั่วไป — ใช้ id ที่ระบุก่อน (ถ้ามองเห็น+ใช้งานได้)
    ไม่งั้นเลือกปุ่ม 'ถัดไป' ที่มองเห็น+ใช้งานได้ตัวสุดท้ายบนหน้า
    (ใช้กับ Step 4 = เอกสารนายจ้าง 2/2 และ Step 5.1 = สรุปคำขอ 1/2 ที่ปุ่มไม่มี id)
    """
    res = {"ok": False, "note": ""}
    _bt30_dismiss_news(page, log=log)
    clicked = page.evaluate(
        r"""(pid) => {
        const vis = el => { const r = el.getBoundingClientRect(); const s = getComputedStyle(el);
            return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden'; };
        if (pid) { const b = document.getElementById(pid); if (b && !b.disabled && vis(b)) { b.click(); return 'id'; } }
        const cand = Array.from(document.querySelectorAll('button,a'))
            .filter(e => vis(e) && !e.disabled && /ถัดไป/.test((e.textContent || '').trim())
                && (e.textContent || '').trim().length < 20
                && !/ย้อนกลับ|ยกเลิก/.test((e.textContent || '').trim()));
        if (cand.length) { cand[cand.length - 1].click(); return 'generic'; }
        return '';
    }""",
        prefer_id,
    )
    if not clicked:
        res["note"] = f"ไม่พบปุ่มถัดไป ({label})"
        return res
    page.wait_for_timeout(wait_ms)
    _bt30_dismiss_news(page, log=log)
    res["ok"] = True
    res["note"] = f"กดถัดไป ({label}) แล้ว [{clicked}]"
    return res


def _bt30_step5_confirm_next(page: Page, log=print) -> dict[str, Any]:
    """Step 5.2 (สรุปคำขอ หน้า 2/2): ติ๊ก checkbox 'ข้าพเจ้าได้ตรวจสอบข้อมูล...ถูกต้อง' (#check_truth)
    แล้วกดถัดไป (#nextstepfour59 — จะ enable หลังติ๊ก)
    """
    res = {"ok": False, "note": ""}
    _bt30_dismiss_news(page, log=log)
    ck = page.evaluate(
        r"""() => {
        const c = document.getElementById('check_truth');
        if (!c) return 'no-checkbox';
        if (!c.checked) { c.click(); if (!c.checked) { c.checked = true; c.dispatchEvent(new Event('change', {bubbles:true})); } }
        return c.checked ? 'checked' : 'fail';
    }"""
    )
    if ck != "checked":
        res["note"] = f"ติ๊กยืนยันข้อมูลไม่ได้ ({ck})"
        return res
    page.wait_for_timeout(800)
    # รอปุ่มถัดไป enable
    for _ in range(12):
        st = page.evaluate(
            "() => { const b = document.getElementById('nextstepfour59'); "
            "return b ? (b.disabled ? 'disabled' : 'enabled') : 'none'; }"
        )
        if st in ("enabled", "none"):
            break
        page.wait_for_timeout(400)
    nx = page.evaluate(
        r"""() => {
        const b = document.getElementById('nextstepfour59');
        if (b && !b.disabled) { b.click(); return 'id'; }
        const vis = el => { const r = el.getBoundingClientRect(); const s = getComputedStyle(el);
            return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden'; };
        const cand = Array.from(document.querySelectorAll('button,a'))
            .filter(e => vis(e) && !e.disabled && /ถัดไป/.test((e.textContent || '').trim())
                && (e.textContent || '').trim().length < 20
                && !/ย้อนกลับ|ยกเลิก/.test((e.textContent || '').trim()));
        if (cand.length) { cand[cand.length - 1].click(); return 'generic'; }
        return '';
    }"""
    )
    if not nx:
        res["note"] = "ติ๊กยืนยันข้อมูลแล้วแต่กดถัดไปไม่ได้"
        return res
    page.wait_for_timeout(3500)
    _bt30_dismiss_news(page, log=log)
    res["ok"] = True
    res["note"] = f"ติ๊กยืนยันข้อมูล + กดถัดไป แล้ว [{nx}]"
    return res


def _bt30_step6_upload_identity(page: Page, log=print) -> dict[str, Any]:
    """Step 6.1 (วิธีการยืนยันตัวตน): กด 'อัปโหลดภาพ' (OncOpenModalAddFileIden) → โมดอล #AddFileIden
    → กด 'บันทึก' โดยไม่แนบไฟล์จริง (เรียก SaveAddFileIden() → identify() ซึ่ง enable ปุ่มถัดไป
    #button_skip_next เสมอ ไม่ว่าผลตรวจใบหน้าจะผ่านหรือไม่) → กดถัดไป

    *** ขอบเขตใหม่ — หยุดหลังกดถัดไปนี้ ไม่ส่งคำขอ/ไม่ยืนยันตัวตนจริง/ไม่ชำระเงิน ***
    """
    res = {"ok": False, "note": ""}
    _bt30_dismiss_news(page, log=log)
    # เปิดโมดอล 'อัปโหลดภาพ'
    opened = page.evaluate(
        r"""() => {
        const vis = el => { const r = el.getBoundingClientRect(); const s = getComputedStyle(el);
            return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden'; };
        const b = Array.from(document.querySelectorAll('button,a'))
            .find(e => vis(e) && /OncOpenModalAddFileIden/.test(e.getAttribute('onclick') || ''));
        if (b) { b.click(); return 'btn'; }
        if (typeof window.OncOpenModalAddFileIden === 'function') { window.OncOpenModalAddFileIden(); return 'fn'; }
        return '';
    }"""
    )
    if not opened:
        res["note"] = "ไม่พบปุ่ม 'อัปโหลดภาพ' (OncOpenModalAddFileIden)"
        return res
    # รอโมดอลแสดง
    shown = False
    for _ in range(15):
        if page.evaluate(
            "() => { const m = document.getElementById('AddFileIden'); return !!(m && m.classList.contains('show')); }"
        ):
            shown = True
            break
        page.wait_for_timeout(300)
    if not shown:
        res["note"] = "เปิดโมดอลอัปโหลดภาพไม่สำเร็จ"
        return res
    page.wait_for_timeout(600)
    # กด 'บันทึก' — เรียก SaveAddFileIden() ตรงๆ (ปุ่ม #buttonSaveFileSelect มีหลายตัวบนหน้า เหมือน 4.2)
    saved = page.evaluate(
        r"""() => {
        try {
            if (typeof window.SaveAddFileIden === 'function') { window.SaveAddFileIden(); return 'called'; }
        } catch (e) { return 'ERR:' + e; }
        const b = document.querySelector('#AddFileIden #buttonSaveFileSelect');
        if (b) { b.click(); return 'clicked'; }
        return 'none';
    }"""
    )
    log(f"        · กดบันทึก (ยืนยันตัวตน ไม่แนบไฟล์จริง) → {saved}")
    # identify() ยิง AJAX แล้ว enable #button_skip_next — รอจน enable (~12s)
    enabled = False
    for _ in range(30):
        st = page.evaluate(
            "() => { const b = document.getElementById('button_skip_next'); "
            "return b ? (b.disabled ? 'disabled' : 'enabled') : 'none'; }"
        )
        if st == "enabled":
            enabled = True
            break
        if st == "none":
            break
        page.wait_for_timeout(400)
    if not enabled:
        # สำรอง: ปุ่มยังไม่ enable (เช่น AJAX ตรวจใบหน้าไม่ตอบ) — ใช้เส้นทางข้ามการยืนยัน
        # identify() เองก็ตั้ง disabled=false ปุ่มนี้ทุกกรณีอยู่แล้ว จึงบังคับ enable ได้สอดคล้องกับตรรกะหน้าเว็บ
        page.evaluate(
            r"""() => {
            try { if (typeof window.SaveAddFileIden_none_verify === 'function') window.SaveAddFileIden_none_verify(); } catch (e) {}
            const b = document.getElementById('button_skip_next'); if (b) b.disabled = false;
        }"""
        )
        page.wait_for_timeout(800)
        log("        · ปุ่มถัดไปยังไม่พร้อม — ใช้เส้นทางข้ามการยืนยัน (none_verify) แล้วบังคับ enable")
    # ปิดโมดอลถ้ายังค้าง
    page.evaluate("() => { try { jQuery('#AddFileIden').modal('hide'); } catch (e) {} }")
    page.wait_for_timeout(500)
    # กดถัดไป (#button_skip_next) — ขอบเขตสุดท้าย (หยุดที่นี่)
    nx = page.evaluate(
        r"""() => {
        const b = document.getElementById('button_skip_next');
        if (b) { b.disabled = false; b.click(); return 'id'; }
        return '';
    }"""
    )
    if not nx:
        res["note"] = "กดถัดไป (ยืนยันตัวตน #button_skip_next) ไม่ได้"
        return res
    page.wait_for_timeout(3000)
    alert = _capture_register_alert(page)
    res["ok"] = True
    res["note"] = (
        alert[:150] if alert else f"อัปโหลดภาพ+บันทึก(ไม่แนบไฟล์จริง)+ถัดไป แล้ว [save={saved}, enabled={enabled}]"
    )
    return res


def _bt30_step7_submit(page: Page, log=print) -> dict[str, Any]:
    """Step 7 (ชำระเงิน): เลือกช่องทาง e-Payment (ถ้ายังไม่เลือก) แล้วกด 'ถัดไป' = *ส่งคำขอจริง*

    *** สำคัญมาก — ขั้นตอนนี้ย้อนกลับไม่ได้ (เป็นการยื่นคำขอจริง) เรียกเฉพาะเมื่อ do_submit=True ***
    มี guard นิรภัย: ต้องตรวจพบว่าเป็นหน้า 'ชำระเงิน' จริง (พบ 'วิธีการชำระเงิน'/'e-Payment'/
    'รายการชำระเงิน'/'ค่ายื่นคำขอ') ก่อนจึงจะกดถัดไป ไม่งั้นยกเลิกเพื่อกันส่งผิดหน้า
    """
    res = {"ok": False, "note": ""}
    _bt30_dismiss_news(page, log=log)
    # ── SAFETY GUARD: ยืนยันว่าอยู่หน้า 'ชำระเงิน' จริงก่อนกดส่ง ──
    on_payment = page.evaluate(
        r"""() => /วิธีการชำระเงิน|e-?Payment|รายการชำระเงิน|ค่ายื่นคำขอ/i.test(document.body.innerText || '')"""
    )
    if not on_payment:
        res["note"] = "ไม่ใช่หน้าชำระเงิน — ยกเลิกการส่งคำขอเพื่อความปลอดภัย"
        log("      ⛔ Step 7: ไม่พบหน้าชำระเงิน — ไม่กดส่งคำขอ (กันส่งผิดหน้า)")
        return res
    # เลือกช่องทาง e-Payment ถ้ายังไม่ถูกเลือก
    page.evaluate(
        r"""() => {
        const cks = Array.from(document.querySelectorAll('input[type=checkbox],input[type=radio]'));
        for (const c of cks) {
            const lbl = ((c.closest('label') && c.closest('label').innerText) || (c.parentElement && c.parentElement.innerText) || '');
            if (/e-?Payment|ชำระเงินผ่าน/i.test(lbl) && !c.checked) {
                c.click();
                if (!c.checked) { c.checked = true; c.dispatchEvent(new Event('change', {bubbles:true})); }
            }
        }
    }"""
    )
    page.wait_for_timeout(500)
    # กด 'ถัดไป' (ปุ่มส่งคำขอ) — เลือกปุ่มที่มองเห็น+ใช้งานได้ตัวสุดท้าย
    clicked = page.evaluate(
        r"""() => {
        const vis = el => { const r = el.getBoundingClientRect(); const s = getComputedStyle(el);
            return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden'; };
        const cand = Array.from(document.querySelectorAll('button,a'))
            .filter(e => vis(e) && !e.disabled && /ถัดไป/.test((e.textContent || '').trim())
                && (e.textContent || '').trim().length < 20
                && !/ย้อนกลับ|ยกเลิก/.test((e.textContent || '').trim()));
        if (cand.length) { cand[cand.length - 1].click(); return true; }
        return false;
    }"""
    )
    if not clicked:
        res["note"] = "ไม่พบปุ่มถัดไป (หน้าชำระเงิน)"
        return res
    log("      · Step 7: กดถัดไป (ส่งคำขอ) แล้ว — รอหน้าผลสำเร็จ...")
    # รอหน้าผลสำเร็จ (เผื่อมีโมดอลยืนยันก็กดยืนยันให้)
    ok_success = False
    for _ in range(25):
        if page.evaluate(
            r"""() => /เลขที่คำขอ|ส่งใบคำขอ.*เรียบร้อย|เรียบร้อยแล้ว|E-?Tracking|พิมพ์แบบฟอร์มการชำระเงิน/i.test(document.body.innerText || '')"""
        ):
            ok_success = True
            break
        page.evaluate(
            r"""() => {
            const vis = el => { const r = el.getBoundingClientRect(); const s = getComputedStyle(el);
                return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden'; };
            const b = Array.from(document.querySelectorAll('.swal2-confirm,button,a'))
                .find(e => vis(e) && /^(ยืนยัน|ตกลง|confirm|ใช่|ส่งคำขอ)/i.test((e.textContent || '').trim())
                    && (e.textContent || '').trim().length < 20);
            if (b) b.click();
        }"""
        )
        page.wait_for_timeout(1000)
    _bt30_dismiss_news(page, log=log)
    res["ok"] = ok_success
    res["note"] = ("ส่งคำขอแล้ว — พบหน้าผลสำเร็จ" if ok_success
                   else "กดถัดไปแล้ว แต่ยังไม่พบหน้าผลสำเร็จ (โปรดตรวจสอบด้วยตนเอง)")
    return res


def _bt30_parse_payment_pdf(pdf_bytes: bytes) -> dict[str, str]:
    """แยกข้อมูลสำคัญจากไฟล์ PDF 'ใบแจ้งชำระเงิน' (Bill Payment)

    หน้าผลสำเร็จของระบบโหลด 'ค่า' (เลขที่คำขอ/วันที่/ยอดเงิน) แบบ async ทีหลัง ทำให้ scrape DOM
    ไม่เสถียร แต่ไฟล์ PDF ใบแจ้งชำระเงินมีโครงสร้างคงที่และเชื่อถือได้ จึงใช้เป็นแหล่งข้อมูลหลัก
    คืน dict: request_no, bill_no, amount, due, txn_date, ref1, ref2, alien_ref (หาไม่พบ = '')
    """
    out = {"request_no": "", "bill_no": "", "amount": "", "due": "",
           "txn_date": "", "ref1": "", "ref2": "", "alien_ref": ""}
    if not pdf_bytes:
        return out
    try:
        import io
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(pdf_bytes))
        txt = "\n".join((pg.extract_text() or "") for pg in reader.pages)
    except Exception:
        return out
    import re as _re

    def g(pat: str) -> str:
        m = _re.search(pat, txt)
        return m.group(1).strip() if m else ""

    out["request_no"] = g(r"\(Request No\.\)\s*([0-9]{8,})")
    out["bill_no"] = g(r"\(Bill Payment No\.\)\s*([0-9]{8,})")
    out["due"] = g(r"\(Payment Due Date\)\s*([0-9]{2}-[0-9]{2}-[0-9]{4}\s+[0-9]{1,2}:[0-9]{2})")
    out["txn_date"] = g(r"\(Transaction Date\)\s*([0-9]{2}-[0-9]{2}-[0-9]{4}\s+[0-9]{1,2}:[0-9]{2})")
    mref = _re.search(r"Ref\. No\. 1\s*Ref\. No\. 2\s*([0-9]{6,})\s+([0-9]{6,})", txt)
    if mref:
        out["ref1"], out["ref2"] = mref.group(1), mref.group(2)
    out["amount"] = (g(r"Total Amount\s*([0-9,]+\.[0-9]{2})")
                     or g(r"\(Amount\)\s*([0-9,]+\.[0-9]{2})"))
    out["alien_ref"] = g(r"\b([A-Z]{2}[0-9]{12,})\b")
    return out


def _bt30_step8_capture_print(page: Page, rec: dict[str, Any], screenshot_dir: Path, log=print) -> dict[str, Any]:
    """Step 8 (หน้าผลสำเร็จ): 8.1 เก็บรายละเอียดทั้งหมด + 8.2 ดาวน์โหลด 'ใบแจ้งชำระเงิน'

    8.2 ดาวน์โหลดได้ครั้งเดียว → ลองซ้ำสูงสุด 3 ครั้งถ้ายังไม่ได้ไฟล์ และเตือนชัดเจนถ้าล้มเหลว
    บันทึกไฟล์ลง reports/bt30_submitted/  (PDF + JSON ข้อมูลหน้า)
    """
    res: dict[str, Any] = {"ok": False, "note": "", "request_no": "", "pdf_file": "", "data_file": ""}
    submitted_dir = screenshot_dir.parent / "bt30_submitted"
    submitted_dir.mkdir(parents=True, exist_ok=True)
    seq = rec.get("seq", "?")
    name = rec.get("name", "")

    # ── รอหน้าผลสำเร็จโหลดค่าจริง (กันจับตอนยังเป็น spinner/โหลด async ไม่เสร็จ) ──
    for _ in range(20):
        loaded = page.evaluate(
            r"""() => {
            const t = document.body.innerText || '';
            // ปรากฏเลขชุดยาว (เลขที่คำขอ/อ้างอิง) หรือ ปุ่มพิมพ์ใบชำระเงิน = โหลดเสร็จแล้ว
            return /\d{10,}/.test(t) || /พิมพ์.{0,8}ชำระเงิน/.test(t);
        }"""
        )
        if loaded:
            break
        page.wait_for_timeout(1000)
    page.wait_for_timeout(800)

    # ── 8.1 เก็บข้อมูลหน้าผลสำเร็จ (best-effort; แหล่งข้อมูลหลักคือ PDF ใน 8.3) ──
    data = page.evaluate(
        r"""() => {
        const lines = (document.body.innerText || '').split('\n').map(s => s.trim()).filter(s => s.length);
        const after = (re) => {
            for (let i = 0; i < lines.length; i++) {
                if (re.test(lines[i])) {
                    const same = lines[i].replace(re, '').trim();
                    if (same) return same;
                    if (i + 1 < lines.length) return lines[i + 1];
                }
            }
            return '';
        };
        const out = {};
        out.request_no = (after(/^เลขที่คำขอ/) || '').replace(/[^0-9]/g, '');
        out.subject = after(/^ระบบได้รับคำขอเรื่อง/);
        out.submit_date = after(/^วันที่ยื่นคำขอ/);
        out.alien = after(/^คนต่างด้าว/);
        out.pay_method = after(/^วิธีการชำระเงิน/);
        out.ref1 = (after(/^หมายเลขอ้างอิง\s*1/) || '').replace(/[^0-9]/g, '');
        out.ref2 = (after(/^หมายเลขอ้างอิง\s*2/) || '').replace(/[^0-9]/g, '');
        out.amount = after(/^ยอดชำระ/);
        const m = (document.body.innerText || '').match(/ภายในวันที่\s*([^\n]+)/);
        out.due = m ? m[1].trim() : '';
        out.full_text = (document.body.innerText || '');
        return out;
    }"""
    )
    dom_request_no = (data.get("request_no") or "").strip()

    base = _safe_filename(f"{seq}_{name}_step8_success")
    try:
        page.screenshot(path=str(screenshot_dir / f"{base}.png"), full_page=True)
    except Exception:
        pass

    # ── 8.2 ดาวน์โหลด 'พิมพ์แบบฟอร์มการชำระเงิน' (ดาวน์โหลดได้ครั้งเดียว → ลองซ้ำถ้าพลาด) ──
    def _do_click():
        page.evaluate(
            r"""() => {
            const vis = el => { const r = el.getBoundingClientRect(); const s = getComputedStyle(el);
                return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden'; };
            const b = Array.from(document.querySelectorAll('button,a'))
                .find(e => vis(e) && /พิมพ์.{0,8}(ฟอร์ม).{0,8}ชำระเงิน|พิมพ์.{0,8}ชำระเงิน/.test((e.textContent || '').trim()));
            if (b) b.click();
        }"""
        )

    body: bytes = b""
    err_last = ""
    for attempt in range(1, 4):
        b, err = _grab_pdf_after_click(page, _do_click, log=lambda *a: None)
        if b and len(b) >= 500 and b[:5] == b"%PDF-":
            body = b
            break
        err_last = err or "ไม่ได้ไฟล์"
        log(f"        · ⚠ 8.2 ดาวน์โหลดใบชำระเงินครั้งที่ {attempt} ไม่สำเร็จ ({err_last}) — ลองใหม่")
        page.wait_for_timeout(2500)

    # ── 8.3 แยกข้อมูลจากไฟล์ PDF (แหล่งข้อมูลหลัก เชื่อถือได้กว่า DOM ที่โหลดแบบ async) ──
    pdf_info = _bt30_parse_payment_pdf(body) if body else {}
    request_no = dom_request_no or pdf_info.get("request_no", "")
    if pdf_info:
        data["request_no"] = request_no
        for k in ("bill_no", "amount", "due", "txn_date", "ref1", "ref2", "alien_ref"):
            v = pdf_info.get(k, "")
            if v and not data.get(k):
                data[k] = v
    res["request_no"] = request_no
    res["amount"] = pdf_info.get("amount", "") or data.get("amount", "")
    res["due"] = pdf_info.get("due", "") or data.get("due", "")
    res["data"] = data
    log(f"        · 8.1 เลขที่คำขอ={request_no or '-'} | ยอด={res['amount'] or '-'} | "
        f"ชำระภายใน={res['due'] or '-'} | อ้างอิง1={data.get('ref1','-')} อ้างอิง2={data.get('ref2','-')}"
        + ("  [จาก PDF]" if pdf_info.get("request_no") and not dom_request_no else ""))

    # ── บันทึกไฟล์ด้วยชื่อที่อิงเลขที่คำขอจริง ──
    stem = _safe_filename(f"{request_no or seq}_{name}_submitted")
    try:
        (submitted_dir / f"{stem}.json").write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        res["data_file"] = f"{stem}.json"
    except Exception as e:
        log(f"        · ⚠ บันทึก JSON ไม่สำเร็จ: {e}")

    if body:
        # ชื่อไฟล์: {เลขที่คำขอ}_{เลขที่เอกสาร}_{ชื่อ}_payment.pdf
        # ถ้าไม่มี 'เลขที่เอกสาร' (2.3 doc_no) → ใช้ 'เลขที่ (*เอกสารแสดงการอนุญาตหรือการรับรอง*)' (auth_no) แทน
        doc_id = (rec.get("doc_no") or rec.get("auth_no") or "").strip()
        name_parts: list[str] = []
        for p in (request_no or seq, doc_id, name):
            s = _safe_filename(p)
            if p and s and s != "x":
                name_parts.append(s)
        pdf_name = "_".join(name_parts) + "_payment.pdf"
        try:
            (submitted_dir / pdf_name).write_bytes(body)
            res["pdf_file"] = pdf_name
            log(f"        · ✓ 8.2 ดาวน์โหลดใบชำระเงินแล้ว → bt30_submitted/{pdf_name} ({len(body)//1024} KB)")
        except Exception as e:
            log(f"        · ⚠ เขียนไฟล์ PDF ไม่สำเร็จ: {e}")
    else:
        log(f"        · ⛔ 8.2 ดาวน์โหลดใบชำระเงินไม่สำเร็จหลังลอง 3 ครั้ง ({err_last}) — "
            f"กรุณาดาวน์โหลดเองทันที (เลขคำขอ {request_no or '-'}) เพราะดาวน์โหลดได้ครั้งเดียว")

    res["ok"] = bool(request_no) and bool(res["pdf_file"])
    res["note"] = (f"เลขคำขอ={request_no or '-'} | ยอด={res['amount'] or '-'} | "
                   f"ชำระภายใน={res['due'] or '-'} | ใบชำระเงิน="
                   + ("OK" if res["pdf_file"] else f"ล้มเหลว({err_last})"))
    return res


def _bt30_do_step2(
    page: Page,
    rec: dict[str, Any],
    screenshot_dir: Path,
    log=print,
    do_submit: bool = False,
) -> dict[str, Any]:
    """รวมขั้นตอนที่ 2 (2.1–2.5) + หน้า 2/2 + แนบเอกสาร + Step 4–6.1 สำหรับคนต่างด้าว 1 คน
    หลังเพิ่มข้อมูลในขั้นตอนที่ 1 สำเร็จ

    do_submit=False (ค่าเริ่มต้น/ปลอดภัย): หยุดหลัง Step 6.1 — ไม่ส่งคำขอ/ไม่ชำระเงิน
    do_submit=True  (*** ส่งคำขอจริง ***): ทำ Step 7 (ชำระเงิน+ส่งคำขอ) + Step 8
      (เก็บข้อมูลหน้าผลสำเร็จ + ดาวน์โหลดใบแจ้งชำระเงิน) — ย้อนกลับไม่ได้
    คืน {step2_status, step2_note, step2_screenshot, request_no, submit_pdf, submit_status}
    """
    seq = rec.get("seq", "?")
    name = rec.get("name", "")
    base = _safe_filename(f"{seq}_{name}_step2")
    parts: list[str] = []
    out = {"step2_status": "", "step2_note": "", "step2_screenshot": "",
           "request_no": "", "submit_pdf": "", "submit_status": ""}
    try:
        # 2.1 รับรอง + ถัดไป
        if not _bt30_consent_next(page, log=log):
            out["step2_status"] = "FAIL"
            out["step2_note"] = "2.1 รับรอง/ถัดไป ไม่สำเร็จ"
            try:
                page.screenshot(path=str(screenshot_dir / f"{base}_consentfail.png"), full_page=True)
                out["step2_screenshot"] = f"{base}_consentfail.png"
            except Exception:
                pass
            return out
        parts.append("2.1 รับรอง+ถัดไป: OK")
        log("      ✓ 2.1 รับรอง + ถัดไป → หน้ารายละเอียด")

        # 2.2 ที่อยู่ที่ติดต่อได้
        r_addr = _bt30_fill_address(page, rec, log=log)
        parts.append(f"2.2 ที่อยู่: {'OK' if r_addr['ok'] else 'X'} {r_addr['note']}".strip())
        log(f"      {'✓' if r_addr['ok'] else '⚠'} 2.2 ที่อยู่ที่ติดต่อได้ — {r_addr['note']}")

        # 2.3+2.4 ข้อมูลเพิ่มเติม (เอกสาร/วีซ่า/ตม.)
        r_stay = _bt30_fill_staypermit(page, rec, log=log)
        parts.append(f"2.3+2.4 ข้อมูลเพิ่มเติม: {'OK' if r_stay['ok'] else 'X'} {r_stay['note']}".strip())
        log(f"      {'✓' if r_stay['ok'] else '⚠'} 2.3+2.4 ข้อมูลเพิ่มเติม — {r_stay['note']}")

        # screenshot ก่อนกดถัดไป
        try:
            page.screenshot(path=str(screenshot_dir / f"{base}_filled.png"), full_page=True)
            out["step2_screenshot"] = f"{base}_filled.png"
        except Exception:
            pass

        # 2.5 ถัดไป
        r_next = _bt30_step2_next(page, log=log)
        parts.append(f"2.5 ถัดไป: {'OK' if r_next['ok'] else 'X'} {r_next['note']}".strip())
        log(f"      {'✓' if r_next['ok'] else '⚠'} 2.5 ถัดไป — {r_next['note']}")
        try:
            page.screenshot(path=str(screenshot_dir / f"{base}_afternext.png"), full_page=True)
            out["step2_screenshot"] = f"{base}_afternext.png"
        except Exception:
            pass

        # 2.6 หน้า 2/2: โดยจะมาทำงาน (สถานที่ทำงาน/ประเภทกิจการ) + เอกสารแสดงการอนุญาตหรือการรับรอง
        r_page2 = {"ok": False, "note": "ข้าม (2.5 ไม่สำเร็จ)"}
        r_p2next = {"ok": False, "note": "ข้าม"}
        if r_next["ok"]:
            page.wait_for_timeout(1500)
            r_page2 = _bt30_fill_page2(page, rec, log=log)
            parts.append(f"2.6 หน้า2/2: {'OK' if r_page2['ok'] else 'X'} {r_page2['note']}".strip())
            log(f"      {'✓' if r_page2['ok'] else '⚠'} 2.6 หน้า 2/2 (โดยจะมาทำงาน/เอกสารอนุญาต) — {r_page2['note']}")
            try:
                page.screenshot(path=str(screenshot_dir / f"{base}_page2filled.png"), full_page=True)
                out["step2_screenshot"] = f"{base}_page2filled.png"
            except Exception:
                pass

            # 2.7 ถัดไป (หน้า 2/2 → แนบเอกสาร)
            r_p2next = _bt30_page2_next(page, log=log)
            parts.append(f"2.7 ถัดไป(2/2): {'OK' if r_p2next['ok'] else 'X'} {r_p2next['note']}".strip())
            log(f"      {'✓' if r_p2next['ok'] else '⚠'} 2.7 ถัดไป (หน้า 2/2) — {r_p2next['note']}")
            try:
                page.screenshot(path=str(screenshot_dir / f"{base}_page2next.png"), full_page=True)
                out["step2_screenshot"] = f"{base}_page2next.png"
            except Exception:
                pass

        # 2.8 แนบเอกสาร (3.1–4.1 + 4.2 เอกสารอื่นๆ) + 2.9 ถัดไป
        r_attach = {"ok": False, "note": "ข้าม (2.7 ไม่สำเร็จ)"}
        r_attnext = {"ok": False, "note": "ข้าม"}
        r_emp_next = {"ok": False, "note": "ข้าม"}       # Step 4 เอกสารนายจ้าง 2/2
        r_sum1_next = {"ok": False, "note": "ข้าม"}      # Step 5.1 สรุปคำขอ 1/2
        r_sum2_confirm = {"ok": False, "note": "ข้าม"}   # Step 5.2 สรุปคำขอ 2/2 (ยืนยันข้อมูล)
        r_iden = {"ok": False, "note": "ข้าม"}           # Step 6.1 วิธีการยืนยันตัวตน
        r_step7 = {"ok": False, "note": "ข้าม (ไม่ส่งคำขอ)"}   # Step 7 ชำระเงิน+ส่งคำขอ
        r_step8 = {"ok": False, "note": "ข้าม"}           # Step 8 เก็บข้อมูล+ดาวน์โหลดใบชำระเงิน
        if r_p2next["ok"]:
            page.wait_for_timeout(2500)
            r_attach = _bt30_fill_attachments(page, rec, log=log)
            parts.append(f"2.8 แนบเอกสาร: {'OK' if r_attach['ok'] else 'X'} {r_attach['note']}".strip())
            log(f"      {'✓' if r_attach['ok'] else '⚠'} 2.8 แนบเอกสาร — {r_attach['note']}")
            try:
                page.screenshot(path=str(screenshot_dir / f"{base}_attachfilled.png"), full_page=True)
                out["step2_screenshot"] = f"{base}_attachfilled.png"
            except Exception:
                pass

            # 2.9 ถัดไป (แนบเอกสาร) — เฉพาะเมื่อแนบครบ เพื่อกันถูกบล็อกจากเอกสารบังคับที่ขาด
            if r_attach["ok"]:
                r_attnext = _bt30_attach_next(page, log=log)
                parts.append(f"2.9 ถัดไป(แนบ): {'OK' if r_attnext['ok'] else 'X'} {r_attnext['note']}".strip())
                log(f"      {'✓' if r_attnext['ok'] else '⚠'} 2.9 ถัดไป (แนบเอกสาร) — {r_attnext['note']}")
                try:
                    page.screenshot(path=str(screenshot_dir / f"{base}_attachnext.png"), full_page=True)
                    out["step2_screenshot"] = f"{base}_attachnext.png"
                except Exception:
                    pass
            else:
                parts.append("2.9 ถัดไป(แนบ): ข้าม (แนบไม่ครบ)")

        # Step 4–6.1 (หลังแนบเอกสาร 2.9) — เดินต่อจนถึง 'ยืนยันตัวตน' แล้วหยุด (ขอบเขตใหม่)
        # *** ไม่ส่งคำขอ / ไม่ยืนยันตัวตนจริง / ไม่ชำระเงิน ***
        if r_attnext["ok"]:
            # Step 4: เอกสารนายจ้าง (แนบเอกสาร หน้า 2/2) — กดถัดไป
            page.wait_for_timeout(2500)
            r_emp_next = _bt30_click_next_generic(
                page, "NextStepThreePageOneRenew", "4 เอกสารนายจ้าง 2/2", log=log)
            parts.append(f"4 เอกสารนายจ้าง: {'OK' if r_emp_next['ok'] else 'X'} {r_emp_next['note']}".strip())
            log(f"      {'✓' if r_emp_next['ok'] else '⚠'} 4 เอกสารนายจ้าง (หน้า 2/2) — {r_emp_next['note']}")
            try:
                page.screenshot(path=str(screenshot_dir / f"{base}_step4.png"), full_page=True)
                out["step2_screenshot"] = f"{base}_step4.png"
            except Exception:
                pass

            # Step 5.1: สรุปคำขอ หน้า 1/2 — กดถัดไป
            if r_emp_next["ok"]:
                r_sum1_next = _bt30_click_next_generic(page, None, "5.1 สรุปคำขอ 1/2", log=log)
                parts.append(f"5.1 สรุปคำขอ 1/2: {'OK' if r_sum1_next['ok'] else 'X'} {r_sum1_next['note']}".strip())
                log(f"      {'✓' if r_sum1_next['ok'] else '⚠'} 5.1 สรุปคำขอ (หน้า 1/2) — {r_sum1_next['note']}")
                try:
                    page.screenshot(path=str(screenshot_dir / f"{base}_step5_1.png"), full_page=True)
                    out["step2_screenshot"] = f"{base}_step5_1.png"
                except Exception:
                    pass

            # Step 5.2: สรุปคำขอ หน้า 2/2 — ติ๊กยืนยันข้อมูล + กดถัดไป
            if r_sum1_next["ok"]:
                r_sum2_confirm = _bt30_step5_confirm_next(page, log=log)
                parts.append(f"5.2 ยืนยันข้อมูล: {'OK' if r_sum2_confirm['ok'] else 'X'} {r_sum2_confirm['note']}".strip())
                log(f"      {'✓' if r_sum2_confirm['ok'] else '⚠'} 5.2 สรุปคำขอ (หน้า 2/2) ยืนยันข้อมูล — {r_sum2_confirm['note']}")
                try:
                    page.screenshot(path=str(screenshot_dir / f"{base}_step5_2.png"), full_page=True)
                    out["step2_screenshot"] = f"{base}_step5_2.png"
                except Exception:
                    pass

            # Step 6.1: วิธีการยืนยันตัวตน — อัปโหลดภาพ + บันทึก (ไม่แนบไฟล์จริง) + ถัดไป (ขอบเขตสุดท้าย)
            if r_sum2_confirm["ok"]:
                page.wait_for_timeout(1500)
                r_iden = _bt30_step6_upload_identity(page, log=log)
                parts.append(f"6.1 ยืนยันตัวตน: {'OK' if r_iden['ok'] else 'X'} {r_iden['note']}".strip())
                log(f"      {'✓' if r_iden['ok'] else '⚠'} 6.1 วิธีการยืนยันตัวตน — {r_iden['note']}")
                try:
                    page.screenshot(path=str(screenshot_dir / f"{base}_step6.png"), full_page=True)
                    out["step2_screenshot"] = f"{base}_step6.png"
                except Exception:
                    pass

            # Step 7–8 (ชำระเงิน+ส่งคำขอจริง + เก็บข้อมูล+ดาวน์โหลดใบชำระเงิน)
            # *** IRREVERSIBLE — ทำเฉพาะเมื่อ do_submit=True เท่านั้น (ค่าเริ่มต้น False = หยุดที่ 6.1) ***
            if do_submit and r_iden["ok"]:
                page.wait_for_timeout(1500)
                r_step7 = _bt30_step7_submit(page, log=log)
                parts.append(f"7 ส่งคำขอ: {'OK' if r_step7['ok'] else 'X'} {r_step7['note']}".strip())
                log(f"      {'✓' if r_step7['ok'] else '⚠'} 7 ชำระเงิน/ส่งคำขอ — {r_step7['note']}")
                try:
                    page.screenshot(path=str(screenshot_dir / f"{base}_step7.png"), full_page=True)
                    out["step2_screenshot"] = f"{base}_step7.png"
                except Exception:
                    pass
                if r_step7["ok"]:
                    r_step8 = _bt30_step8_capture_print(page, rec, screenshot_dir, log=log)
                    parts.append(f"8 ใบชำระเงิน: {'OK' if r_step8['ok'] else 'X'} {r_step8['note']}".strip())
                    log(f"      {'✓' if r_step8['ok'] else '⚠'} 8 เก็บข้อมูล+ดาวน์โหลดใบชำระเงิน — {r_step8['note']}")
                    out["request_no"] = r_step8.get("request_no", "")
                    out["submit_pdf"] = r_step8.get("pdf_file", "")
                    out["submit_status"] = "SUBMITTED" if r_step8.get("request_no") else "UNKNOWN"

        ok_all = (r_addr["ok"] and r_stay["ok"] and r_next["ok"]
                  and r_page2["ok"] and r_p2next["ok"]
                  and r_attach["ok"] and r_attnext["ok"]
                  and r_emp_next["ok"] and r_sum1_next["ok"]
                  and r_sum2_confirm["ok"] and r_iden["ok"])
        if do_submit:
            ok_all = ok_all and r_step7["ok"] and r_step8["ok"]
        out["step2_status"] = "SUCCESS" if ok_all else "PARTIAL"
        out["step2_note"] = " | ".join(parts)[:900]
        return out
    except Exception as e:
        try:
            page.screenshot(path=str(screenshot_dir / f"{base}_error.png"), full_page=True)
            out["step2_screenshot"] = f"{base}_error.png"
        except Exception:
            pass
        out["step2_status"] = "ERROR"
        out["step2_note"] = (" | ".join(parts) + " | " + str(e))[:500]
        return out


def _save_bt30_report(rows: list[dict[str, Any]], out_path: Path, log=print) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "บต.30"
    headers = ["ลำดับ", "คำนำหน้า", "ชื่อ", "สัญชาติ", "เพศ", "วันเกิด",
               "สถานะ", "หมายเหตุ", "Screenshot",
               "สถานะขั้นตอน2", "หมายเหตุขั้นตอน2", "Screenshot2",
               "สถานะส่งคำขอ", "เลขที่คำขอ", "ไฟล์ใบชำระเงิน"]
    ws.append(headers)
    head_fill = PatternFill("solid", fgColor="305496")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r in rows:
        ws.append([
            r.get("seq", ""), r.get("prefix", ""), r.get("name", ""),
            r.get("nationality", ""), r.get("sex", ""), r.get("birthdate", ""),
            r.get("status", ""), r.get("note", ""), r.get("screenshot", ""),
            r.get("step2_status", ""), r.get("step2_note", ""), r.get("step2_screenshot", ""),
            r.get("submit_status", ""), r.get("request_no", ""), r.get("submit_pdf", ""),
        ])
    widths = [8, 12, 28, 16, 8, 14, 12, 44, 28, 14, 50, 30, 14, 18, 32]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    log(f"บันทึกรายงาน: {out_path}")


# คอลัมน์ผลลัพธ์ที่ระบบเขียนกลับเข้าไฟล์ Excel ต้นทาง (ต่อท้ายขวาสุดถ้ายังไม่มี)
_BT30_WB_COLS = [
    "สถานะระบบ", "สถานะส่งคำขอ", "เลขที่คำขอ",
    "ไฟล์ใบชำระเงิน", "วันที่ทำรายการ", "หมายเหตุระบบ",
]


def _bt30_writeback_excel(excel_path: Path, row: dict[str, Any], log=print) -> bool:
    """เขียนผลลัพธ์รายแถวกลับเข้าไฟล์ Excel ต้นทาง (เพิ่มคอลัมน์ผลที่ขวาสุดถ้ายังไม่มี)
    จับคู่แถวด้วย row['row_index'] (แถว Excel = row_index + 1 เพราะข้อมูลเริ่มแถว 2)
    คืน True ถ้าบันทึกสำเร็จ — ใช้สำหรับติดตามสถานะ + resume/skip คนที่ทำเสร็จแล้ว
    """
    try:
        excel_path = Path(excel_path)
        wb = load_workbook(excel_path)
        ws = wb.active
        hdr = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
        col_idx: dict[str, int] = {}
        for name in _BT30_WB_COLS:
            if name in hdr:
                col_idx[name] = hdr.index(name) + 1
            else:
                new_c = len(hdr) + 1
                ws.cell(row=1, column=new_c, value=name)
                hdr.append(name)
                col_idx[name] = new_c
        try:
            ridx = int(row.get("row_index", 0))
        except (TypeError, ValueError):
            ridx = 0
        if ridx <= 0:
            return False
        excel_row = ridx + 1  # ข้อมูลเริ่มแถว 2
        vals = {
            "สถานะระบบ": row.get("step2_status") or row.get("status") or "",
            "สถานะส่งคำขอ": row.get("submit_status", ""),
            "เลขที่คำขอ": row.get("request_no", ""),
            "ไฟล์ใบชำระเงิน": row.get("submit_pdf", ""),
            "วันที่ทำรายการ": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "หมายเหตุระบบ": (row.get("step2_note") or row.get("note") or "")[:250],
        }
        for name, v in vals.items():
            ws.cell(row=excel_row, column=col_idx[name], value=v)
        wb.save(excel_path)
        return True
    except PermissionError:
        log(f"      ⚠ เขียนผลกลับ Excel ไม่ได้ (ไฟล์ {Path(excel_path).name} อาจเปิดค้างอยู่) — กรุณาปิดไฟล์แล้วลองใหม่")
        return False
    except Exception as e:
        log(f"      ⚠ เขียนผลกลับ Excel ไม่สำเร็จ: {e}")
        return False


def _bt30_resolve_groups(
    records: list[dict[str, Any]],
    accounts: dict[str, dict[str, str]],
    fallback_cfg: dict[str, str],
    log=print,
) -> list[tuple[dict[str, str], list[dict[str, Any]]]]:
    """จัดกลุ่ม records ตามคอลัมน์ Username (รักษาลำดับที่ปรากฏในไฟล์) เพื่อ login ทีละบัญชี
    คืนค่า: [(login_cfg, [records...]), ...]
      - แถวที่เว้น Username ว่าง → ใช้ fallback_cfg (บัญชีหลัก/ค่าจากหน้าโปรแกรม)
      - Username ที่ไม่พบใน UsernameLogin.xlsx → เตือนแล้ว fallback ไปบัญชีหลัก
      - แถว Username เดียวกันถูกรวมไว้กลุ่มเดียว (login ครั้งเดียวต่อบัญชี)
    """
    order: list[str] = []
    buckets: dict[str, list[dict[str, Any]]] = {}
    cfgs: dict[str, dict[str, str]] = {}
    for rec in records:
        uname = (rec.get("username") or "").strip()
        key = uname.lower() or "\x00default"
        if key not in buckets:
            buckets[key] = []
            order.append(key)
            if not uname:
                cfgs[key] = fallback_cfg
            else:
                acct = accounts.get(key)
                if acct:
                    cfgs[key] = {
                        "username": acct["username"],
                        "password": acct["password"],
                        "user_type": acct["type"],
                        "method": acct.get("method") or fallback_cfg.get("method", "E-Workpermit"),
                    }
                else:
                    log(f"  ⚠ ไม่พบบัญชี '{uname}' ใน UsernameLogin.xlsx — "
                        f"ใช้บัญชีหลัก ({fallback_cfg.get('username','')}) แทน")
                    cfgs[key] = fallback_cfg
        buckets[key].append(rec)
    return [(cfgs[k], buckets[k]) for k in order]


def run_bt30(
    cfg: dict,
    excel_input: Path,
    login_excel: Path,
    out_path: Path,
    row_range: str | None = None,
    do_step2: bool = True,
    do_submit: bool = False,
    skip_done: bool = True,
    log=print,
    progress=None,
    is_cancelled=None,
) -> tuple[int, Path]:
    """โหมด บต.30 — ยื่นต่ออายุใบอนุญาตทำงานตาม MoU
    - Login จาก UsernameLogin.xlsx: ถ้า from_bt30.xlsx มีคอลัมน์ 'Username' จะจัดกลุ่มตาม
      บัญชีแล้ว login แยกทีละกลุ่ม (รองรับยื่นหลาย Username ในไฟล์เดียว) — แถวที่เว้นว่าง/
      ไม่พบบัญชี ใช้บัญชีหลัก (บัญชีแรกใน UsernameLogin.xlsx) หรือค่าใน cfg
    - do_step2=False : ขั้นตอนที่ 1 เท่านั้น — เปิดฟอร์มครั้งเดียว วนเพิ่มคนต่างด้าวทุกแถว แล้วกดบันทึก
    - do_step2=True  : ทำครบขั้นตอนที่ 1+2 แบบรายคน (1 แถว = 1 คำขอ): เปิดฟอร์มใหม่ →
      ค้นหา/บันทึกคนต่างด้าว → 2.1 รับรอง+ถัดไป → 2.2 ที่อยู่ → 2.3+2.4 ข้อมูลเพิ่มเติม →
      2.5 ถัดไป → 2.6 หน้า 2/2 → 2.7 ถัดไป → 2.8 แนบเอกสาร (13 ไฟล์) → 2.9 ถัดไป →
      Step 4 เอกสารนายจ้าง (ถัดไป) → Step 5.1 สรุปคำขอ 1/2 (ถัดไป) →
      Step 5.2 สรุปคำขอ 2/2 (ติ๊กยืนยันข้อมูล+ถัดไป) → Step 6.1 ยืนยันตัวตน
      (อัปโหลดภาพ+บันทึก ไม่แนบไฟล์จริง+ถัดไป)
    - do_submit=False (ค่าเริ่มต้น/ปลอดภัย): หยุดหลัง Step 6.1 — ไม่ส่งคำขอ/ไม่ชำระเงิน
    - do_submit=True  (*** ส่งคำขอจริง — ย้อนกลับไม่ได้ ***): ทำ Step 7 (ชำระเงิน e-Payment + ถัดไป = ส่งคำขอ)
      → Step 8 (เก็บข้อมูลหน้าผลสำเร็จทั้งหมด + ดาวน์โหลดใบแจ้งชำระเงิน) → ไฟล์ลง reports/bt30_submitted/
    - เขียนผลกลับเข้าไฟล์ Excel ต้นทางรายแถว (คอลัมน์ขวาสุด: สถานะระบบ/สถานะส่งคำขอ/เลขที่คำขอ/
      ไฟล์ใบชำระเงิน/วันที่ทำรายการ/หมายเหตุระบบ) ทันทีที่ทำแต่ละแถวเสร็จ
    - skip_done=True (ค่าเริ่มต้น): ข้ามแถวที่มี 'เลขที่คำขอ' อยู่แล้ว (ส่งคำขอไปแล้ว) อัตโนมัติ —
      กันส่งซ้ำ/ทำซ้ำ และรองรับการรันต่อ (resume) หลังหยุดกลางคัน
    """
    out_path = _timestamped_path(out_path)
    records = _read_bt30_excel(excel_input)
    accounts = (_read_login_accounts(login_excel)
                if login_excel and Path(login_excel).exists() else {})
    total = len(records)
    indices = _parse_row_range(row_range, total)
    selected = [records[i - 1] for i in indices]
    mode_txt = "ขั้นตอน 1+2 (รายคน)" if do_step2 else "ขั้นตอน 1 เท่านั้น"
    log(f"[1/3] อ่าน {Path(excel_input).name}: {total} แถว → จะทำ {len(selected)} แถว "
        f"({row_range or 'ทั้งหมด'}) | โหมด: {mode_txt}")
    log(f"      ไฟล์รายงาน: {out_path.name}")

    # ตรวจขนาดไฟล์แนบทุกแถวก่อนเริ่ม (ระบบจำกัด 4 MB/ไฟล์) — เฉพาะโหมดที่มีการแนบเอกสาร
    if do_step2:
        _bt30_preflight_doc_sizes(selected, log=log)

    if accounts:
        acct = next(iter(accounts.values()))
        login_cfg = {
            "username": acct["username"], "password": acct["password"],
            "user_type": acct["type"],
            "method": acct.get("method") or cfg.get("method", "E-Workpermit"),
        }
    else:
        login_cfg = {k: cfg.get(k, "") for k in ("username", "password", "user_type", "method")}
    if not login_cfg.get("username") or not login_cfg.get("password"):
        raise ValueError("ไม่พบบัญชี login — กรุณาระบุ UsernameLogin.xlsx หรือกรอก Username/Password")
    log(f"      บัญชีหลัก (fallback): {login_cfg['username']} ({login_cfg.get('user_type','')})")

    screenshot_dir = out_path.parent / "bt30_screenshots"
    screenshot_dir.mkdir(parents=True, exist_ok=True)

    results: list[dict[str, Any]] = []
    success = 0
    if progress:
        try: progress(0, len(selected))
        except Exception: pass

    with sync_playwright() as pw:
        browser = _launch_chromium(pw, cfg, ["--ignore-certificate-errors", "--start-maximized"])
        ctx = browser.new_context(
            locale="th-TH", ignore_https_errors=True,
            viewport={"width": 1920, "height": 1080},
        )
        page = ctx.new_page()
        try:
            if do_step2:
                # --- รายคน: จัดกลุ่มตามคอลัมน์ Username → login แยกทีละบัญชี →
                #     เปิดฟอร์มใหม่ทุกแถว แล้วทำขั้นตอน 1 + 2 ให้ครบ ---
                groups = _bt30_resolve_groups(selected, accounts, login_cfg, log=log)
                multi = len(groups) > 1
                if multi:
                    log(f"[2/3] พบ {len(groups)} บัญชี (Username) ในรายการ — จะ login แยกทีละกลุ่ม")
                k = 0
                cancelled = False
                for gi, (gcfg, grecs) in enumerate(groups, start=1):
                    if cancelled or (is_cancelled and is_cancelled()):
                        break
                    gtag = f"[กลุ่ม {gi}/{len(groups)}] " if multi else ""
                    log(f"[2/3] {gtag}เข้าสู่ระบบ: {gcfg['username']} "
                        f"({gcfg.get('user_type','')}) — {len(grecs)} รายการ")
                    try:
                        if gi > 1:
                            _logout_safely(page)
                            page.wait_for_timeout(800)
                        login(page, gcfg)
                        page.wait_for_timeout(1500)
                    except Exception as e:
                        log(f"      ✗ login ไม่สำเร็จ ({gcfg['username']}): {e} — "
                            f"ข้ามกลุ่มนี้ {len(grecs)} รายการ")
                        for rec in grecs:
                            k += 1
                            row = {**rec, "status": "LOGIN_FAIL",
                                   "note": f"login ไม่สำเร็จ ({gcfg['username']}): {str(e)[:150]}",
                                   "screenshot": "",
                                   "step2_status": "LOGIN_FAIL",
                                   "step2_note": f"ข้าม — login บัญชี {gcfg['username']} ไม่สำเร็จ"}
                            results.append(row)
                            _bt30_writeback_excel(excel_input, row, log=log)
                            if progress:
                                try: progress(k, len(selected))
                                except Exception: pass
                            _save_bt30_report(results, out_path, log=lambda *a: None)
                        continue
                    for rec in grecs:
                        if is_cancelled and is_cancelled():
                            log("[!] ผู้ใช้ยกเลิก — หยุด")
                            cancelled = True
                            break
                        k += 1
                        log(f"  ({k}/{len(selected)}) แถว {rec['row_index']}: "
                            f"{rec.get('prefix','')} {rec.get('name','')} | "
                            f"{rec.get('nationality','')} | {rec.get('sex','')} | {rec.get('birthdate','')}")
                        if skip_done and rec.get("done_request_no"):
                            log(f"      ⏭ ข้าม — ส่งคำขอไปแล้ว (เลขที่คำขอ {rec['done_request_no']})")
                            results.append({**rec, "status": "SKIP_DONE",
                                            "note": f"ส่งคำขอแล้ว (เลขที่คำขอ {rec['done_request_no']})",
                                            "screenshot": "",
                                            "step2_status": "SKIP_DONE",
                                            "step2_note": "ข้าม — ส่งคำขอแล้วก่อนหน้า",
                                            "submit_status": rec.get("done_submit_status") or "SUBMITTED",
                                            "request_no": rec["done_request_no"],
                                            "submit_pdf": ""})
                            if progress:
                                try: progress(k, len(selected))
                                except Exception: pass
                            _save_bt30_report(results, out_path, log=lambda *a: None)
                            continue
                        if not _open_bt30_form(page, log=log):
                            results.append({**rec, "status": "FORM_FAIL",
                                            "note": "เปิดฟอร์ม บต.30 ไม่สำเร็จ", "screenshot": ""})
                            _save_bt30_report(results, out_path, log=lambda *a: None)
                            continue
                        res = _bt30_fill_search_one(page, rec, screenshot_dir, log=log)
                        row = {**rec, **res}
                        log(f"      → ขั้นตอน1: {res.get('status')}"
                            + (f" | {res['note']}" if res.get("note") else ""))
                        if res.get("status") == "SUCCESS":
                            res2 = _bt30_do_step2(page, rec, screenshot_dir, log=log, do_submit=do_submit)
                            row.update(res2)
                            if res2.get("step2_status") == "SUCCESS":
                                success += 1
                            log(f"      → ขั้นตอน2: {res2.get('step2_status')}")
                        else:
                            row.update({"step2_status": "SKIP",
                                        "step2_note": "ข้ามขั้นตอน 2 เพราะขั้นตอน 1 ไม่สำเร็จ"})
                        results.append(row)
                        _bt30_writeback_excel(excel_input, row, log=log)
                        if progress:
                            try: progress(k, len(selected))
                            except Exception: pass
                        _save_bt30_report(results, out_path, log=lambda *a: None)
            else:
                # --- ขั้นตอน 1 เท่านั้น: login บัญชีหลัก เปิดฟอร์มครั้งเดียว วนเพิ่มทุกแถว ---
                log("[2/3] เข้าสู่ระบบ...")
                login(page, login_cfg)
                page.wait_for_timeout(1500)
                if not _open_bt30_form(page, log=log):
                    log("[!] เปิดฟอร์ม บต.30 ไม่สำเร็จ — ยุติ")
                    for rec in selected:
                        results.append({**rec, "status": "FORM_FAIL",
                                        "note": "เปิดฟอร์ม บต.30 ไม่สำเร็จ", "screenshot": ""})
                    _save_bt30_report(results, out_path, log=log)
                    return 0, out_path
                log("      ✓ เปิดฟอร์ม บต.30 สำเร็จ — เริ่มกรอกข้อมูลคนต่างด้าว")
                for k, rec in enumerate(selected, start=1):
                    if is_cancelled and is_cancelled():
                        log("[!] ผู้ใช้ยกเลิก — หยุด")
                        break
                    log(f"  ({k}/{len(selected)}) แถว {rec['row_index']}: "
                        f"{rec.get('prefix','')} {rec.get('name','')} | "
                        f"{rec.get('nationality','')} | {rec.get('sex','')} | {rec.get('birthdate','')}")
                    res = _bt30_fill_search_one(page, rec, screenshot_dir, log=log)
                    row1 = {**rec, **res}
                    results.append(row1)
                    _bt30_writeback_excel(excel_input, row1, log=log)
                    if res.get("status") == "SUCCESS":
                        success += 1
                    log(f"      → {res.get('status')}" + (f" | {res['note']}" if res.get("note") else ""))
                    if progress:
                        try: progress(k, len(selected))
                        except Exception: pass
                    _save_bt30_report(results, out_path, log=lambda *a: None)
        finally:
            ctx.close(); browser.close()

    _save_bt30_report(results, out_path, log=log)
    log(f"[3/3] สรุป: สำเร็จ {success} / {len(selected)} รายการ "
        f"(ดู screenshots ใน {screenshot_dir.name}/)")
    return success, out_path


# ════════════════════════════════════════════════════════════════════════════
#  โหมด บต.44 — การแจ้งการทำงาน และการยื่นคำขอเปลี่ยนรายการในใบอนุญาตทำงาน
#  ซึ่งไม่กระทบในใบอนุญาต (CHANGE_EMPLOYER) — ขั้นตอนที่ 1: ค้นหา+บันทึกข้อมูลคนต่างด้าว
#  ใช้ฟอร์มค้นหาคนต่างด้าวชุดเดียวกับ บต.30 ต่างกันแค่เมนูที่นำทางเข้าฟอร์ม
# ════════════════════════════════════════════════════════════════════════════
def _read_bt44_excel(path: Path) -> list[dict[str, Any]]:
    """อ่าน from_bt44.xlsx (ขั้นตอน 1-3)
    คอลัมน์: No., คำนำหน้า, หมายเลขอ้างอิงของคนต่างด้าว (ออปชัน), ชื่อ, สัญชาติ, เพศ, เกิดวันที่,
             เลขที่ใบอนุญาตทำงาน, เลขที่-ที่อยู่ที่ติดต่อได้, หมู่ที่/อาคาร, ซอย, ถนน, จังหวัด, เขต/อำเภอ, แขวง/ตำบล,
             ประเภทการค้นหา-เปลี่ยนนายจ้าง, ระบุ-เปลี่ยนนายจ้าง, เหตุผลการเปลี่ยนนายจ้าง,
             อื่น ๆ (โปรดระบุ), สถานที่ทำงาน/สาขา, ประเภทกิจการ, ประเภทงานที่ขออนุญาต, ลักษณะงาน
    """
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    hdr = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    def col(*names: str) -> int:
        for nm in names:
            for i, h in enumerate(hdr):
                if h == nm:
                    return i
        for nm in names:
            for i, h in enumerate(hdr):
                if nm and nm in h:
                    return i
        return -1

    i_no = col("No.", "No", "ลำดับ", "ลําดับ")
    i_prefix = col("คำนำหน้า", "คํานําหน้า", "Prefix", "Title")
    i_ref = col("หมายเลขอ้างอิงของคนต่างด้าว", "หมายเลขอ้างอิง", "เลขอ้างอิง", "alien_id")
    i_name = col("ชื่อ", "ชื่อ-สกุล", "Name")
    i_nat = col("สัญชาติ", "Nationality")
    i_sex = col("เพศ", "Sex", "Gender")
    i_birth = col("เกิดวันที่", "วันเกิด", "วันเดือนปีเกิด", "BirthDate", "DOB", "Birth")
    i_username = col("Username", "username", "ชื่อผู้ใช้")
    
    # Step 2 fields
    i_workpermit = col("เลขที่ใบอนุญาตทำงาน", "ใบอนุญาตทำงาน", "WorkPermitNo")
    i_addr_no = col("เลขที่-ที่อยู่ที่ติดต่อได้", "เลขที่")
    i_addr_moo = col("หมู่ที่/อาคาร-ที่อยู่ที่ติดต่อได้", "หมู่ที่")
    i_addr_soi = col("ซอย-ที่อยู่ที่ติดต่อได้", "ซอย")
    i_addr_road = col("ถนน-ที่อยู่ที่ติดต่อได้", "ถนน")
    i_addr_prov = col("จังหวัด-ที่อยู่ที่ติดต่อได้", "จังหวัด")
    i_addr_dist = col("เขต/อำเภอ-ที่อยู่ที่ติดต่อได้", "เขต/อำเภอ")
    i_addr_subdist = col("แขวง/ตำบล-ที่อยู่ที่ติดต่อได้", "แขวง/ตำบล")
    
    # Step 3 fields
    i_change_emp_search_type = col("ประเภทการค้นหา-เปลี่ยนนายจ้าง")
    i_change_emp_keyword = col("ระบุ-เปลี่ยนนายจ้าง")
    i_change_emp_reason = col("เหตุผลการเปลี่ยนนายจ้าง")
    i_change_emp_reason_other = col("อื่น ๆ (โปรดระบุ)")
    i_workplace_branch = col("สถานที่ทำงาน/สาขา", "สถานที่ทำงาน")
    i_work_biz = col("ประเภทกิจการ")
    i_work_permit_job = col("ประเภทงานที่ขออนุญาต")
    i_work_detail = col("ลักษณะงาน")

    # Step 4 — เอกสารแนบ (cols 23-33 เอกสารหลัก, 34-38 เอกสารอื่นๆ)
    # จับคู่ช่องอัปโหลดบนเว็บด้วย data-document-th (ค่า TH) = หัวคอลัมน์ Excel (ตัด ' *' ออก)
    # required อิงเครื่องหมาย * ในหัวคอลัมน์ + ช่องที่เว็บบังคับ (ใบอนุญาตทำงาน, รูปถ่าย ฯลฯ)
    _bt44_doc_specs = [
        # (header_base, required, is_photo, group)
        ("สำเนาเอกสารสำคัญประจำตัวของคนต่างด้าวที่ราชการออกให้", True, False, ""),
        ("สำเนาเอกสารหรือหลักฐานที่แสดงให้เห็นว่ามีการเปลี่ยนรายการในใบอนุญาตทำงานจริง", True, False, ""),
        ("ใบอนุญาตทำงาน", True, False, ""),
        ("สำเนาหนังสือเดินทาง", False, False, "passport"),
        ("สำเนาเอกสารใช้แทนหนังสือเดินทาง", False, False, "passport"),
        ("สำเนาหลักฐานการอนุญาตให้เข้ามาในราชอาณาจักร", False, False, "passport"),
        ("รูปถ่าย ขนาด 3 x 4 ซม.", True, True, ""),
        ("หนังสือมอบอำนาจซึ่งระบุข้อความมอบอำนาจให้ผู้รับอนุญาตนำคนต่างด้าวมาทำงานเป็นผู้ดำเนินการแทน", True, False, ""),
        ("ใบมอบอำนาจพร้อมติดอากรแสตมป์ของผู้รับมอบอำนาจแทนบริษัทนำเข้าคนต่างด้าว (เอกสารเพิ่มเติมสำหรับผู้กระทำการแทน)", True, False, ""),
        ("สำเนาบัตรประจำตัวประชาชนของผู้มอบอำนาจ (บริษัทนำเข้าคนต่างด้าว) (เอกสารเพิ่มเติมสำหรับผู้กระทำการแทน)", True, False, ""),
        ("สำเนาบัตรประจำตัวประชาชนของผู้รับมอบอำนาจแทนบริษัทนำเข้าคนต่างด้าว (เอกสารเพิ่มเติมสำหรับผู้กระทำการแทน)", True, False, ""),
    ]

    def _strip_star(s: str) -> str:
        s = (s or "").strip()
        while s and s[-1] in ("*", " ", "\u00a0"):
            s = s[:-1]
        return s.strip()

    def doc_col(base: str) -> int:
        tgt = _strip_star(base)
        for i, h in enumerate(hdr):
            if _strip_star(h) == tgt:
                return i
        return -1

    _bt44_doc_cols = [(doc_col(b), b, req, photo, grp) for (b, req, photo, grp) in _bt44_doc_specs]
    i_other_docs = [col(f"เอกสารอื่นๆที่เกี่ยวข้อง {n}") for n in range(1, 6)]
    base_dir = path.parent

    def resolve_path(raw: str) -> str:
        raw = (raw or "").strip().strip('"')
        if not raw:
            return ""
        p = Path(raw)
        if not p.is_absolute():
            p = base_dir / raw
        return str(p)

    def fmt_date(v: Any) -> str:
        if v is None or v == "":
            return ""
        if isinstance(v, (datetime, date)):
            return v.strftime("%d/%m/%Y")
        return str(v).strip()

    rows: list[dict[str, Any]] = []
    for ridx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=1):
        if not any(v not in (None, "") for v in row):
            continue

        def g(i: int) -> str:
            return str(row[i]).strip() if 0 <= i < len(row) and row[i] is not None else ""

        rec = {
            "seq": g(i_no) or str(ridx),
            "prefix": g(i_prefix),
            "alien_ref": g(i_ref),
            "name": g(i_name),
            "nationality": g(i_nat),
            "sex": g(i_sex),
            "birthdate": fmt_date(row[i_birth] if 0 <= i_birth < len(row) else None),
            "username": g(i_username),
            "row_index": ridx,
            # Step 2 fields
            "workpermit_no": g(i_workpermit),
            "addr_no": g(i_addr_no),
            "addr_moo": g(i_addr_moo),
            "addr_soi": g(i_addr_soi),
            "addr_road": g(i_addr_road),
            "addr_prov": g(i_addr_prov),
            "addr_dist": g(i_addr_dist),
            "addr_subdist": g(i_addr_subdist),
            # Step 3 fields
            "change_emp_search_type": g(i_change_emp_search_type),
            "change_emp_keyword": g(i_change_emp_keyword),
            "change_emp_reason": g(i_change_emp_reason),
            "change_emp_reason_other": g(i_change_emp_reason_other),
            "workplace_branch": g(i_workplace_branch),
            "work_biz": g(i_work_biz),
            "work_permit_job": g(i_work_permit_job),
            "work_detail": g(i_work_detail),
        }
        # Step 4 — เอกสารแนบ
        rec["docs"] = [
            {
                "th_name": _strip_star(base),
                "raw_path": (g(cidx) if cidx >= 0 else ""),
                "path": resolve_path(g(cidx) if cidx >= 0 else ""),
                "required": req,
                "is_photo": photo,
                "group": grp,
            }
            for (cidx, base, req, photo, grp) in _bt44_doc_cols
        ]
        rec["other_docs"] = [
            {"raw_path": g(cidx), "path": resolve_path(g(cidx))}
            for cidx in i_other_docs
            if cidx >= 0 and g(cidx)
        ]
        if rec["name"] or rec["birthdate"]:
            rows.append(rec)
    return rows


def _open_bt44_form(page: Page, log=print) -> bool:
    """เปิดฟอร์ม บต.44 (CHANGE_EMPLOYER) ผ่านเมนูบริการ
    flow: หน้าหลัก → 'เมนูบริการ' → openCity('tab_CHANGE_REQ') → คลิก #CHANGE_EMPLOYER
    (นำทางไป /WorkPermit?...ft=CHANGE_REQ — ใช้ฟอร์มค้นหาคนต่างด้าวชุดเดียวกับ บต.30)
    """
    try:
        page.goto("https://eworkpermit.doe.go.th/", wait_until="domcontentloaded", timeout=30_000)
        page.wait_for_timeout(2000)
    except Exception as e:
        log(f"      ✗ เปิดหน้าหลักไม่สำเร็จ: {e}")
        return False

    # 1) คลิก 'เมนูบริการ'
    clicked = page.evaluate(r"""() => {
        const a = document.querySelector('a.lang_menu_service')
          || Array.from(document.querySelectorAll('a,button'))
               .find(x => /เมนูบริการ/.test((x.textContent || '').trim()));
        if (a) { a.click(); return true; }
        return false;
    }""")
    if not clicked:
        log("      ✗ ไม่พบลิงก์ 'เมนูบริการ'")
        return False
    page.wait_for_timeout(1200)

    # 2) เปิดหมวด 'การยื่นขอเปลี่ยนรายการในใบอนุญาตทำงาน' (tab_CHANGE_REQ) แล้วคลิก #CHANGE_EMPLOYER
    page.evaluate(r"""() => {
        try { if (typeof openCity === 'function') openCity('tab_CHANGE_REQ', new Event('click')); } catch (e) {}
    }""")
    page.wait_for_timeout(900)
    ok = page.evaluate(r"""() => {
        const t = document.querySelector('#CHANGE_EMPLOYER');
        if (t) { t.click(); return true; }
        return false;
    }""")
    if not ok:
        log("      ✗ ไม่พบเมนู #CHANGE_EMPLOYER (แบบ บต.44)")
        return False

    # 3) รอเข้าหน้าฟอร์ม + ปุ่ม 'ค้นหาข้อมูลคนต่างด้าว' พร้อม
    try:
        page.wait_for_url("**/WorkPermit**", timeout=30_000)
    except PWTimeoutError:
        log(f"      ⚠ ยังไม่เข้าฟอร์ม (URL={page.url}) — ลองรอ element ต่อ")
    page.wait_for_timeout(2500)
    try:
        page.wait_for_function(
            r"""() => Array.from(document.querySelectorAll('button,a'))
                  .some(e => /search_alien_modal\.show/.test(e.getAttribute('onclick') || ''))""",
            timeout=20_000,
        )
        return True
    except PWTimeoutError:
        log(f"      ✗ เปิดฟอร์ม บต.44 ไม่สำเร็จ (URL={page.url})")
        return False


def _bt44_fill_search_one(
    page: Page,
    rec: dict[str, Any],
    screenshot_dir: Path,
    log=print,
) -> dict[str, Any]:
    """เปิด modal 'ค้นหาข้อมูลคนต่างด้าว' → กรอกข้อมูล 1 คน → กดบันทึก → จัดประเภทผลตาม modal
    - modal แจ้ง 'เสร็จสมบูรณ์/เรียบร้อยแล้ว' → SUCCESS (ดำเนินการคนถัดไป)
    - modal แจ้ง Error → ALERT/REVIEW (เก็บข้อความแจ้งเตือนใส่รายงาน แล้วข้ามไปคนถัดไป)
    คืน {status, note, screenshot}
    """
    seq = rec.get("seq", "?")
    name = rec.get("name", "")
    base = _safe_filename(f"{seq}_{name}")
    res: dict[str, Any] = {"status": "", "note": "", "screenshot": ""}

    try:
        # เปิด modal ค้นหาข้อมูลคนต่างด้าว
        opened = page.evaluate(r"""() => {
            const b = Array.from(document.querySelectorAll('button,a'))
              .find(e => /search_alien_modal\.show/.test(e.getAttribute('onclick') || ''));
            if (b) { b.click(); return true; }
            return false;
        }""")
        if not opened:
            res["status"] = "FAIL"
            res["note"] = "ไม่พบปุ่ม 'ค้นหาข้อมูลคนต่างด้าว'"
            return res
        page.wait_for_selector("#btn_search_alien_submit", state="visible", timeout=10_000)
        page.wait_for_timeout(600)

        # เคลียร์ทุก field ใน modal ก่อนกรอก (กัน state ค้างจาก record ก่อน)
        try:
            page.evaluate(r"""() => {
                const ids = ['alien_id','other_name','birthDateCheck'];
                for (const id of ids) {
                    const t = document.getElementById(id);
                    if (t) {
                        t.value = '';
                        t.dispatchEvent(new Event('input', { bubbles: true }));
                        t.dispatchEvent(new Event('change', { bubbles: true }));
                    }
                }
                // reset select dropdowns ให้ค่าว่าง (option แรก) + sync select2 ถ้ามี
                for (const sid of ['alien_prefix','nationality_al','sexCheck']) {
                    const s = document.getElementById(sid);
                    if (s && s.options.length > 0) {
                        s.selectedIndex = 0;
                        s.dispatchEvent(new Event('change', { bubbles: true }));
                        if (window.jQuery) {
                            try { jQuery(s).val('').trigger('change.select2'); } catch(e){}
                        }
                    }
                }
            }""")
        except Exception:
            pass
        page.wait_for_timeout(300)

        # หมายเลขอ้างอิงของคนต่างด้าว (#alien_id) — ออปชัน: กรอกถ้ามี / ล้างถ้าไม่มี (กันค่าค้างจากคนก่อน)
        page.evaluate(
            r"""(v) => {
              const t = document.getElementById('alien_id');
              if (t) {
                t.value = v;
                t.dispatchEvent(new Event('input', { bubbles: true }));
                t.dispatchEvent(new Event('change', { bubbles: true }));
              }
            }""", rec.get("alien_ref", "") or "",
        )
        # คำนำหน้า (select)
        if rec.get("prefix") and not _select_option_by_text(page, "alien_prefix", rec["prefix"]):
            log(f"      ⚠ เลือกคำนำหน้า '{rec['prefix']}' ไม่ได้")
        # ชื่อ (text) — ใส่ชื่อเต็มตาม Excel
        page.evaluate(
            r"""(v) => {
              const t = document.getElementById('other_name');
              if (t) {
                t.value = v;
                t.dispatchEvent(new Event('input', { bubbles: true }));
                t.dispatchEvent(new Event('change', { bubbles: true }));
              }
            }""", name,
        )
        # สัญชาติ (select)
        if rec.get("nationality") and not _select_option_by_text(page, "nationality_al", rec["nationality"]):
            log(f"      ⚠ เลือกสัญชาติ '{rec['nationality']}' ไม่ได้")
        # เพศ (select) — try multiple selectors
        sex_val = rec.get("sex", "")
        if sex_val:
            sex_ok = _select_option_by_text(page, "sexCheck", sex_val)
            if not sex_ok:
                # fallback: try radio button หรือ button group
                sex_ok = bool(page.evaluate(r"""(want) => {
                    const norm = s => (s||'').replace(/\s+/g,' ').trim().toLowerCase();
                    const w = norm(want);
                    // ลองหา radio/checkbox ที่มี label match
                    for (const rb of document.querySelectorAll('input[type="radio"], input[type="checkbox"]')) {
                        if (!rb.offsetParent) continue;
                        const lbl = rb.closest('label') || (rb.id ? document.querySelector(`label[for="${rb.id}"]`) : null);
                        if (lbl && norm(lbl.textContent).includes(w)) {
                            if (!rb.checked) rb.click();
                            return true;
                        }
                    }
                    return false;
                }""", sex_val))
            if not sex_ok:
                log(f"      ⚠ เลือกเพศ '{sex_val}' ไม่ได้")
        # วันเกิด (datepicker text, dd/mm/yyyy ค.ศ.)
        if rec.get("birthdate"):
            page.evaluate(
                r"""(v) => {
                  const t = document.getElementById('birthDateCheck');
                  if (t) {
                    t.removeAttribute('readonly');
                    t.value = v;
                    t.dispatchEvent(new Event('input', { bubbles: true }));
                    t.dispatchEvent(new Event('change', { bubbles: true }));
                    t.dispatchEvent(new Event('blur', { bubbles: true }));
                    if (window.jQuery) { try { jQuery(t).trigger('change'); } catch (e) {} }
                  }
                }""", rec["birthdate"],
            )
        page.wait_for_timeout(400)
        try:
            page.screenshot(path=str(screenshot_dir / f"{base}_filled.png"), full_page=True)
            res["screenshot"] = f"{base}_filled.png"
        except Exception:
            pass

        # กดบันทึก (#btn_search_alien_submit)
        page.evaluate(r"""() => { const b = document.getElementById('btn_search_alien_submit'); if (b) b.click(); }""")

        # รอผลแบบสมาร์ท: รอจนกว่าจะมี 'alert/swal modal โผล่' หรือ 'modal ค้นหาปิด'
        # (เดิม wait แบบ static 5s — record 2+ บางครั้ง AJAX ยังไม่เสร็จ → false positive)
        try:
            page.wait_for_function(
                r"""() => {
                    for (const sel of ['.swal2-popup', '.modal.show .alert', '.alert.show', '.toast.show']) {
                        for (const el of document.querySelectorAll(sel)) {
                            if (el.offsetParent !== null) return true;
                        }
                    }
                    const modal = document.getElementById('search_alien_modal');
                    return !(modal && modal.offsetParent !== null);
                }""",
                timeout=15_000,
            )
        except Exception:
            pass
        page.wait_for_timeout(1500)

        alert = _capture_register_alert(page)
        errs = _bt30_field_errors(page)
        modal_open = page.evaluate(
            r"""() => { const m = document.getElementById('search_alien_modal'); return !!(m && m.offsetParent !== null); }"""
        )

        # DEBUG: เก็บ snippet ของหน้าหลัง submit เพื่อ diagnose ว่าฟอร์มหลักมี text อะไรจริงๆ
        try:
            page_snippet = page.evaluate(
                r"""() => {
                    const t = document.body.innerText || '';
                    return t.replace(/\s+/g, ' ').slice(0, 400);
                }"""
            )
            log(f"      [Step1-DEBUG] modal_open={modal_open} | alert={alert!r} | "
                f"page_snippet={page_snippet[:300]!r}")
        except Exception:
            pass

        # บันทึก screenshot หลังกด (เก็บภาพ modal แจ้งเตือนไว้ในรายงาน)
        try:
            page.screenshot(path=str(screenshot_dir / f"{base}_aftersave.png"), full_page=True)
            res["screenshot"] = f"{base}_aftersave.png"
        except Exception:
            pass

        if alert and any(k in alert for k in ("เสร็จสมบูรณ์", "เรียบร้อยแล้ว", "เรียบร้อย", "สำเร็จแล้ว")):
            # 'การค้นหาข้อมูลคนต่างด้าวเสร็จสมบูรณ์' = สำเร็จ → ดำเนินการคนถัดไป
            res["status"] = "SUCCESS"
            res["note"] = alert[:400]
            _close_register_alert(page)  # กดปุ่ม 'ปิด' — ไม่กด 'ยินยอม' (อยู่นอกขอบเขต Step 1)
        elif alert and any(k in alert for k in ("ไม่พบ", "ไม่ถูกต้อง", "ผิดพลาด", "ไม่สำเร็จ", "ไม่สามารถ", "ซ้ำ", "กรอกข้อมูล")):
            # modal แจ้ง Error → เก็บข้อความไว้ทำรายงาน แล้วข้ามไปคนถัดไป
            res["status"] = "ALERT"
            res["note"] = alert[:400]
            _close_register_alert(page)
        elif alert:
            res["status"] = "REVIEW"
            res["note"] = alert[:400]
            _close_register_alert(page)
        elif errs:
            res["status"] = "VALIDATE"
            res["note"] = f"ฟอร์มแจ้งเตือน: {errs}"[:300]
        elif not modal_open:
            # modal ปิด → ถือเป็น SUCCESS (กลับมาเหมือน behavior เดิมที่ record 1 เคยทำงาน)
            res["status"] = "SUCCESS"
            res["note"] = "บันทึกค้นหา (modal ปิด)"
        else:
            res["status"] = "REVIEW"
            res["note"] = "modal ยังเปิดอยู่หลังกดบันทึก — ตรวจสอบ screenshot"

        # ถ้าไม่สำเร็จ ปิด modal ค้นหาที่ค้างอยู่ เพื่อเริ่มกรอกคนถัดไปแบบสะอาด
        if res["status"] != "SUCCESS":
            try:
                page.evaluate(r"""() => {
                    const b = Array.from(document.querySelectorAll('button,a'))
                      .find(e => /search_alien_modal\.close/.test(e.getAttribute('onclick') || ''));
                    if (b) b.click();
                }""")
                page.wait_for_timeout(500)
            except Exception:
                pass
        return res
    except Exception as e:
        try:
            page.screenshot(path=str(screenshot_dir / f"{base}_error.png"), full_page=True)
            res["screenshot"] = f"{base}_error.png"
        except Exception:
            pass
        res["status"] = "ERROR"
        res["note"] = str(e)[:300]
        return res


# ====================== Helper: Loading Protection ======================

def _wait_loading_disappeared(page: Page, timeout_ms: int = 12000, log=print, allow_modal: bool = False) -> bool:
    """รอให้ loading spinner หายไป (อาจจะช้า/เร็วตามเว็บ)
    - ทดสอบหลายรูปแบบของ spinner
    - Retry หลายครั้ง
    - ป้องกันการหมดเวลาโดยคืนค่า True ถ้ารอนานเกินไป
    - allow_modal=True: ไม่นับ Bootstrap modal / SweetAlert popup ว่าเป็น loading
      (แก้บั๊ก: เปิด modal อยู่แล้ว detector เข้าใจผิดว่าเป็น overlay loading → รอจนครบ timeout)
    """
    start = time.time()
    max_wait = timeout_ms / 1000.0
    
    while time.time() - start < max_wait:
        try:
            # ตรวจสอบ spinner/overlay หลายรูปแบบ รวมทั้ง loading แบบจุดกลางจอ
            has_spinner = page.evaluate(r"""(allowModal) => {
                const isVisible = (elem) => {
                    if (!elem) return false;
                    const style = window.getComputedStyle(elem);
                    return elem.offsetParent !== null &&
                        style.display !== 'none' &&
                        style.visibility !== 'hidden' &&
                        style.opacity !== '0';
                };
                // element นี้เป็น/อยู่ใน modal หรือ swal popup หรือไม่ (ไม่ใช่ loading จริง)
                const MODAL_SEL = '.modal, .modal-dialog, .modal-content, [role="dialog"], .swal2-popup, .swal2-modal, .swal2-container';
                const inModal = (el) => !!(el && el.closest && el.closest(MODAL_SEL));
                const modalVisible = () => {
                    const ms = document.querySelectorAll('.modal.show, .modal.in, .modal-dialog, [role="dialog"], .swal2-popup, .swal2-container');
                    for (const m of ms) { if (isVisible(m)) return true; }
                    return false;
                };
                const modalOpen = allowModal && modalVisible();

                // (1) ตัวบ่งชี้ "loading จริง" — class/aria เฉพาะเจาะจง → ตรวจเสมอ แม้อยู่ใน modal
                //     เช่น spinner ระหว่าง AJAX ตอนกด 'ค้นหานายจ้าง' จะยังถูกจับได้ ไม่ถูกข้าม
                const realSpinner = [
                    '[class*="loading"]',
                    '[class*="spinner"]',
                    '.loader',
                    '[role="progressbar"]',
                    '[aria-busy="true"]'
                ];
                for (const selector of realSpinner) {
                    for (const elem of document.querySelectorAll(selector)) {
                        if (isVisible(elem)) return true;
                    }
                }

                // (2) overlay/backdrop ที่กำกวม — ถ้าเป็น chrome ของ modal เองให้ข้ามเมื่อ allowModal
                //     แต่ overlay loading เต็มจอ (ไม่ได้อยู่ใน .modal) จะยังถูกจับได้ตามปกติ
                const ambiguous = ['.overlay'];
                if (!allowModal) {
                    ambiguous.push('[class*="modal-backdrop"]', '[class*="fade in"]');
                }
                for (const selector of ambiguous) {
                    for (const elem of document.querySelectorAll(selector)) {
                        if (!isVisible(elem)) continue;
                        if (allowModal && inModal(elem)) continue;  // chrome ของ modal ไม่ใช่ loading
                        return true;
                    }
                }

                // ตรวจสอบ element กลางจอที่ดูเหมือน spinner (เช่น จุดวิ่ง) — ตรวจได้แม้ใน modal
                // (เนื้อหา modal ที่เป็นข้อความ/ฟอร์มจะไม่เข้าเงื่อนไข spinner ด้านล่าง จึงปลอดภัย)
                const cx = Math.floor(window.innerWidth / 2);
                const cy = Math.floor(window.innerHeight / 2);
                const center = document.elementFromPoint(cx, cy);
                if (center) {
                    const cls = (center.className || '').toString().toLowerCase();
                    const style = window.getComputedStyle(center);
                    const rect = center.getBoundingClientRect();
                    const isDotLike = rect.width <= 40 && rect.height <= 40 &&
                        (style.borderRadius.includes('50%') || style.borderRadius.includes('999'));

                    if (isVisible(center) && (
                        /load|spin|overlay|progress/.test(cls) ||
                        center.getAttribute('role') === 'progressbar' ||
                        center.closest('[class*="loading"], [class*="spinner"], .loader, [aria-busy="true"]') ||
                        isDotLike
                    )) {
                        return true;
                    }
                }

                // ตรวจสอบมี overlay ใหญ่ปิดจอ (ข้ามถ้ามี modal เปิดอยู่ — overlay คือ backdrop ของ modal)
                if (!modalOpen) {
                    const vw = window.innerWidth;
                    const vh = window.innerHeight;
                    const all = document.querySelectorAll('body *');
                    for (const el of all) {
                        if (!isVisible(el)) continue;
                        if (allowModal && (inModal(el) || (el.querySelector && el.querySelector(MODAL_SEL)))) continue;
                        const style = window.getComputedStyle(el);
                        if (!['fixed', 'absolute'].includes(style.position)) continue;
                        const rect = el.getBoundingClientRect();
                        const areaRatio = (rect.width * rect.height) / Math.max(1, (vw * vh));
                        const z = Number.parseInt(style.zIndex || '0', 10);
                        if (areaRatio >= 0.35 && z >= 10) return true;
                    }
                }

                return false;
            }""", allow_modal)
            
            if not has_spinner:
                elapsed = time.time() - start
                log(f"      ✓ Loading finished ({elapsed:.1f}s)")
                return True
                
            page.wait_for_timeout(300)  # Check every 300ms
            
        except Exception as e:
            log(f"      ⚠ Loading check error: {str(e)[:80]}")
            page.wait_for_timeout(500)
    
    elapsed = time.time() - start
    log(f"      ⚠ Loading timeout ({elapsed:.1f}s/{timeout_ms}ms) - continuing anyway")
    return False


def _screenshot_when_ready(page: Page, path: Path, timeout_ms: int = 15000, log=print, allow_modal: bool = False) -> bool:
    """Capture screenshot only when page is visually ready (no center loading spinner/overlay)."""
    try:
        _wait_loading_disappeared(page, timeout_ms=timeout_ms, log=log, allow_modal=allow_modal)

        start = time.time()
        max_wait = timeout_ms / 1000.0
        while time.time() - start < max_wait:
            center_busy = page.evaluate(r"""(allowModal) => {
                const cx = Math.floor(window.innerWidth / 2);
                const cy = Math.floor(window.innerHeight / 2);
                const el = document.elementFromPoint(cx, cy);
                if (!el) return false;
                if (allowModal && el.closest && el.closest('.modal, .modal-dialog, [role="dialog"], .swal2-popup, .swal2-container')) return false;

                const style = window.getComputedStyle(el);
                const cls = (el.className || '').toString().toLowerCase();
                const txt = (el.textContent || '').trim();
                const visible = el.offsetParent !== null && style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0';
                if (!visible) return false;

                if (/load|spin|overlay|progress/.test(cls)) return true;
                if (el.getAttribute('role') === 'progressbar' || el.closest('[aria-busy="true"]')) return true;
                if (/^\.{3,}$/.test(txt) || /\u2022/.test(txt)) return true;
                return false;
            }""", allow_modal)
            if not center_busy:
                break
            page.wait_for_timeout(300)

        page.screenshot(path=str(path), full_page=True)
        return True
    except Exception:
        return False


def _safe_click_element(page: Page, find_script: str, element_name: str, retries: int = 3, log=print) -> bool:
    """Click element with visibility check, scroll-into-view, and retries
    
    Args:
        find_script: JavaScript code that assigns element to 'elem' variable
                    e.g. "const elem = document.getElementById('btn');"
    """
    for attempt in range(retries):
        try:
            # First: find and scroll to element
            found = page.evaluate(f"""() => {{
                {find_script}
                if (!elem) return false;
                elem.scrollIntoView({{ behavior: 'instant', block: 'center' }});
                return true;
            }}""")
            
            if not found:
                log(f"      ⚠ ไม่พบ {element_name} (attempt {attempt + 1}/{retries})")
                if attempt < retries - 1:
                    page.wait_for_timeout(400)
                continue
            
            page.wait_for_timeout(200)
            
            # Second: verify visibility and click
            clicked = page.evaluate(f"""() => {{
                {find_script}
                if (!elem) return false;
                
                // Check visibility
                const style = window.getComputedStyle(elem);
                if (style.display === 'none' || style.visibility === 'hidden' || 
                    style.opacity === '0' || elem.offsetParent === null) {{
                    return false;
                }}

                // Must be top-most at element center to avoid click through overlay
                const r = elem.getBoundingClientRect();
                const cx = Math.floor(r.left + (r.width / 2));
                const cy = Math.floor(r.top + (r.height / 2));
                const topElem = document.elementFromPoint(cx, cy);
                if (!topElem || (topElem !== elem && !elem.contains(topElem))) {{
                    return false;
                }}
                
                // Click
                elem.click();
                return true;
            }}""")
            
            if clicked:
                log(f"      ✓ Clicked '{element_name}'")
                return True
            
            log(f"      ⚠ {element_name} ไม่ visible หรือ click ไม่สำเร็จ")
            if attempt < retries - 1:
                page.wait_for_timeout(400)
                
        except Exception as e:
            log(f"      ⚠ Error clicking {element_name}: {str(e)[:80]}")
            if attempt < retries - 1:
                page.wait_for_timeout(400)
    
    return False


def _bt44_step2_consent(page: Page, log=print) -> bool:
    """ขั้นตอน 2.1: หา checkbox ข้อมูลใจสำสัญญา แล้วกด ถัดไป
    ข้อความ: 'ข้าพเจ้าขอรับรองว่า มีความประสงค์ในการยื่นคำขอใบอนุญาตแทนคนต่างด้าวหรือนายจ้าง...'
    
    ป้องกัน Loading: - ใช้ wait_loading_disappeared() เพื่ออรประเป็นการอร wait loading ได้อย่างเหมาะสม
    - Multiple retries สำหรับการหา element
    - Visibility verification ก่อน click
    """
    try:
        log("      Step 2.1: Checking consent...")

        # Step 1: เช็ก checkbox แบบ user interaction (เหมือนคนกดจริง)
        checkbox_clicked = False
        for attempt in range(1, 4):
            try:
                chk = page.locator("#check_truth").first
                if chk.count() > 0:
                    chk.scroll_into_view_if_needed()
                    # check() เป็น interaction จริงพร้อม firing events
                    if not chk.is_checked():
                        chk.check(timeout=5000, force=True)
                    checkbox_clicked = True
                
                # Fallback: กด label ถ้า check() ไม่สำเร็จ
                if not checkbox_clicked or (chk.count() > 0 and not chk.is_checked()):
                    lbl = page.locator("label", has_text="ข้าพเจ้าขอรับรอง").first
                    if lbl.count() > 0:
                        lbl.scroll_into_view_if_needed()
                        lbl.click(timeout=5000, force=True)
                        checkbox_clicked = True

                # ยืนยันผลหลัง interaction
                if chk.count() > 0 and chk.is_checked():
                    checkbox_clicked = True
                    break
            except Exception:
                checkbox_clicked = False

            page.wait_for_timeout(350)

        # Verify ว่าติ๊กสำเร็จจริง
        checkbox_checked = False
        try:
            checkbox_checked = page.locator("#check_truth").first.is_checked()
        except Exception:
            checkbox_checked = bool(page.evaluate(r"""() => {
                const chk = document.getElementById('check_truth');
                return !!(chk && chk.checked);
            }"""))

        if not checkbox_clicked or not checkbox_checked:
            log("      ⚠ ไม่สามารถติ๊ก checkbox ข้อมูลใจสำสัญญาแบบ user click ได้")
            return False

        log("      ✓ Consent checkbox clicked by user-like interaction")
        page.wait_for_timeout(500)

        # ปิดโมดัลที่อาจค้างอยู่ก่อนกดถัดไป
        try:
            page.evaluate(r"""() => {
                const visibleModals = Array.from(document.querySelectorAll('.modal, .swal2-popup, [role="dialog"]'))
                  .filter(m => m.offsetParent !== null);
                for (const m of visibleModals) {
                    const btn = Array.from(m.querySelectorAll('button,a'))
                      .find(b => /ปิด|ยืนยัน|ตกลง|ok/i.test((b.textContent || '').trim()) && b.offsetParent !== null);
                    if (btn) btn.click();
                }
            }""")
            page.wait_for_timeout(400)
        except Exception:
            pass

        # Step 2: รอ loading รอบแรก
        _wait_loading_disappeared(page, timeout_ms=10000, log=log)

        # Step 3: คลิกปุ่ม 'ถัดไป' แบบ user-like (Playwright click) ก่อน
        click_info = None
        for attempt in range(1, 5):
            try:
                next_btn = page.locator("#gonextSubmit").first
                next_btn.click(timeout=5000, force=True)
                click_info = {
                    "clicked": True,
                    "reason": "ok",
                    "text": "ถัดไป",
                    "id": "gonextSubmit",
                    "cls": ""
                }
                break
            except Exception:
                pass

            # Fallback: ค้นหาปุ่มถัดไปจาก DOM
            click_info = page.evaluate(r"""() => {
            const chk = document.getElementById('check_truth') ||
              Array.from(document.querySelectorAll('input[type="checkbox"]'))
              .find(c => {
                const lbl = c.closest('label') || document.querySelector(`label[for="${c.id}"]`);
                const txt = (lbl?.textContent || c.parentElement?.textContent || '').trim();
                return /ข้าพเจ้าขอรับรอง/.test(txt);
              });
            if (!chk) return { clicked: false, reason: 'checkbox-not-found' };

            const chkRect = chk.getBoundingClientRect();
            const candidates = Array.from(document.querySelectorAll('button, a, [role="button"]'))
              .filter(b => {
                const txt = (b.textContent || '').trim();
                if (!/ถัดไป|NEXT|Next/i.test(txt)) return false;
                const style = window.getComputedStyle(b);
                return b.offsetParent !== null && style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0';
              })
              .map(b => {
                const r = b.getBoundingClientRect();
                const dx = (r.left + r.width / 2) - (chkRect.left + chkRect.width / 2);
                const dy = (r.top + r.height / 2) - (chkRect.top + chkRect.height / 2);
                return { b, dist: Math.sqrt(dx * dx + dy * dy) };
              })
              .sort((x, y) => x.dist - y.dist);

            if (!candidates.length) return { clicked: false, reason: 'next-not-found' };

            const btn = candidates[0].b;
            btn.scrollIntoView({ behavior: 'instant', block: 'center' });

            try {
                btn.click();
            } catch (e) {
                btn.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
            }

            return {
                clicked: true,
                reason: 'ok',
                text: (btn.textContent || '').trim(),
                id: btn.id || '',
                cls: (btn.className || '').toString().slice(0, 80)
            };
        }""")

            if click_info and click_info.get("clicked"):
                break

            reason = (click_info or {}).get("reason", "unknown")
            log(f"      ⚠ Click 'ถัดไป' attempt {attempt}/4 failed ({reason})")

            # ถ้าถูก overlay บัง ให้รอ loading แล้วลองใหม่
            if reason in ("next-covered-by-overlay", "next-not-found"):
                _wait_loading_disappeared(page, timeout_ms=6000, log=log)
            page.wait_for_timeout(600)

        if not click_info or not click_info.get("clicked"):
            reason = (click_info or {}).get("reason", "unknown")
            log(f"      ⚠ Click 'ถัดไป' failed ({reason})")
            return False

        log(
            "      ✓ Clicked Next | "
            f"text='{click_info.get('text', '')}' id='{click_info.get('id', '')}' class='{click_info.get('cls', '')}'"
        )

        # Step 4: รอ loading หลัง click
        page.wait_for_timeout(1000)
        _wait_loading_disappeared(page, timeout_ms=12000, log=log)
        page.wait_for_timeout(500)

        # Step 5: ยืนยันว่า transition ออกจากหน้า checkbox แล้วจริง
        ready = page.evaluate(r"""() => {
            const isVisible = (el) => {
                if (!el) return false;
                const st = window.getComputedStyle(el);
                return el.offsetParent !== null && st.display !== 'none' && st.visibility !== 'hidden' && st.opacity !== '0';
            };

            const consentChk = document.getElementById('check_truth');
            const consentStillVisible = isVisible(consentChk);

            const headingEls = Array.from(document.querySelectorAll('h1,h2,h3,h4,h5,label,span,div,section'));
            const hasThaiAddressHeading = headingEls.some(el => {
                const txt = (el.textContent || '').trim();
                return /ที่อยู่ในประเทศไทย/.test(txt) && isVisible(el);
            });

            const hasEmployerSection = headingEls.some(el => {
                const txt = (el.textContent || '').trim();
                return /ข้อมูลนายจ้างที่ต้องการดำเนินการแทน/.test(txt) && isVisible(el);
            });

            const hasEditInAddressBlock = Array.from(document.querySelectorAll('button,a,[role="button"],span,div'))
              .some(el => {
                if (!isVisible(el)) return false;
                const txt = (el.textContent || '').trim();
                if (!/แก้ไขข้อมูล|แก้ไข|edit/i.test(txt)) return false;
                const block = el.closest('section,div,fieldset');
                const btxt = (block?.textContent || '').trim();
                return /ที่อยู่ในประเทศไทย/.test(btxt);
              });

            return { consentStillVisible, hasThaiAddressHeading, hasEditInAddressBlock, hasEmployerSection };
        }""")

        if ready.get("consentStillVisible"):
            try:
                dbg_dir = Path("reports") / "bt44_screenshots"
                dbg_dir.mkdir(parents=True, exist_ok=True)
                page.screenshot(path=str(dbg_dir / "step2_1_transition_fail_consent_visible.png"), full_page=True)
            except Exception:
                pass
            log("      ⚠ หลังคลิก 'ถัดไป' ยังอยู่หน้า consent (checkbox ยังมองเห็น)")
            return False

        # บางเคสจะเข้าหน้า "ข้อมูลนายจ้าง" ก่อน ต้องกดถัดไปอีกครั้ง
        if (not ready.get("hasThaiAddressHeading")
                and not ready.get("hasEditInAddressBlock")
                and ready.get("hasEmployerSection")):
            log("      ⚠ อยู่หน้าข้อมูลนายจ้าง — ลองกด 'ถัดไป' อีกครั้ง")
            extra_next = page.evaluate(r"""() => {
                const btn = Array.from(document.querySelectorAll('button,a,[role="button"]'))
                    .find(b => /ถัดไป|NEXT|Next/i.test((b.textContent || '').trim()) && b.offsetParent !== null);
                if (!btn) return false;
                btn.click();
                return true;
            }""")
            if extra_next:
                page.wait_for_timeout(1000)
                _wait_loading_disappeared(page, timeout_ms=10000, log=log)
                ready = page.evaluate(r"""() => {
                    const isVisible = (el) => {
                        if (!el) return false;
                        const st = window.getComputedStyle(el);
                        return el.offsetParent !== null && st.display !== 'none' && st.visibility !== 'hidden' && st.opacity !== '0';
                    };
                    const hasThaiAddressHeading = Array.from(document.querySelectorAll('h1,h2,h3,h4,h5,label,span,div,section'))
                        .some(el => /ที่อยู่ในประเทศไทย/.test((el.textContent || '').trim()) && isVisible(el));
                    const hasEditInAddressBlock = Array.from(document.querySelectorAll('button,a,[role="button"],span,div'))
                        .some(el => {
                            if (!isVisible(el)) return false;
                            const txt = (el.textContent || '').trim();
                            if (!/แก้ไขข้อมูล|แก้ไข|edit/i.test(txt)) return false;
                            const block = el.closest('section,div,fieldset');
                            const btxt = (block?.textContent || '').trim();
                            return /ที่อยู่ในประเทศไทย/.test(btxt);
                        });
                    return { hasThaiAddressHeading, hasEditInAddressBlock };
                }""")

        if not ready.get("hasThaiAddressHeading") and not ready.get("hasEditInAddressBlock"):
            try:
                dbg_dir = Path("reports") / "bt44_screenshots"
                dbg_dir.mkdir(parents=True, exist_ok=True)
                page.screenshot(path=str(dbg_dir / "step2_1_transition_fail_no_address_section.png"), full_page=True)
                headings = page.evaluate(r"""() => Array.from(document.querySelectorAll('h1,h2,h3,h4,h5'))
                    .map(h => (h.textContent || '').trim())
                    .filter(Boolean)
                    .slice(0, 8)""")
                log(f"      DEBUG headings after next: {headings}")
            except Exception:
                pass
            log("      ⚠ หลังคลิก 'ถัดไป' ยังไม่เจอ section ที่อยู่ในประเทศไทย")
            return False

        log("      ✓ Step 2.1 confirmed (transitioned to address step)")
        return True

    except Exception as e:
        log(f"      ✗ Error in Step 2.1: {str(e)[:200]}")
        return False


def _bt44_step2_edit_address(page: Page, rec: dict[str, str], screenshot_dir: Path, log=print) -> dict[str, str]:
    """ขั้นตอน 2.2: แก้ไขที่อยู่ในส่วน 'ที่อยู่ในประเทศไทย'
    
    ป้องกัน Loading:
    - รอให้ loading หายไป ก่อน click ปุ่ม
    - Multiple retries สำหรับการหา element
    - Visibility verification ก่อน interaction
    - Enhanced button finding: ค้นหาปุ่มแบบรอบด้านยิ่งขึ้น
    """
    res = {"status": "", "note": ""}
    try:
        log("      Step 2.2: Editing address...")
        
        # Wait for page loading after Step 2.1 navigation
        _wait_loading_disappeared(page, timeout_ms=10000, log=log)
        page.wait_for_timeout(500)

        # หา และกด ปุ่ม 'แก้ไขข้อมูล' ของส่วน "ที่อยู่ในประเทศไทย" โดยเฉพาะ
        # (ห้ามไปจับปุ่มแก้ไขของส่วนอื่น เช่น ใบอนุญาตทำงาน, นายจ้าง, สถานประกอบการ)
        edit_btn_script = r"""
            let btn = null;

            // ฟังก์ชันช่วย: ดึง "ข้อความเฉพาะของ element นั้น" (ไม่รวมข้อความลูก)
            const ownText = (el) => Array.from(el.childNodes || [])
                .filter(n => n.nodeType === 3)
                .map(n => (n.textContent || '').trim())
                .join(' ').trim();

            // 1) หา heading element ที่ข้อความตรง "ที่อยู่ในประเทศไทย"
            let heading = null;
            const headEls = document.querySelectorAll(
                'h1,h2,h3,h4,h5,h6,strong,b,legend,label,div,span,p'
            );
            for (const el of headEls) {
                const t = ownText(el);
                if (t === 'ที่อยู่ในประเทศไทย' || t === 'ที่อยู่ปัจจุบัน' || t === 'ที่อยู่ที่ติดต่อได้') {
                    heading = el;
                    break;
                }
            }

            // 2) ถ้าเจอ heading → walk up parents (ระยะ 6 ชั้น) แล้วหาปุ่ม
            //    "แก้ไขข้อมูล"/"แก้ไข" ที่อยู่ใน scope ใกล้กันเท่านั้น
            const looksLikeEdit = (cand) => {
                const txt = (cand.textContent || '').trim();
                if (!txt || txt.length > 40) return false;
                if (/^แก้ไขข้อมูล$/.test(txt)) return true;
                if (/^แก้ไข$/.test(txt)) return true;
                // กันชนกับปุ่ม "แก้ไขใบอนุญาตทำงาน", "แก้ไขเปลี่ยนนายจ้าง", "แก้ไขสถานประกอบการ"
                if (/^แก้ไขข้อมูล\s*$/.test(txt)) return true;
                return false;
            };

            if (heading) {
                let scope = heading.parentElement;
                for (let i = 0; i < 6 && scope && !btn; i++) {
                    const candidates = scope.querySelectorAll(
                        'button, a, [role="button"], [onclick], [ng-click]'
                    );
                    for (const cand of candidates) {
                        if (looksLikeEdit(cand)) { btn = cand; break; }
                    }
                    scope = scope.parentElement;
                }
            }

            // 3) Fallback: หา button ทั่วหน้าที่ข้อความตรง "แก้ไขข้อมูล" เป๊ะ
            //    (ถ้ามีอันเดียวก็ใช้, ถ้ามีหลายอันให้เลือกอันที่ใกล้ heading ที่สุด)
            if (!btn) {
                const exact = [];
                const all = document.querySelectorAll('button, a, [role="button"], [onclick]');
                for (const cand of all) {
                    const txt = (cand.textContent || '').trim();
                    if (txt === 'แก้ไขข้อมูล') exact.push(cand);
                }
                if (exact.length === 1) {
                    btn = exact[0];
                } else if (exact.length > 1 && heading) {
                    // เลือกอันที่ DOM-distance ใกล้ heading ที่สุด
                    const hRect = heading.getBoundingClientRect();
                    let bestDist = Infinity;
                    for (const c of exact) {
                        const cr = c.getBoundingClientRect();
                        const dy = Math.abs(cr.top - hRect.top);
                        const dx = Math.abs(cr.left - hRect.left);
                        const d = dy * 4 + dx;  // weight แนวตั้งมากกว่า
                        if (d < bestDist) { bestDist = d; btn = c; }
                    }
                } else if (exact.length > 1) {
                    btn = exact[0];
                }
            }

            elem = btn || null;
        """
        
        if not _safe_click_element(page, edit_btn_script, "ปุ่ม 'แก้ไขข้อมูล'", retries=4, log=log):
            log("      ⚠ ไม่พบปุ่ม 'แก้ไขข้อมูล' - ข้ามไป Step 2.3")
            res["status"] = "ALERT"
            res["note"] = "ไม่พบปุ่มแก้ไขที่อยู่ - ข้ามไป"
            return res

        # รอ modal เปิด + พร้อมกรอก — ใช้ rect-based visibility (รองรับ position:fixed
        # ที่ offsetParent เป็น null เสมอ → เดิมรอจนครบ 8s ฟรี ๆ ทำให้ช้า ~15s)
        try:
            page.wait_for_function(
                r"""() => {
                    const vis = (el) => {
                        if (!el) return false;
                        const st = getComputedStyle(el);
                        if (st.display === 'none' || st.visibility === 'hidden' || st.opacity === '0') return false;
                        const r = el.getBoundingClientRect();
                        return r.width > 0 && r.height > 0;
                    };
                    // โมดัลที่มองเห็น (รองรับทั้ง .modal มาตรฐานและ custom dialog)
                    const sels = '.modal, .modal-dialog, .modal-content, [role="dialog"], .swal2-popup';
                    for (const m of document.querySelectorAll(sels)) { if (vis(m)) return true; }
                    // หรือมีปุ่ม 'บันทึก' ที่มองเห็นแล้ว = โมดัลพร้อม
                    for (const b of document.querySelectorAll('button, a')) {
                        if (/บันทึก/.test((b.textContent || '').trim()) && vis(b)) return true;
                    }
                    return false;
                }""",
                timeout=4000,
            )
        except Exception:
            pass
        _wait_loading_disappeared(page, timeout_ms=3000, log=log, allow_modal=True)

        # ---- กรอกข้อมูลที่อยู่จาก Excel ----
        # แยกเป็น 4 จังหวะ: text inputs → จังหวัด (รอ AJAX) → อำเภอ (รอ AJAX) → ตำบล
        addr_data = {
            "addr_no": str(rec.get("addr_no", "") or "").strip(),
            "addr_moo": str(rec.get("addr_moo", "") or "").strip(),
            "addr_soi": str(rec.get("addr_soi", "") or "").strip(),
            "addr_road": str(rec.get("addr_road", "") or "").strip(),
            "addr_prov": str(rec.get("addr_prov", "") or "").strip(),
            "addr_dist": str(rec.get("addr_dist", "") or "").strip(),
            "addr_subdist": str(rec.get("addr_subdist", "") or "").strip(),
        }
        log(f"      [Step2.2-DATA] addr_no={addr_data['addr_no']!r} moo={addr_data['addr_moo']!r} "
            f"prov={addr_data['addr_prov']!r} dist={addr_data['addr_dist']!r} subdist={addr_data['addr_subdist']!r}")

        # 1) Text inputs (เลขที่/หมู่/ซอย/ถนน) — match ด้วย name+id+placeholder+label
        text_filled = page.evaluate(
            r"""(d) => {
                const out = { addr_no: false, addr_moo: false, addr_soi: false, addr_road: false };
                // หา input ที่อยู่ใน modal เปิดอยู่ก่อน (ถ้าไม่มี modal ใช้ทั้ง document)
                const visMod = Array.from(document.querySelectorAll(
                    '.modal.show, .modal[style*="block"], .modal-dialog:not([style*="display: none"]), [role="dialog"]'
                )).find(m => {
                    const r = m.getBoundingClientRect();
                    return r.width > 0 && r.height > 0;
                });
                const root = visMod || document;
                const inputs = root.querySelectorAll('input[type="text"], input:not([type]), textarea');

                const labelOf = (inp) => {
                    let lab = '';
                    if (inp.id) {
                        const l = document.querySelector(`label[for="${inp.id}"]`);
                        if (l) lab = (l.textContent || '').trim();
                    }
                    if (!lab) {
                        const wrap = inp.closest('.form-group, .form-row, .row, div');
                        if (wrap) {
                            const l = wrap.querySelector('label');
                            if (l) lab = (l.textContent || '').trim();
                        }
                    }
                    return lab;
                };

                const fire = (inp) => {
                    inp.dispatchEvent(new Event('input', { bubbles: true }));
                    inp.dispatchEvent(new Event('change', { bubbles: true }));
                    inp.dispatchEvent(new Event('blur', { bubbles: true }));
                };

                for (const inp of inputs) {
                    const name = (inp.name || inp.id || '').toLowerCase();
                    const placeholder = (inp.placeholder || '').toLowerCase();
                    const label = labelOf(inp);
                    const combined = (name + ' ' + placeholder + ' ' + label).toLowerCase();

                    // ข้าม input ที่ไม่ใช่ของฟอร์มที่อยู่ (เช่น search box)
                    if (inp.type === 'hidden' || inp.disabled) continue;

                    // 'เลขที่' — ระวังชนกับ 'เลขที่เอกสาร', 'เลขประจำตัว', 'เลขที่ใบอนุญาต'
                    if (!out.addr_no && d.addr_no &&
                        (/(^|\s)เลขที่(\s|$)/.test(label) ||
                         /address.*no|addr.*no|house.*no|no\.?$/.test(name) ||
                         (combined.includes('เลขที่') &&
                          !combined.includes('เอกสาร') &&
                          !combined.includes('ใบอนุญาต') &&
                          !combined.includes('ประจำตัว') &&
                          !combined.includes('นิติ')))) {
                        inp.value = d.addr_no; fire(inp); out.addr_no = true; continue;
                    }
                    if (!out.addr_moo && d.addr_moo &&
                        (combined.includes('หมู่') || combined.includes('moo') || combined.includes('village'))) {
                        inp.value = d.addr_moo; fire(inp); out.addr_moo = true; continue;
                    }
                    if (!out.addr_soi && d.addr_soi &&
                        (combined.includes('ซอย') || combined.includes('soi') || combined.includes('alley'))) {
                        inp.value = d.addr_soi; fire(inp); out.addr_soi = true; continue;
                    }
                    if (!out.addr_road && d.addr_road &&
                        (combined.includes('ถนน') || combined.includes('road') || combined.includes('street'))) {
                        inp.value = d.addr_road; fire(inp); out.addr_road = true; continue;
                    }
                }
                return out;
            }""",
            addr_data,
        )
        log(f"      [Step2.2-FILL-TEXT] {text_filled}")

        # 2) Dropdown chain: prov → wait → dist → wait → subdist (Dice fuzzy match)
        def _select_addr_dropdown(kind: str, want: str) -> dict:
            """kind: 'prov' | 'dist' | 'subdist' """
            if not want:
                return {"ok": False, "reason": "no_want"}
            keywords_map = {
                "prov": ["จังหวัด", "province", "provid"],
                "dist": ["อำเภอ", "เขต", "district", "amphur", "amphoe"],
                "subdist": ["ตำบล", "แขวง", "subdistrict", "tambon", "tambol"],
            }
            anti_keywords_map = {
                "prov": ["อำเภอ", "เขต", "ตำบล", "แขวง"],
                "dist": ["จังหวัด", "ตำบล", "แขวง"],
                "subdist": ["จังหวัด", "อำเภอ"],
            }
            return page.evaluate(
                r"""({ keys, antiKeys, want }) => {
                    const norm = (s) => (s || '').replace(/\s+/g, ' ').trim().toLowerCase();
                    const compact = (s) => (s || '').replace(/\s+/g, '').toLowerCase();
                    const bigrams = (s) => {
                        const t = compact(s); const out = new Map();
                        for (let i = 0; i < t.length - 1; i++) {
                            const bg = t.slice(i, i + 2);
                            out.set(bg, (out.get(bg) || 0) + 1);
                        }
                        return out;
                    };
                    const dice = (a, b) => {
                        const A = bigrams(a), B = bigrams(b);
                        if (A.size === 0 || B.size === 0) return 0;
                        let inter = 0, totA = 0, totB = 0;
                        for (const [, v] of A) totA += v;
                        for (const [, v] of B) totB += v;
                        for (const [bg, va] of A) {
                            const vb = B.get(bg);
                            if (vb) inter += Math.min(va, vb);
                        }
                        return (2 * inter) / (totA + totB);
                    };

                    // หา select ที่มองเห็น + ตรงประเภท (มี keyword + ไม่มี anti-keyword)
                    const visMod = Array.from(document.querySelectorAll(
                        '.modal.show, .modal[style*="block"], .modal-dialog:not([style*="display: none"]), [role="dialog"]'
                    )).find(m => {
                        const r = m.getBoundingClientRect();
                        return r.width > 0 && r.height > 0;
                    });
                    const root = visMod || document;
                    const selects = Array.from(root.querySelectorAll('select')).filter(s => {
                        const r = s.getBoundingClientRect();
                        return r.width > 0 && r.height > 0 && !s.disabled;
                    });

                    const labelOf = (sel) => {
                        let lab = '';
                        if (sel.id) {
                            const l = document.querySelector(`label[for="${sel.id}"]`);
                            if (l) lab = (l.textContent || '').trim();
                        }
                        if (!lab) {
                            const wrap = sel.closest('.form-group, .form-row, .row, div');
                            if (wrap) {
                                const l = wrap.querySelector('label');
                                if (l) lab = (l.textContent || '').trim();
                            }
                        }
                        return lab;
                    };

                    let target = null;
                    for (const sel of selects) {
                        const ctx = (sel.name + ' ' + sel.id + ' ' + labelOf(sel)).toLowerCase();
                        const hasKey = keys.some(k => ctx.includes(k.toLowerCase()));
                        const hasAnti = antiKeys.some(k => ctx.includes(k.toLowerCase()));
                        if (hasKey && !hasAnti) { target = sel; break; }
                    }
                    if (!target) return { ok: false, reason: 'no_select' };

                    const opts = Array.from(target.options).filter(o => o.value && o.value !== '');
                    if (opts.length === 0) return { ok: false, reason: 'no_option', selectId: target.id || target.name };

                    const w = norm(want);
                    let best = null, bestScore = -1;
                    for (const o of opts) {
                        const txt = norm(o.textContent || '');
                        if (!txt) continue;
                        let score;
                        if (txt === w) score = 1.0;
                        else if (txt.includes(w) || w.includes(txt)) score = 0.95;
                        else score = dice(txt, w);
                        if (score > bestScore) { bestScore = score; best = o; }
                    }
                    if (!best || bestScore < 0.4) {
                        return {
                            ok: false, reason: 'no_match', bestScore,
                            bestText: best ? best.textContent.trim() : '',
                            optCount: opts.length,
                            sampleOpts: opts.slice(0, 5).map(o => o.textContent.trim()),
                        };
                    }
                    target.value = best.value;
                    target.dispatchEvent(new Event('input', { bubbles: true }));
                    target.dispatchEvent(new Event('change', { bubbles: true }));
                    // ถ้าเป็น select2/jQuery ลอง trigger ผ่าน $ ด้วย
                    try {
                        if (window.jQuery) {
                            window.jQuery(target).trigger('change');
                        }
                    } catch (e) {}
                    return {
                        ok: true,
                        score: bestScore,
                        chosen: best.textContent.trim(),
                        selectId: target.id || target.name,
                    };
                }""",
                {"keys": keywords_map[kind], "antiKeys": anti_keywords_map[kind], "want": want},
            )

        # 2a) จังหวัด
        prov_res = _select_addr_dropdown("prov", addr_data["addr_prov"])
        log(f"      [Step2.2-FILL-PROV] {prov_res}")
        # รอ AJAX โหลด options ของอำเภอ
        if prov_res.get("ok"):
            try:
                page.wait_for_function(
                    r"""() => {
                        const sels = Array.from(document.querySelectorAll('select')).filter(s => {
                            const ctx = (s.name + ' ' + s.id + ' ').toLowerCase();
                            const lab = (() => {
                                if (s.id) {
                                    const l = document.querySelector(`label[for="${s.id}"]`);
                                    if (l) return (l.textContent || '').toLowerCase();
                                }
                                return '';
                            })();
                            const all = ctx + lab;
                            return /(อำเภอ|เขต|district|amphur|amphoe)/.test(all)
                                && !/จังหวัด|ตำบล|แขวง/.test(all);
                        });
                        // อย่างน้อย 1 select มี options > 1 (รวม placeholder)
                        return sels.some(s => Array.from(s.options).filter(o => o.value).length > 0);
                    }""",
                    timeout=8000,
                )
            except Exception:
                log("      ⚠ Timeout รอ options อำเภอโหลด")
            page.wait_for_timeout(400)

        # 2b) อำเภอ
        dist_res = _select_addr_dropdown("dist", addr_data["addr_dist"])
        log(f"      [Step2.2-FILL-DIST] {dist_res}")
        if dist_res.get("ok"):
            try:
                page.wait_for_function(
                    r"""() => {
                        const sels = Array.from(document.querySelectorAll('select')).filter(s => {
                            const ctx = (s.name + ' ' + s.id + ' ').toLowerCase();
                            const lab = (() => {
                                if (s.id) {
                                    const l = document.querySelector(`label[for="${s.id}"]`);
                                    if (l) return (l.textContent || '').toLowerCase();
                                }
                                return '';
                            })();
                            const all = ctx + lab;
                            return /(ตำบล|แขวง|subdistrict|tambon|tambol)/.test(all)
                                && !/จังหวัด|อำเภอ/.test(all);
                        });
                        return sels.some(s => Array.from(s.options).filter(o => o.value).length > 0);
                    }""",
                    timeout=8000,
                )
            except Exception:
                log("      ⚠ Timeout รอ options ตำบลโหลด")
            page.wait_for_timeout(400)

        # 2c) ตำบล
        subdist_res = _select_addr_dropdown("subdist", addr_data["addr_subdist"])
        log(f"      [Step2.2-FILL-SUBDIST] {subdist_res}")

        # debug: ถ้า dropdown หาไม่เจอเลย → dump form HTML ลงไฟล์เพื่อ probe
        if (not prov_res.get("ok")) or (not dist_res.get("ok")) or (not subdist_res.get("ok")):
            try:
                form_html = page.evaluate(
                    r"""() => {
                        const m = Array.from(document.querySelectorAll(
                            '.modal.show, .modal[style*="block"], [role="dialog"]'
                        )).find(x => {
                            const r = x.getBoundingClientRect();
                            return r.width > 0 && r.height > 0;
                        });
                        return (m || document.body).outerHTML;
                    }"""
                )
                debug_path = screenshot_dir / f"step2_2_debug_modal_{rec.get('seq','?')}.html"
                debug_path.write_text(form_html, encoding="utf-8")
                log(f"      [Step2.2-DEBUG] dump modal HTML → {debug_path}")
            except Exception as e:
                log(f"      [Step2.2-DEBUG] dump failed: {e!r}")

        addr_filled = {
            **{k: v for k, v in (text_filled or {}).items()},
            "addr_prov": bool(prov_res.get("ok")),
            "addr_dist": bool(dist_res.get("ok")),
            "addr_subdist": bool(subdist_res.get("ok")),
        }

        page.wait_for_timeout(300)

        # Screenshot หลังกรอก (modal เปิดอยู่ → allow_modal=True ไม่ต้องรอนาน)
        _screenshot_when_ready(page, screenshot_dir / "step2_address_filled.png", timeout_ms=4000, log=log, allow_modal=True)

        # กดปุ่ม 'บันทึก' — เลือกเฉพาะปุ่มที่ "มองเห็นได้" + timeout สั้น
        # (บั๊กเดิม: locator ทั่วหน้า .last ไปโดนปุ่มบันทึกที่ซ่อนอยู่ และ
        #  scroll_into_view_if_needed() ไม่ใส่ timeout → ใช้ค่า default 30s → ค้าง ~30 วิ)
        save_clicked = False
        try:
            save_btn = page.locator("button:visible:has-text('บันทึก'), a:visible:has-text('บันทึก')").last
            if save_btn.count() > 0:
                save_btn.scroll_into_view_if_needed(timeout=2000)
                save_btn.click(timeout=4000, force=True)
                save_clicked = True
        except Exception:
            save_clicked = False

        if not save_clicked:
            try:
                save_clicked = bool(page.evaluate(r"""() => {
                    const candidates = Array.from(document.querySelectorAll('button, a'));
                    const btn = candidates.find(el => {
                        const txt = (el.textContent || '').trim();
                        const st = window.getComputedStyle(el);
                        return /บันทึก/.test(txt) && el.offsetParent !== null && st.display !== 'none' && st.visibility !== 'hidden';
                    });
                    if (!btn) return false;
                    btn.scrollIntoView({ block: 'center' });
                    btn.click();
                    return true;
                }"""))
            except Exception:
                save_clicked = False

        if not save_clicked:
            log("      ⚠ ไม่พบปุ่ม 'บันทึก'")
            res["status"] = "ALERT"
            res["note"] = "ไม่พบปุ่มบันทึกที่อยู่"
            return res

        log("      ✓ Clicked 'บันทึก'")

        page.wait_for_timeout(800)
        _wait_loading_disappeared(page, timeout_ms=5000, log=log, allow_modal=True)

        # Screenshot หลังบันทึก (allow_modal=True → ไม่รอ swal นาน)
        _screenshot_when_ready(page, screenshot_dir / "step2_address_saved.png", timeout_ms=5000, log=log, allow_modal=True)

        # ต้องเจอโมดัลสำเร็จแล้วกด 'ยืนยัน'
        confirm_clicked = False
        try:
            page.wait_for_function(
                r"""() => {
                    const bodyText = document.body.innerText || '';
                    return /สำเร็จ/.test(bodyText) && /ยืนยัน/.test(bodyText);
                }""",
                timeout=8000,
            )

            confirm_btn = page.locator("button:visible:has-text('ยืนยัน'), a:visible:has-text('ยืนยัน')").last
            if confirm_btn.count() > 0:
                confirm_btn.scroll_into_view_if_needed(timeout=2000)
                confirm_btn.click(timeout=4000, force=True)
                confirm_clicked = True
        except Exception:
            confirm_clicked = False

        if not confirm_clicked:
            try:
                confirm_clicked = bool(page.evaluate(r"""() => {
                    const visible = (el) => {
                        if (!el) return false;
                        const st = window.getComputedStyle(el);
                        return el.offsetParent !== null && st.display !== 'none' && st.visibility !== 'hidden' && st.opacity !== '0';
                    };
                    const buttons = Array.from(document.querySelectorAll('button, a'));
                    const btn = buttons.find(el => visible(el) && /ยืนยัน|ตกลง|ปิด/i.test((el.textContent || '').trim()));
                    if (!btn) return false;
                    btn.click();
                    return true;
                }"""))
            except Exception:
                confirm_clicked = False

        if not confirm_clicked:
            log("      ⚠ ไม่พบโมดัลบันทึกสำเร็จหรือปุ่ม 'ยืนยัน'")
            res["status"] = "ALERT"
            res["note"] = "ไม่พบโมดัลบันทึกสำเร็จ/ปุ่มยืนยัน"
            return res

        log("      ✓ กด 'ยืนยัน' ในโมดัลบันทึกสำเร็จแล้ว")
        page.wait_for_timeout(800)

        res["status"] = "SUCCESS"
        res["note"] = "แก้ไขที่อยู่และยืนยันบันทึกเสร็จ"
        log("      ✓ Address edited successfully")
        return res

    except Exception as e:
        log(f"      ✗ Error in Step 2.2: {str(e)[:200]}")
        res["status"] = "ERROR"
        res["note"] = str(e)[:200]
        return res


def _bt44_step2_verify_permit(page: Page, rec: dict[str, str], log=print) -> tuple[bool, str]:
    """ขั้นตอน 2.3: ตรวจสอบว่าเลขที่ใบอนุญาตในฟอร์มตรงกับ Excel หรือไม่
    Return: (match_ok, current_permit_number)
    """
    try:
        expected_permit = "".join(ch for ch in (rec.get("workpermit_no", "") or "") if ch.isdigit())

        # เก็บ candidate permit จากหลายแหล่ง แล้ว normalize เป็นตัวเลขล้วน
        candidates = page.evaluate(r"""() => {
            const outPermitContext = [];
            const outGlobal = [];
            const pushNumsFromText = (txt) => {
                const list = [];
                if (!txt) return;
                // รองรับรูปแบบมีช่องว่าง/ขีด เช่น 5692-0005-5150
                const chunks = txt.match(/(?:\d[\d\s\-]{9,}\d)/g) || [];
                for (const c of chunks) {
                    const digits = c.replace(/\D/g, '');
                    if (digits.length >= 10 && digits.length <= 16) list.push(digits);
                }
                const direct = txt.match(/\d{10,16}/g) || [];
                for (const d of direct) list.push(d);
                return list;
            };

            // 1) เน้นใน section ใบอนุญาตทำงาน
            const permitSections = Array.from(document.querySelectorAll('div,section,fieldset,table'))
                .filter(sec => /ข้อมูลใบอนุญาตทำงานปัจจุบัน|ใบอนุญาตทำงาน|work\s*permit|permit/i.test((sec.textContent || '').trim()));
            for (const sec of permitSections) {
                outPermitContext.push(...(pushNumsFromText(sec.textContent || '') || []));
                for (const el of sec.querySelectorAll('input,span,p,label,td,th,div')) {
                    outPermitContext.push(...(pushNumsFromText((el.value || el.textContent || '').trim()) || []));
                }
            }

            // 2) label ใกล้เคียงคำว่า ใบอนุญาตทำงาน เท่านั้น
            const labelish = Array.from(document.querySelectorAll('label,th,td,span,div,p'));
            for (const el of labelish) {
                const txt = (el.textContent || '').trim();
                if (/ใบอนุญาตทำงาน|work\s*permit|permit/i.test(txt)) {
                    outPermitContext.push(...(pushNumsFromText(txt) || []));
                    const parent = el.closest('tr,div,section,fieldset');
                    if (parent) outPermitContext.push(...(pushNumsFromText(parent.textContent || '') || []));
                }
            }

            // 3) ทั้งหน้า
            outGlobal.push(...(pushNumsFromText(document.body.innerText || '') || []));

            // unique preserve order
            const seen = new Set();
            const uniqPermit = outPermitContext.filter(v => {
                if (seen.has(v)) return false;
                seen.add(v);
                return true;
            });

            const seenGlobal = new Set();
            const uniqGlobal = outGlobal.filter(v => {
                if (seenGlobal.has(v)) return false;
                seenGlobal.add(v);
                return true;
            });

            return { permitContext: uniqPermit, global: uniqGlobal };
        }""")

        # เลือก candidate ที่ยาว 12 เฉพาะ permit context
        cands = [c for c in ((candidates or {}).get("permitContext") or []) if isinstance(c, str)]
        cands12 = [c for c in cands if len(c) == 12]
        global_cands = [c for c in ((candidates or {}).get("global") or []) if isinstance(c, str)]

        has_permit_context = len(cands) > 0

        def _same_permit(a: str, b: str) -> bool:
            if not a or not b:
                return False
            return a == b or a.lstrip("0") == b.lstrip("0")

        current_permit = ""
        if expected_permit:
            # เจอเลข expected ใน candidate ใด ๆ ให้ถือว่าตรง
            for c in cands:
                if _same_permit(c, expected_permit):
                    current_permit = c
                    break

        if not current_permit and has_permit_context:
            if cands12:
                current_permit = cands12[0]
            elif cands:
                current_permit = cands[0]

        # ถ้าไม่เจอใน permit context ให้ถือว่าไม่พบ (ไม่ใช้เลขทั่วหน้าเพื่อลด false positive)
        if not current_permit and expected_permit and expected_permit in global_cands:
            current_permit = expected_permit

        if not current_permit:
            log("      ⚠ ไม่พบเลขที่ใบอนุญาต ในฟอร์ม")
            return (False, "")

        match_ok = bool(expected_permit) and _same_permit(current_permit, expected_permit)
        if match_ok:
            log(f"      ✓ เลขที่ใบอนุญาตตรงกัน: {current_permit} (expected {expected_permit})")
            return (True, current_permit)

        log(f"      ✗ เลขที่ใบอนุญาตไม่ตรง — คาดหวัง: {expected_permit}, ได้: {current_permit}")
        return (False, current_permit)

    except Exception as e:
        log(f"      ✗ Error in Step 2.3 permit verify: {str(e)[:200]}")
        return (False, str(e)[:100])


def _bt44_step3_change_employer(page: Page, rec: dict[str, Any], log=print) -> dict[str, str]:
    """Step 3.1: เปลี่ยนนายจ้างและค้นหาตามข้อมูลใน Excel
    - ถ้าค้นหาแล้วพบ modal 'ไม่สามารถดำเนินการต่อได้...' ให้ mark SKIP
    - ถ้าค้นหาไม่สำเร็จในเชิงเทคนิค ให้ mark ALERT
    """
    res = {"status": "", "note": ""}
    try:
        search_type = (rec.get("change_emp_search_type") or "").strip()
        keyword = (rec.get("change_emp_keyword") or "").strip()
        if not search_type or not keyword:
            res["status"] = "ALERT"
            res["note"] = "Step 3.1: missing employer search inputs"
            return res

        log("      [Step 3.1] ค้นหานายจ้าง...")
        opened = False
        try:
            page.locator("#changeEmployer").first.click(timeout=5000, force=True)
            opened = True
        except Exception:
            pass
        if not opened:
            try:
                opened = bool(page.evaluate(r"""() => {
                    const btn = document.querySelector('#changeEmployer') ||
                      Array.from(document.querySelectorAll('button,a,[role="button"]'))
                        .find(b => /เปลี่ยนนายจ้าง/.test((b.textContent || '').trim()) && b.offsetParent !== null);
                    if (!btn) return false;
                    btn.click();
                    return true;
                }"""))
            except Exception:
                opened = False

        if not opened:
            res["status"] = "ALERT"
            res["note"] = "Step 3.1: open employer modal failed"
            return res

        page.wait_for_timeout(600)
        _wait_loading_disappeared(page, timeout_ms=5000, log=log, allow_modal=True)

        if not _select2_pick(page, "search-type", search_type, log=log):
            res["status"] = "ALERT"
            res["note"] = f"Step 3.1: employer search type not found ({search_type})"
            return res
        page.wait_for_timeout(500)

        # กรอก keyword แบบพิมพ์จริง (fill) — เว็บอ่านค่าจาก event จริง ไม่ใช่ .value ที่ set ด้วย JS
        try:
            kw = page.locator("#search-keyword").first
            kw.scroll_into_view_if_needed(timeout=4000)
            kw.click(timeout=4000)
            kw.fill("")
            kw.fill(keyword)
        except Exception:
            page.evaluate(r"""(v) => {
                const t = document.getElementById('search-keyword');
                if (!t) return;
                t.value = v;
                t.dispatchEvent(new Event('input', { bubbles: true }));
                t.dispatchEvent(new Event('change', { bubbles: true }));
            }""", keyword)
        page.wait_for_timeout(300)

        # กดปุ่ม 'ค้นหา' แบบคลิกจริง (auto-scroll + actionability) — เลี่ยง JS click ที่ไม่กระตุ้น handler
        clicked = False
        try:
            sa = page.locator("#search-action").first
            sa.scroll_into_view_if_needed(timeout=4000)
            sa.click(timeout=6000)
            clicked = True
        except Exception:
            # fallback: JS click
            clicked = bool(page.evaluate(r"""() => {
                const btn = document.getElementById('search-action');
                if (!btn) return false;
                btn.scrollIntoView({ block: 'center' });
                btn.click();
                return true;
            }"""))
        if not clicked:
            res["status"] = "ALERT"
            res["note"] = "Step 3.1: employer search click failed"
            return res

        try:
            page.wait_for_function(
                r"""() => {
                    const txt = document.body.innerText || '';
                    return /ข้อมูลนายจ้างใหม่/.test(txt)
                        || /ไม่สามารถดำเนินการต่อได้ เนื่องจากไม่พบข้อมูลนายจ้างในระบบ/.test(txt)
                        || /ข้อมูลนายจ้างเดิม/.test(txt)
                        || /คืนค่าเดิม/.test(txt)
                        || /นายจ้างหลัก/.test(txt);
                }""",
                timeout=8000,
            )
        except Exception:
            pass
        _wait_loading_disappeared(page, timeout_ms=8000, log=log, allow_modal=True)

        alert = _capture_register_alert(page)
        body_text = page.evaluate(r"""() => (document.body.innerText || '').slice(0, 5000)""") or ""
        not_found_text = "ไม่สามารถดำเนินการต่อได้ เนื่องจากไม่พบข้อมูลนายจ้างในระบบ"
        if not_found_text in alert or not_found_text in body_text:
            _close_register_alert(page)
            res["status"] = "SKIP"
            res["note"] = "Step 3.1: employer not found or not eligible"
            return res

        # ---- เลือกเหตุผลการเปลี่ยนนายจ้าง (reason-type) ในโมดัล ก่อนบันทึก ----
        # ⚠ สำคัญ: หลังกด 'ค้นหา' เว็บหน่วง 2-3 วิรอข้อมูลนายจ้างกลับมา ระหว่างนั้น
        #   ส่วนเลือกเหตุผล/ช่อง 'อื่นๆ (โปรดระบุ)' ยังไม่พร้อม ถ้า Select เร็วเกินไป
        #   change handler จะไม่ผูก → ช่อง #reason-other ไม่โผล่ → กรอกไม่ได้ → บันทึกไม่ผ่าน
        #   จึงต้อง (1) รอ select reason-type พร้อมก่อน (2) หลังเลือก 'อื่นๆ' รอ #reason-other โผล่จริง
        reason = (rec.get("change_emp_reason") or "").strip()
        reason_other = (rec.get("change_emp_reason_other") or "").strip()

        # (1) รอให้ select 'เหตุผลการเปลี่ยนนายจ้าง' พร้อม (มี option จริง) + settle หน่วงเพิ่ม
        try:
            page.wait_for_function(
                r"""() => {
                    const s = document.getElementById('reason-type');
                    return s && s.options && s.options.length > 1;
                }""",
                timeout=8000,
            )
        except Exception:
            pass
        page.wait_for_timeout(1500)  # settle: เผื่อ change handler ผูกหลังข้อมูลค้นกลับมา

        if reason:
            if not _select2_pick(page, "reason-type", reason, log=log):
                log(f"      ⚠ Step 3.1: เลือกเหตุผล '{reason}' ไม่ได้")
            page.wait_for_timeout(800)

        # (2) ถ้ามีข้อความ 'อื่นๆ (โปรดระบุ)' → ต้องรอช่อง #reason-other โผล่จริงก่อนกรอก
        #     ถ้ายังไม่โผล่ ให้เลือกเหตุผลซ้ำเพื่อ re-trigger change (สูงสุด 3 รอบ)
        if reason_other:
            ro_filled = False
            for attempt in range(3):
                try:
                    ro = page.locator("#reason-other:visible").first
                    ro.wait_for(state="visible", timeout=4000)
                    ro.scroll_into_view_if_needed(timeout=2000)
                    ro.click(timeout=3000)
                    ro.fill("")
                    ro.fill(reason_other)
                    val = ro.input_value(timeout=2000)
                    if (val or "").strip() == reason_other:
                        ro_filled = True
                        log(f"      [Step 3.1] กรอกเหตุผลอื่นๆ: {reason_other}")
                        break
                except Exception:
                    pass
                # re-trigger: เลือกเหตุผลซ้ำเพื่อให้ change event ยิงอีกครั้ง
                if reason:
                    _select2_pick(page, "reason-type", reason, log=lambda *a: None)
                page.wait_for_timeout(1200)
            if not ro_filled:
                # JS fallback: เซ็ตค่าโดยตรงบน element ที่มองเห็น
                rr = page.evaluate(r"""(val) => {
                    const vis = el => { if(!el) return false; const st=getComputedStyle(el); const rc=el.getBoundingClientRect(); return st.display!=='none'&&st.visibility!=='hidden'&&rc.width>0&&rc.height>0; };
                    const els = Array.from(document.querySelectorAll('#reason-other, [name="reason-other"]')).filter(vis);
                    const el = els[els.length - 1];
                    if (!el) return { ok:false };
                    el.focus(); el.value = val;
                    el.dispatchEvent(new Event('input', { bubbles:true }));
                    el.dispatchEvent(new Event('change', { bubbles:true }));
                    el.dispatchEvent(new Event('blur', { bubbles:true }));
                    return { ok:true };
                }""", reason_other)
                ro_filled = bool(rr and rr.get("ok"))
            if not ro_filled:
                res["status"] = "ALERT"
                res["note"] = "Step 3.1: ช่อง 'อื่นๆ (โปรดระบุ)' ไม่โผล่/กรอกไม่ได้ (เลือกเหตุผลเร็วเกินไป?)"
                return res
            page.wait_for_timeout(400)

        # ---- บันทึกการเปลี่ยนนายจ้าง (ปุ่ม 'บันทึก' ในโมดัล) — ใช้ :visible + timeout สั้น ----
        save_clicked = False
        try:
            save_btn = page.locator("button:visible:has-text('บันทึก'), a:visible:has-text('บันทึก')").last
            if save_btn.count() > 0:
                save_btn.scroll_into_view_if_needed(timeout=2000)
                save_btn.click(timeout=4000, force=True)
                save_clicked = True
        except Exception:
            save_clicked = False
        if not save_clicked:
            try:
                save_clicked = bool(page.evaluate(r"""() => {
                    const btn = Array.from(document.querySelectorAll('button, a')).find(el => {
                        const st = getComputedStyle(el);
                        return /บันทึก/.test((el.textContent || '').trim()) && el.offsetParent !== null && st.display !== 'none' && st.visibility !== 'hidden';
                    });
                    if (!btn) return false;
                    btn.scrollIntoView({ block: 'center' });
                    btn.click();
                    return true;
                }"""))
            except Exception:
                save_clicked = False
        if not save_clicked:
            res["status"] = "ALERT"
            res["note"] = "Step 3.1: save (บันทึก) button not found"
            return res

        page.wait_for_timeout(800)
        _wait_loading_disappeared(page, timeout_ms=6000, log=log, allow_modal=True)
        # ปิด swal สำเร็จ (ถ้ามี) แล้วยืนยันว่าโมดัลปิด + เข้าสู่ขั้น 'เลือกสถานที่ทำงาน'
        try:
            page.wait_for_function(
                r"""() => /สำเร็จ|เลือกสถานที่ทำงาน|ข้อมูลนายจ้างเดิม/.test(document.body.innerText || '')""",
                timeout=6000,
            )
        except Exception:
            pass
        _close_register_alert(page)
        _wait_loading_disappeared(page, timeout_ms=4000, log=log, allow_modal=True)

        log("      [Step 3.1] ✅ เปลี่ยนนายจ้างและบันทึกแล้ว")
        res["status"] = "SUCCESS"
        res["note"] = "Step 3.1 complete (employer changed + saved)"
        return res
    except Exception as e:
        res["status"] = "ERROR"
        res["note"] = f"Step 3.1 error: {str(e)[:180]}"
        return res


def _bt44_step3_workplace(page: Page, rec: dict[str, Any], log=print) -> dict[str, str]:
    """Step 3.3: เลือกสถานที่ทำงาน/สาขา + ประเภทกิจการ + ตรวจสอบประเภทงาน + กรอกลักษณะงาน
    Step 3.4: กดถัดไป → Step 4

    - 3.3:   เปิดโมดัล 'เลือกสถานที่ทำงาน' แล้วเลือกสาขา (Excel: สถานที่ทำงาน/สาขา)
    - 3.3.1: เลือกประเภทกิจการ (Excel: ประเภทกิจการ) แล้วตรวจสอบประเภทงานที่ขออนุญาต
             ต้องตรงกับ Excel (ประเภทงานที่ขออนุญาต) — ถ้าไม่ตรง → SKIP (mark report)
    - 3.3.2: กรอกลักษณะงาน (Excel: ลักษณะงาน)
    - 3.4:   กดถัดไป
    """
    res = {"status": "", "note": ""}

    def _close_workplace_modal():
        try:
            page.evaluate(r"""() => {
                const norm = s => (s||'').replace(/\s+/g,' ').trim();
                const btn = Array.from(document.querySelectorAll('button, a, .close, .btn-close'))
                  .find(b => /^(ปิด|ยกเลิก|×|✕)$/.test(norm(b.textContent)) && b.offsetParent !== null);
                if (btn) btn.click();
            }""")
        except Exception:
            pass

    try:
        work_branch = (rec.get("workplace_branch") or "").strip()
        work_biz = (rec.get("work_biz") or "").strip()
        work_job = (rec.get("work_permit_job") or "").strip()
        work_detail = (rec.get("work_detail") or "").strip()

        log("      [Step 3.3] เลือกสถานที่ทำงาน...")

        # ---- เปิดโมดัล 'เลือกสถานที่ทำงาน' ----
        # ปุ่มจริง: <button data-action="addEmployerAddr" class="...lang_select_work_address">
        # ไม่มี onclick (ใช้ delegated handler) — real click เปิดได้ โมดัลโผล่ทันที (มี heading)
        # แต่รายการสาขา (option) โหลดผ่าน AJAX ช้ากว่า → ต้องแยกการตรวจ "เปิดแล้ว" ออกจาก "option พร้อม"
        # ⚠ ห้ามคลิกซ้ำหลังโมดัลเปิดแล้ว เพราะปุ่มเดียวกันอาจ toggle/รีเซ็ต AJAX
        # หมายเหตุ: option สาขาในเว็บมีหลายฟอร์แมต — บางบัญชีขึ้นต้น 'สำนักงาน...' บางบัญชีเป็นที่อยู่ตรงๆ
        # (เช่น '616/16 หมู่ที่ 1 แขวง/ตำบล แม่น้ำคู้...') → ตรวจด้วย keyword ที่อยู่ทั่วไปแทน
        def _branch_option_present() -> bool:
            return bool(page.evaluate(r"""() => {
                const KW = /(สำนักงาน|ตำบล|อำเภอ|จังหวัด|รหัสไปรษณีย์|หมู่ที่)/;
                return Array.from(document.querySelectorAll('select'))
                  .some(s => Array.from(s.options).some(o => o.value && KW.test(o.textContent || '')));
            }"""))

        def _workplace_modal_open() -> bool:
            # โมดัลเปิด = มี heading 'สถานที่ทำงาน/สาขา' หรือ 'ประเภทกิจการ' หรือมี option สาขาแล้ว
            return bool(page.evaluate(r"""() => {
                const t = document.body.innerText || '';
                if (/สถานที่ทำงาน\/สาขา|ประเภทกิจการ/.test(t)) return true;
                const KW = /(สำนักงาน|ตำบล|อำเภอ|จังหวัด|รหัสไปรษณีย์|หมู่ที่)/;
                return Array.from(document.querySelectorAll('select'))
                  .some(s => Array.from(s.options).some(o => o.value && KW.test(o.textContent || '')));
            }"""))

        def _click_workplace_btn() -> bool:
            for sel in (
                "button[data-action='addEmployerAddr']:visible",
                "button.lang_select_work_address:visible, a.lang_select_work_address:visible",
            ):
                try:
                    b = page.locator(sel).last
                    if b.count() > 0:
                        b.scroll_into_view_if_needed(timeout=2000)
                        b.click(timeout=4000)
                        return True
                except Exception:
                    continue
            try:
                b = page.locator(
                    "button:visible:has-text('เลือกสถานที่ทำงาน'), "
                    "a:visible:has-text('เลือกสถานที่ทำงาน')"
                ).last
                if b.count() > 0:
                    b.scroll_into_view_if_needed(timeout=2000)
                    b.click(timeout=4000)
                    return True
            except Exception:
                pass
            try:
                return bool(page.evaluate(r"""() => {
                    const b = document.querySelector("button[data-action='addEmployerAddr'], .lang_select_work_address");
                    if (!b) return false;
                    b.scrollIntoView({ block: 'center' });
                    b.click();
                    return true;
                }"""))
            except Exception:
                return False

        # ปิด swal/overlay ที่อาจค้างหลัง Step 3.1 (กันคลิกโดน overlay)
        try:
            page.evaluate(r"""() => {
                const btn = Array.from(document.querySelectorAll('.swal2-confirm, .swal2-close'))
                  .find(b => b.offsetParent !== null);
                if (btn) btn.click();
            }""")
        except Exception:
            pass
        page.wait_for_timeout(500)

        # คลิกเปิดโมดัล (คลิกซ้ำได้สูงสุด 3 ครั้ง เฉพาะเมื่อโมดัล "ยังไม่เปิด")
        modal_open = False
        for _attempt in range(3):
            if not _workplace_modal_open():
                _click_workplace_btn()
            # รอ heading โผล่ (โมดัลเปิด) สูงสุด ~6s
            for _i in range(6):
                page.wait_for_timeout(1000)
                if _workplace_modal_open():
                    modal_open = True
                    break
            if modal_open:
                break
            _wait_loading_disappeared(page, timeout_ms=3000, log=log, allow_modal=True)
        if not modal_open:
            try:
                _dbg = Path("reports") / "bt44_screenshots"
                _dbg.mkdir(parents=True, exist_ok=True)
                page.screenshot(path=str(_dbg / "step3_3_modal_fail.png"), full_page=True)
                _btns = page.evaluate(r"""() => Array.from(document.querySelectorAll('button,a'))
                    .filter(b => /เลือกสถานที่ทำงาน|addEmployerAddr/.test((b.textContent||'')+' '+(b.getAttribute('data-action')||'')))
                    .map(b => ({tag:b.tagName, vis:b.offsetParent!==null, da:b.getAttribute('data-action')||'', cls:(b.className||'').slice(0,80)}))""")
                log(f"      [Step 3.3] debug btns: {_btns}")
            except Exception:
                pass
            res["status"] = "ALERT"
            res["note"] = "Step 3.3: เปิดโมดัลเลือกสถานที่ทำงานไม่สำเร็จ (ไม่พบ heading/โมดัล)"
            return res

        # โมดัลเปิดแล้ว — รอรายการสาขา (option) โหลดผ่าน AJAX สูงสุด ~18s (ห้ามคลิกปุ่มซ้ำ)
        opened_wp = False
        for _i in range(18):
            if _branch_option_present():
                opened_wp = True
                break
            page.wait_for_timeout(1000)
        if not opened_wp:
            res["status"] = "ALERT"
            res["note"] = "Step 3.3: โมดัลเปิดแต่รายการสาขาไม่โหลด (timeout)"
            return res
        _wait_loading_disappeared(page, timeout_ms=6000, log=log, allow_modal=True)

        # ---- เลือกสาขา: select ที่ visible และมี option เป็นที่อยู่/สาขา ----
        # กลยุทธ์: คำนวณคะแนนความเหมือนทุก option แล้วเลือกตัวคะแนนสูงสุด
        #   - normalize ก่อน: ตัด prefix สัญลักษณ์ (เช่น "- :"), whitespace, dash, slash, comma, dot, วงเล็บ
        #   - exact (compact) match → ให้คะแนนเต็ม 1.0
        #   - substring (option ⊃ excel หรือ excel ⊃ option) → 0.95
        #   - Dice bigram similarity บนตัวอักษร (รองรับไทย+Eng+digit) → 0..1
        #   - +bonus 0.05 ถ้ารหัสไปรษณีย์ 5 หลักตรงกัน
        #   - threshold 0.30 → เลือกตัวคะแนนสูงสุด ถ้าต่ำกว่านี้ → fallback option แรก
        branch_res = page.evaluate(
            r"""(want) => {
            const norm = s => (s||'').replace(/\s+/g,' ').trim();
            // ตัด prefix อย่าง "- :", "-:", "- ", " : ", "•", bullet, dot leaders ออกหัวสตริง
            const stripPrefix = s => (s||'').replace(/^[\s\-:•·\.\u2013\u2014]+/,'').trim();
            const compact = s => stripPrefix(s||'').replace(/[\s\-\/.,()\\:]+/g,'').trim();
            const vis = el => { if(!el) return false; const st=getComputedStyle(el); const r=el.getBoundingClientRect(); return st.display!=='none'&&st.visibility!=='hidden'&&(el.offsetParent!==null||r.width>0); };
            const KW = /(สำนักงาน|ตำบล|อำเภอ|จังหวัด|รหัสไปรษณีย์|หมู่ที่)/;
            // เลือก select ที่ option อย่างน้อย 1 ตัวเป็นที่อยู่/สาขา
            const candidates = Array.from(document.querySelectorAll('select')).filter(s => {
                const opts = Array.from(s.options).filter(o => o.value);
                return opts.length > 0 && opts.some(o => KW.test(o.textContent || ''));
            });
            const sel = candidates.find(vis) || candidates[0];
            if (!sel) return { ok:false, reason:'no_select' };
            const opts = Array.from(sel.options).filter(o => o.value);
            if (opts.length === 0) return { ok:false, reason:'no_option' };

            // --- Dice bigram similarity (char-level, รองรับ unicode ไทย/Eng/digit) ---
            const bigrams = (s) => {
                const t = compact(s);
                const out = new Map();
                for (let i = 0; i < t.length - 1; i++) {
                    const bg = t.slice(i, i + 2);
                    out.set(bg, (out.get(bg) || 0) + 1);
                }
                return out;
            };
            const dice = (a, b) => {
                const A = bigrams(a), B = bigrams(b);
                if (A.size === 0 || B.size === 0) return 0;
                let inter = 0, totalA = 0, totalB = 0;
                for (const [, v] of A) totalA += v;
                for (const [, v] of B) totalB += v;
                for (const [bg, va] of A) {
                    const vb = B.get(bg);
                    if (vb) inter += Math.min(va, vb);
                }
                return (2 * inter) / (totalA + totalB);
            };

            const w = norm(want);
            const wc = compact(want);
            const wantZip = (want.match(/\b(\d{5})\b/) || [])[1] || '';
            const debug = [];
            let best = null, bestScore = -1, bestMethod = '';

            if (w && wc) {
                for (const o of opts) {
                    const oText = o.textContent || '';
                    const oc = compact(oText);
                    if (!oc) continue;
                    let score = 0, method = '';
                    if (oc === wc) {
                        score = 1.0; method = 'exact';
                    } else if (oc.includes(wc) || (wc.length > 6 && wc.includes(oc))) {
                        // ratio ปรับตามความยาว: ยิ่งความยาวใกล้กัน ยิ่งสูง
                        const ratio = Math.min(oc.length, wc.length) / Math.max(oc.length, wc.length);
                        score = 0.85 + 0.10 * ratio; method = 'substring';
                    } else {
                        score = dice(want, oText); method = 'dice';
                    }
                    // bonus: รหัสไปรษณีย์ 5 หลักตรง
                    const oZip = (oText.match(/\b(\d{5})\b/) || [])[1] || '';
                    if (wantZip && oZip && wantZip === oZip) score = Math.min(1.0, score + 0.05);
                    debug.push({ text: norm(oText).slice(0,80), score: +score.toFixed(3), method });
                    if (score > bestScore) { bestScore = score; best = o; bestMethod = method; }
                }
            }
            // threshold ขั้นต่ำ → ถ้าต่ำกว่านี้แสดงว่าไม่มีตัวที่ใกล้เคียงเลย ให้ fallback option แรก
            const THRESHOLD = 0.30;
            if (!best || bestScore < THRESHOLD) {
                best = opts[0];
                bestMethod = 'fallback_first';
                bestScore = 0;
            }

            sel.value = best.value;
            sel.dispatchEvent(new Event('input', { bubbles: true }));
            sel.dispatchEvent(new Event('change', { bubbles: true }));
            if (window.jQuery) { try { jQuery(sel).val(best.value).trigger('change'); jQuery(sel).trigger({type:'select2:select', params:{data:{id:best.value, text:best.textContent}}}); } catch(e){} }

            // เก็บ top-3 debug รายการที่เปรียบเทียบ
            debug.sort((a,b) => b.score - a.score);
            return {
                ok: true,
                text: norm(best.textContent).slice(0,120),
                matched_by: bestMethod,
                score: +bestScore.toFixed(3),
                option_count: opts.length,
                top3: debug.slice(0, 3),
            };
        }""",
            work_branch,
        )
        if not branch_res or not branch_res.get("ok"):
            _close_workplace_modal()
            res["status"] = "ALERT"
            res["note"] = f"Step 3.3: ไม่พบสาขาให้เลือก ({branch_res.get('reason') if branch_res else 'no_response'})"
            return res
        page.wait_for_timeout(1000)
        _wait_loading_disappeared(page, timeout_ms=6000, log=log, allow_modal=True)
        log(
            f"      [Step 3.3] เลือกสาขา: {branch_res.get('text','')} "
            f"[{branch_res.get('matched_by','')} score={branch_res.get('score','?')}, "
            f"options={branch_res.get('option_count','?')}]"
        )
        # Log top-3 candidates เมื่อไม่ใช่ exact match (ช่วย debug)
        if branch_res.get("matched_by") not in ("exact", "fallback_first"):
            for cand in branch_res.get("top3", []):
                log(f"        - cand: score={cand.get('score')} [{cand.get('method')}] {cand.get('text','')}")

        # ---- 3.3.1: เลือกประเภทกิจการ (businessType) ตาม Excel ----
        # ใช้ similarity scoring เดียวกับสาขา → ทนต่อความต่างของ whitespace/อักขระ/prefix
        # หา select ของประเภทกิจการ:
        #   1) id/name = businessType หรือมีคำว่า biz/business
        #   2) select ที่ visible และมี option หลายตัวเริ่มต้นด้วย 'การผลิต|การให้บริการ|การค้า|กิจการ|การเกษตร'
        if work_biz:
            # ⚠ option ของ businessType โหลดผ่าน AJAX หลังเลือกสาขา → ต้องรอจนมี option ≥ 2 ตัว
            # (ก่อนหน้านี้ JS รันเร็วเกินไป → ได้ select เปล่า/ตัวเลือกเดียว → fallback ผิด)
            log("      [Step 3.3.1] รอ option ประเภทกิจการโหลดจาก AJAX...")
            biz_ready = False
            for _i in range(30):  # สูงสุด ~15 วินาที
                try:
                    ready = page.evaluate(r"""() => {
                        const norm = s => (s||'').replace(/\s+/g,' ').trim();
                        const vis = el => { if(!el) return false; const st=getComputedStyle(el); const r=el.getBoundingClientRect(); return st.display!=='none'&&st.visibility!=='hidden'&&(el.offsetParent!==null||r.width>0); };
                        const BIZ_KW = /^(การผลิต|การให้บริการ|การค้า|กิจการ|การเกษตร|การก่อสร้าง|การประมง|การขนส่ง|การทำเหมือง|งานบ้าน)/;
                        let sels = Array.from(document.querySelectorAll('#businessType, select[name="businessType"], select[id*="biz" i], select[name*="biz" i], select[id*="business" i], select[name*="business" i]')).filter(vis);
                        if (sels.length === 0) {
                            sels = Array.from(document.querySelectorAll('select')).filter(s => {
                                if (!vis(s)) return false;
                                const opts = Array.from(s.options).filter(o => o.value);
                                if (opts.length < 2) return false;
                                const matches = opts.filter(o => BIZ_KW.test(norm(o.textContent))).length;
                                return matches >= 2;
                            });
                        }
                        const sel = sels[0];
                        if (!sel) return { ok:false, count:0 };
                        const opts = Array.from(sel.options).filter(o => o.value);
                        return { ok: opts.length >= 2, count: opts.length };
                    }""")
                    if ready and ready.get("ok"):
                        biz_ready = True
                        break
                except Exception as e:
                    # navigation/context destroyed — รอ DOM พร้อมแล้วค่อยลองใหม่
                    if "context was destroyed" in str(e).lower() or "navigation" in str(e).lower():
                        try:
                            page.wait_for_load_state("domcontentloaded", timeout=3000)
                        except Exception:
                            pass
                page.wait_for_timeout(500)
            if not biz_ready:
                log(f"      [Step 3.3.1] ⚠ option ประเภทกิจการอาจยังโหลดไม่ครบ — จะลองเลือกต่อ")
            _wait_loading_disappeared(page, timeout_ms=4000, log=log, allow_modal=True)

            # ลอง evaluate ตัวเลือกหลัก — retry 3 ครั้งหากเจอ navigation/context destroyed
            biz_res = None
            for _retry in range(3):
                try:
                    biz_res = page.evaluate(
                r"""(want) => {
                const norm = s => (s||'').replace(/\s+/g,' ').trim();
                const stripPrefix = s => (s||'').replace(/^[\s\-:•·\.\u2013\u2014]+/,'').trim();
                const compact = s => stripPrefix(s||'').replace(/[\s\-\/.,()\\:]+/g,'').trim();
                const vis = el => { if(!el) return false; const st=getComputedStyle(el); const r=el.getBoundingClientRect(); return st.display!=='none'&&st.visibility!=='hidden'&&(el.offsetParent!==null||r.width>0); };
                const BIZ_KW = /^(การผลิต|การให้บริการ|การค้า|กิจการ|การเกษตร|การก่อสร้าง|การประมง|การขนส่ง|การทำเหมือง|งานบ้าน)/;
                // 1) id/name explicit
                let sels = Array.from(document.querySelectorAll('#businessType, select[name="businessType"], select[id*="biz" i], select[name*="biz" i], select[id*="business" i], select[name*="business" i]')).filter(vis);
                // 2) heuristic: visible select with multiple options matching BIZ_KW
                if (sels.length === 0) {
                    sels = Array.from(document.querySelectorAll('select')).filter(s => {
                        if (!vis(s)) return false;
                        const opts = Array.from(s.options).filter(o => o.value);
                        if (opts.length < 2) return false;
                        const matches = opts.filter(o => BIZ_KW.test(norm(o.textContent))).length;
                        return matches >= 2;
                    });
                }
                const sel = sels[0];
                if (!sel) return { ok:false, reason:'no_businessType_select' };
                const opts = Array.from(sel.options).filter(o => o.value);
                if (opts.length === 0) return { ok:false, reason:'no_option' };

                // Dice bigram similarity (char-level)
                const bigrams = (s) => {
                    const t = compact(s);
                    const out = new Map();
                    for (let i = 0; i < t.length - 1; i++) {
                        const bg = t.slice(i, i + 2);
                        out.set(bg, (out.get(bg) || 0) + 1);
                    }
                    return out;
                };
                const dice = (a, b) => {
                    const A = bigrams(a), B = bigrams(b);
                    if (A.size === 0 || B.size === 0) return 0;
                    let inter = 0, totalA = 0, totalB = 0;
                    for (const [, v] of A) totalA += v;
                    for (const [, v] of B) totalB += v;
                    for (const [bg, va] of A) {
                        const vb = B.get(bg);
                        if (vb) inter += Math.min(va, vb);
                    }
                    return (2 * inter) / (totalA + totalB);
                };

                const w = norm(want), wc = compact(want);
                const debug = [];
                let best = null, bestScore = -1, bestMethod = '';
                for (const o of opts) {
                    const oText = o.textContent || '';
                    const oc = compact(oText);
                    if (!oc) continue;
                    let score = 0, method = '';
                    if (oc === wc) {
                        score = 1.0; method = 'exact';
                    } else if (oc.includes(wc) || (wc.length > 6 && wc.includes(oc))) {
                        const ratio = Math.min(oc.length, wc.length) / Math.max(oc.length, wc.length);
                        score = 0.85 + 0.10 * ratio; method = 'substring';
                    } else {
                        score = dice(want, oText); method = 'dice';
                    }
                    debug.push({ text: norm(oText).slice(0,80), score: +score.toFixed(3), method });
                    if (score > bestScore) { bestScore = score; best = o; bestMethod = method; }
                }
                const THRESHOLD = 0.30;
                if (!best || bestScore < THRESHOLD) {
                    return { ok:false, reason:'low_score', best_score: +bestScore.toFixed(3),
                             top3: debug.sort((a,b)=>b.score-a.score).slice(0,3) };
                }
                sel.value = best.value;
                sel.dispatchEvent(new Event('input', { bubbles: true }));
                sel.dispatchEvent(new Event('change', { bubbles: true }));
                if (window.jQuery) { try { jQuery(sel).val(best.value).trigger('change'); jQuery(sel).trigger({type:'select2:select', params:{data:{id:best.value, text:best.textContent}}}); } catch(e){} }
                debug.sort((a,b)=>b.score-a.score);
                return {
                    ok: true,
                    text: norm(best.textContent).slice(0,80),
                    matched_by: bestMethod,
                    score: +bestScore.toFixed(3),
                    option_count: opts.length,
                    top3: debug.slice(0,3),
                };
            }""",
                        work_biz,
                    )
                    break  # สำเร็จ — ออกจาก retry loop
                except Exception as e:
                    msg = str(e)
                    if "context was destroyed" in msg.lower() or "navigation" in msg.lower():
                        log(f"      [Step 3.3.1] ⚠ context destroyed during evaluate — retry {_retry+1}/3")
                        try:
                            page.wait_for_load_state("domcontentloaded", timeout=5000)
                        except Exception:
                            pass
                        page.wait_for_timeout(800)
                        continue
                    raise  # error อื่น → re-raise
            if not biz_res or not biz_res.get("ok"):
                # Log top-3 ก่อน fail เพื่อช่วย debug
                if biz_res and biz_res.get("top3"):
                    log(f"      [Step 3.3.1] ✗ ไม่พบประเภทกิจการที่ใกล้ '{work_biz}' (best={biz_res.get('best_score')}):")
                    for cand in biz_res.get("top3", []):
                        log(f"        - cand: score={cand.get('score')} [{cand.get('method')}] {cand.get('text','')}")
                _close_workplace_modal()
                res["status"] = "ALERT"
                res["note"] = f"Step 3.3.1: เลือกประเภทกิจการ '{work_biz}' ไม่ได้ ({biz_res.get('reason') if biz_res else 'no_response'})"
                return res
            log(
                f"      [Step 3.3.1] ประเภทกิจการ: {biz_res.get('text','')} "
                f"[{biz_res.get('matched_by','')} score={biz_res.get('score','?')}, "
                f"options={biz_res.get('option_count','?')}]"
            )
            if biz_res.get("matched_by") not in ("exact",):
                for cand in biz_res.get("top3", []):
                    log(f"        - cand: score={cand.get('score')} [{cand.get('method')}] {cand.get('text','')}")
            page.wait_for_timeout(900)
            _wait_loading_disappeared(page, timeout_ms=6000, log=log, allow_modal=True)

        # ---- 3.3.1: ตรวจสอบประเภทงานที่ขออนุญาต (permitCateWork) ต้องตรงกับ Excel ----
        if work_job:
            verify = page.evaluate(r"""(want) => {
                const norm = s => (s||'').replace(/\s+/g,' ').trim();
                const vis = el => { if(!el) return false; const st=getComputedStyle(el); const r=el.getBoundingClientRect(); return st.display!=='none'&&st.visibility!=='hidden'&&(el.offsetParent!==null||r.width>0); };
                let sels = Array.from(document.querySelectorAll('#permitCateWork, select[name="permitCateWork"]')).filter(vis);
                let sel = sels[sels.length - 1];
                if (!sel) sel = Array.from(document.querySelectorAll('select')).filter(vis)
                    .find(s => Array.from(s.options).some(o => /กรรมกร|งาน/.test(o.textContent || '')));
                if (!sel) return { ok:false, reason:'no permitCateWork select' };
                const w = norm(want);
                const opt = Array.from(sel.options).find(o => norm(o.textContent) === w)
                         || Array.from(sel.options).find(o => w && norm(o.textContent).includes(w));
                const available = Array.from(sel.options).map(o => norm(o.textContent)).filter(t => t && t!=='-- กรุณาเลือก --' && t!=='กรุณาเลือก');
                if (!opt) return { ok:false, reason:'mismatch', options: available.slice(0,10) };
                sel.value = opt.value;
                sel.dispatchEvent(new Event('input', { bubbles: true }));
                sel.dispatchEvent(new Event('change', { bubbles: true }));
                if (window.jQuery) { try { jQuery(sel).val(opt.value).trigger('change'); jQuery(sel).trigger({type:'select2:select', params:{data:{id:opt.value, text:opt.textContent}}}); } catch(e){} }
                return { ok:true, text: norm(opt.textContent).slice(0,60) };
            }""", work_job)
            if not verify or not verify.get("ok"):
                opts = (verify or {}).get("options", [])
                _close_workplace_modal()
                res["status"] = "SKIP"
                res["note"] = (f"Step 3.3.1: ประเภทงานที่ขออนุญาตไม่ตรง Excel "
                               f"(ต้องการ '{work_job}', มีให้เลือก: {opts})")
                return res
            log(f"      [Step 3.3.1] ✓ ประเภทงานที่ขออนุญาตตรง: {verify.get('text','')}")
            page.wait_for_timeout(700)
            _wait_loading_disappeared(page, timeout_ms=5000, log=log, allow_modal=True)

        # ---- 3.3.2: กรอกลักษณะงาน (jobDescription) ----
        # NOTE: id="jobDescription" มีซ้ำหลายตัวใน DOM (template ที่ซ่อน + ตัวจริง)
        # ต้องเล็งเฉพาะตัวที่ ':visible' มิฉะนั้น .first จะไปโดน template แล้ว wait_for(visible) timeout
        if work_detail:
            jd_filled = False
            for _ in range(3):
                try:
                    jd = page.locator('#jobDescription:visible, input[name="jobDescription"]:visible').last
                    jd.wait_for(state="visible", timeout=4000)
                    jd.scroll_into_view_if_needed(timeout=2000)
                    jd.click(timeout=3000)
                    jd.fill("")
                    jd.fill(work_detail)
                    jd.dispatch_event("input")
                    jd.dispatch_event("change")
                    val = jd.input_value(timeout=2000)
                    if (val or "").strip() == work_detail:
                        jd_filled = True
                        break
                except Exception:
                    pass
                page.wait_for_timeout(600)
            if not jd_filled:
                r = page.evaluate(r"""(val) => {
                    const vis = el => { if(!el) return false; const st=getComputedStyle(el); const rc=el.getBoundingClientRect(); return st.display!=='none'&&st.visibility!=='hidden'&&rc.width>0&&rc.height>0; };
                    const els = Array.from(document.querySelectorAll('#jobDescription, input[name="jobDescription"]')).filter(vis);
                    const el = els[els.length - 1];
                    if (!el) return { ok:false };
                    el.focus(); el.value = val;
                    el.dispatchEvent(new Event('input', { bubbles:true }));
                    el.dispatchEvent(new Event('change', { bubbles:true }));
                    el.dispatchEvent(new Event('blur', { bubbles:true }));
                    return { ok:true };
                }""", work_detail)
                jd_filled = bool(r and r.get("ok"))
            if not jd_filled:
                _close_workplace_modal()
                res["status"] = "ALERT"
                res["note"] = "Step 3.3.2: กรอกลักษณะงานไม่สำเร็จ"
                return res
            log(f"      [Step 3.3.2] ลักษณะงาน: {work_detail}")

        # ---- บันทึกในโมดัล workplace ----
        msave = False
        try:
            sb = page.locator("button:visible:has-text('บันทึก'), a:visible:has-text('บันทึก')").last
            if sb.count() > 0:
                sb.scroll_into_view_if_needed(timeout=2000)
                sb.click(timeout=4000, force=True)
                msave = True
        except Exception:
            msave = False
        if not msave:
            try:
                msave = bool(page.evaluate(r"""() => {
                    const btn = Array.from(document.querySelectorAll('button, a')).reverse().find(el => {
                        const st = getComputedStyle(el);
                        return /บันทึก/.test((el.textContent || '').trim()) && el.offsetParent !== null && st.visibility !== 'hidden';
                    });
                    if (!btn) return false;
                    btn.scrollIntoView({ block: 'center' }); btn.click(); return true;
                }"""))
            except Exception:
                msave = False
        if not msave:
            res["status"] = "ALERT"
            res["note"] = "Step 3.3: ไม่พบปุ่มบันทึกในโมดัลสถานที่ทำงาน"
            return res
        page.wait_for_timeout(1200)
        _wait_loading_disappeared(page, timeout_ms=8000, log=log, allow_modal=True)
        # ปิด swal ยืนยัน/สำเร็จ ถ้ามี
        try:
            page.evaluate(r"""() => {
                const btn = Array.from(document.querySelectorAll('.swal2-confirm, button, a'))
                  .find(b => /ตกลง|ยืนยัน|ปิด|ok/i.test((b.textContent||'').trim()) && b.offsetParent !== null);
                if (btn) btn.click();
            }""")
        except Exception:
            pass
        page.wait_for_timeout(800)
        _wait_loading_disappeared(page, timeout_ms=5000, log=log, allow_modal=True)
        log("      [Step 3.3] ✅ บันทึกข้อมูลการขออนุญาตแล้ว")

        # ---- Step 3.4: กดถัดไป → Step 4 (แนบเอกสาร) ----
        next_ok = False
        try:
            nb = page.locator(
                "button:visible.btn-next:has-text('ถัดไป'), "
                "button:visible:has-text('ถัดไป'), a:visible:has-text('ถัดไป')"
            ).last
            if nb.count() > 0:
                nb.scroll_into_view_if_needed(timeout=2000)
                nb.click(timeout=4000)
                next_ok = True
        except Exception:
            next_ok = False
        if not next_ok:
            try:
                next_ok = bool(page.evaluate(r"""() => {
                    const btn = Array.from(document.querySelectorAll('button, a')).find(el =>
                        /ถัดไป/.test((el.textContent || '').trim()) && el.offsetParent !== null);
                    if (!btn) return false;
                    btn.scrollIntoView({ block: 'center' }); btn.click(); return true;
                }"""))
            except Exception:
                next_ok = False
        if not next_ok:
            res["status"] = "ALERT"
            res["note"] = "Step 3.4: ไม่พบปุ่มถัดไป"
            return res
        page.wait_for_timeout(1200)
        _wait_loading_disappeared(page, timeout_ms=10000, log=log)
        log("      [Step 3.4] ✅ กดถัดไป → Step 4")

        res["status"] = "SUCCESS"
        res["note"] = "Step 3.3-3.4 complete (สาขา+ประเภทกิจการ+ตรวจประเภทงาน+ลักษณะงาน+ถัดไป)"
        return res
    except Exception as e:
        res["status"] = "ERROR"
        res["note"] = f"Step 3.3 error: {str(e)[:180]}"
        return res


# ════════════════════════════════════════════════════════════════════════════
#  Step 4 — แนบเอกสาร (4.1 เอกสารลูกจ้าง + 4.2 เอกสารนายจ้าง)
# ════════════════════════════════════════════════════════════════════════════
_BT44_MAX_DOC_MB: float = 4.0


def _bt44_validate_row_docs(rec: dict[str, Any], limit_mb: float = _BT44_MAX_DOC_MB) -> list[str]:
    """ตรวจไฟล์แนบของ 1 แถว 'ก่อนรัน' ตามกติกา:
    - ไฟล์ห้ามเกิน 4 MB (เกิน → บล็อก ไม่รัน)
    - ช่องที่มี * (required) ต้องมีไฟล์ใน Excel
    - กลุ่มหนังสือเดินทาง (สำเนาหนังสือเดินทาง/ใช้แทน/หลักฐานเข้าราชอาณาจักร) ต้องมีอย่างน้อย 1 ไฟล์
    - ไฟล์ที่ระบุ path ต้องมีอยู่จริงบนดิสก์
    คืน list ข้อความปัญหา (ว่าง = ผ่าน)
    """
    limit = int(limit_mb * 1024 * 1024)
    problems: list[str] = []
    group_seen: dict[str, bool] = {}
    group_has: dict[str, bool] = {}
    for d in rec.get("docs", []):
        th = (d.get("th_name") or "")[:30]
        raw = (d.get("raw_path") or "").strip()
        path = (d.get("path") or "").strip()
        grp = d.get("group") or ""
        if grp:
            group_seen[grp] = True
            if raw:
                group_has[grp] = True
        if not raw:
            if d.get("required"):
                problems.append(f"ขาดไฟล์บังคับ: {th}")
            continue
        p = Path(path)
        if not p.is_file():
            problems.append(f"ไม่พบไฟล์: {p.name}")
            continue
        try:
            if p.stat().st_size > limit:
                problems.append(f"ไฟล์เกิน {limit_mb:.0f}MB ({p.stat().st_size/1024/1024:.2f}MB): {p.name}")
        except OSError:
            pass
    for grp, seen in group_seen.items():
        if seen and not group_has.get(grp):
            problems.append("ต้องแนบอย่างน้อย 1 ไฟล์ (กลุ่มหนังสือเดินทาง)")
    for o in rec.get("other_docs", []):
        path = (o.get("path") or "").strip()
        if not path:
            continue
        p = Path(path)
        if not p.is_file():
            problems.append(f"ไม่พบไฟล์(อื่นๆ): {p.name}")
            continue
        try:
            if p.stat().st_size > limit:
                problems.append(f"ไฟล์เกิน {limit_mb:.0f}MB (อื่นๆ): {p.name}")
        except OSError:
            pass
    return problems


def _bt44_preflight_docs(records: list[dict[str, Any]], log=print) -> dict[int, list[str]]:
    """สแกนไฟล์แนบทุกแถว 'ก่อน' เริ่มทำงาน — รายงานแถวที่ไม่ผ่าน (จะถูกข้าม ไม่รัน)
    คืน dict row_index → list ปัญหา
    """
    blocked: dict[int, list[str]] = {}
    for rec in records:
        probs = _bt44_validate_row_docs(rec)
        if probs:
            blocked[rec.get("row_index", -1)] = probs
    if blocked:
        log(f"  ⚠ ตรวจไฟล์แนบ Step 4: พบ {len(blocked)} แถวมีปัญหา (จะถูกข้าม ไม่รัน):")
        for rec in records:
            ri = rec.get("row_index", -1)
            if ri in blocked:
                log(f"      • แถว {ri} {rec.get('name','')}: " + " | ".join(blocked[ri][:6]))
    else:
        log("  ✓ ตรวจไฟล์แนบ Step 4: ทุกแถวผ่าน (ไฟล์ ≤ 4 MB, ช่องบังคับมีไฟล์ครบ)")
    return blocked


def _bt44_find_doc_input_id(page: Page, th_name: str) -> str:
    """หา id ของช่อง input[type=file] โดยจับคู่จาก data-document-th (ค่า TH) = ชื่อเอกสาร
    (id สุ่มทุก session จึง match จากชื่อเอกสารแทน) — ถ้าไม่มี id จะตั้งให้
    """
    try:
        return page.evaluate(r"""(want) => {
            const norm = s => (s||'').replace(/\s+/g,' ').trim();
            const w = norm(want);
            const inps = Array.from(document.querySelectorAll('input[type=file]'));
            const getTH = inp => {
                const raw = inp.getAttribute('data-document-th') || '';
                try { const j = JSON.parse(raw); return norm(j.TH || ''); } catch (e) { return norm(raw); }
            };
            const ensureId = inp => { if (!inp.id) inp.id = 'bt44doc-' + Math.random().toString(36).slice(2, 8); return inp.id; };
            for (const inp of inps) { if (getTH(inp) === w) return ensureId(inp); }
            for (const inp of inps) { const th = getTH(inp); if (th && (th.includes(w) || w.includes(th))) return ensureId(inp); }
            return '';
        }""", th_name)
    except Exception:
        return ""


def _bt44_handle_crop_modal(page: Page, log=print) -> bool:
    """รูปถ่าย 3x4 เมื่อแนบไฟล์ภาพจะเด้งโมดอลครอป (cropper.js) — ปุ่มบันทึก = #crop-save-button
    ตั้งกรอบครอป 3:4 ใหญ่สุด (best-effort ถ้าเข้าถึง window.cropper ได้) แล้วกดบันทึก
    คืน True ถ้าพบและจัดการโมดอลครอปแล้ว
    """
    appeared = False
    for _ in range(16):  # รอสูงสุด ~8 วินาที
        try:
            shown = bool(page.evaluate(r"""() => {
                const b = document.getElementById('crop-save-button');
                if (!b) return false;
                const r = b.getBoundingClientRect(); const st = getComputedStyle(b);
                return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 0 && r.height > 0;
            }"""))
        except Exception:
            shown = False
        if shown:
            appeared = True
            break
        page.wait_for_timeout(500)
    if not appeared:
        return False
    page.wait_for_timeout(700)
    try:
        page.evaluate(r"""() => {
            try {
                const cr = window.cropper;
                if (cr && typeof cr.getCanvasData === 'function' && typeof cr.setCropBoxData === 'function') {
                    const c = cr.getCanvasData(); const ar = 3 / 4;
                    let w = c.width, h = w / ar;
                    if (h > c.height) { h = c.height; w = h * ar; }
                    cr.setCropBoxData({ left: c.left + (c.width - w) / 2, top: c.top + (c.height - h) / 2, width: w, height: h });
                }
            } catch (e) {}
        }""")
    except Exception:
        pass
    page.wait_for_timeout(300)
    try:
        page.evaluate(r"""() => { const b = document.getElementById('crop-save-button'); if (b) b.click(); }""")
    except Exception:
        pass
    for _ in range(20):  # รอโมดอลปิดสูงสุด ~10 วินาที
        page.wait_for_timeout(500)
        try:
            gone = bool(page.evaluate(r"""() => {
                const b = document.getElementById('crop-save-button');
                if (!b) return true;
                const r = b.getBoundingClientRect(); const st = getComputedStyle(b);
                return !(st.display !== 'none' && r.width > 0 && r.height > 0);
            }"""))
        except Exception:
            gone = False
        if gone:
            break
    page.wait_for_timeout(500)
    log("        ✓ ครอปรูปถ่าย (3:4) แล้วบันทึก")
    return True


def _bt44_upload_one_doc(page: Page, input_id: str, abs_path: str, is_photo: bool, log=print) -> dict[str, Any]:
    """แนบไฟล์ลงช่องเอกสาร 1 ช่อง (set_input_files บน input ที่ซ่อนได้)
    ช่องรูปถ่ายจะเด้งโมดอลครอป → จัดการผ่าน _bt44_handle_crop_modal
    ยืนยันสำเร็จเมื่อ input มีไฟล์ หรือมี chip ชื่อไฟล์ปรากฏ
    """
    res = {"ok": False, "note": ""}
    p = Path(abs_path)
    if not abs_path or not p.exists():
        res["note"] = f"ไม่พบไฟล์ ({p.name if abs_path else 'ว่าง'})"
        return res
    try:
        if p.stat().st_size > _BT44_MAX_DOC_MB * 1024 * 1024:
            res["note"] = f"ไฟล์เกิน {_BT44_MAX_DOC_MB:.0f}MB ({p.stat().st_size/1024/1024:.2f}MB)"
            return res
    except OSError:
        pass
    try:
        page.set_input_files(f"#{input_id}", str(p))
    except Exception as e:
        res["note"] = f"แนบไฟล์ล้มเหลว: {str(e)[:90]}"
        return res
    if is_photo:
        _bt44_handle_crop_modal(page, log=log)
    fn = p.name
    confirmed = False
    for _ in range(16):  # รอยืนยันสูงสุด ~8 วินาที
        try:
            confirmed = bool(page.evaluate(r"""(args) => {
                const norm = s => (s||'').replace(/\s+/g,' ').trim();
                const id = args.id; const fn = norm(args.fn); const photo = args.photo;
                const inp = document.getElementById(id);
                if (!photo && inp && inp.files && inp.files.length > 0) return true;
                const fnShort = fn.length > 12 ? fn.slice(0, 12) : fn;
                return Array.from(document.querySelectorAll('a,span,div,p,li,td')).some(e => {
                    if (!e.offsetParent) return false;
                    const t = norm(e.textContent);
                    return t && (t.includes(fn) || (fnShort.length >= 6 && t.includes(fnShort)));
                });
            }""", {"id": input_id, "fn": fn, "photo": is_photo}))
        except Exception:
            confirmed = False
        if confirmed:
            break
        page.wait_for_timeout(500)
    res["ok"] = confirmed
    res["note"] = fn if confirmed else f"อัปโหลดไม่ยืนยัน ({fn})"
    return res


def _bt44_attach_one_other(page: Page, abs_path: str, idx: int, log=print) -> bool:
    """เอกสารอื่นๆที่เกี่ยวข้อง — แนบทีละไฟล์ผ่านโมดอล 'อัปโหลดเอกสาร' (เปิดด้วยปุ่ม
    [data-action='popup-upload'] / 'เพิ่มเอกสาร'): เลือกไฟล์ → กรอกชื่อเอกสาร (#document) → ยืนยัน
    """
    p = Path(abs_path)
    if not abs_path or not p.exists():
        log(f"        ✗ เอกสารอื่นๆ #{idx}: ไม่พบไฟล์ ({p.name if abs_path else 'ว่าง'})")
        return False
    try:
        if p.stat().st_size > _BT44_MAX_DOC_MB * 1024 * 1024:
            log(f"        ✗ เอกสารอื่นๆ #{idx}: เกิน {_BT44_MAX_DOC_MB:.0f}MB ({p.stat().st_size/1024/1024:.2f}MB)")
            return False
    except OSError:
        pass
    try:
        # เปิดโมดอล
        opened = False
        try:
            ab = page.locator("[data-action='popup-upload']:visible, button:visible:has-text('เพิ่มเอกสาร')").last
            if ab.count() > 0:
                ab.scroll_into_view_if_needed(timeout=2000)
                ab.click(timeout=4000)
                opened = True
        except Exception:
            opened = False
        if not opened:
            opened = bool(page.evaluate(r"""() => {
                const b = document.querySelector("[data-action='popup-upload']")
                  || Array.from(document.querySelectorAll('button,a')).find(e => /เพิ่มเอกสาร/.test((e.textContent||'').trim()));
                if (b) { b.click(); return true; }
                return false;
            }"""))
        if not opened:
            log(f"        ✗ เอกสารอื่นๆ #{idx}: ไม่พบปุ่ม 'เพิ่มเอกสาร'")
            return False
        try:
            page.wait_for_selector("#document", state="visible", timeout=8000)
        except PWTimeoutError:
            log(f"        ✗ เอกสารอื่นๆ #{idx}: โมดอลไม่เปิด (#document ไม่พบ)")
            return False
        page.wait_for_timeout(500)
        # ทำเครื่องหมายโมดอลที่มี #document + ปุ่มยืนยัน
        page.evaluate(r"""() => {
            const norm = s => (s||'').replace(/\s+/g,' ').trim();
            const doc = document.getElementById('document');
            if (!doc) return false;
            let root = doc;
            for (let up = 0; up < 10 && root.parentElement; up++) {
                root = root.parentElement;
                const hasConfirm = Array.from(root.querySelectorAll('button,a')).some(e => norm(e.textContent) === 'ยืนยัน');
                if (hasConfirm) {
                    document.querySelectorAll('[data-bt44-modal]').forEach(e => e.removeAttribute('data-bt44-modal'));
                    root.setAttribute('data-bt44-modal', '1');
                    return true;
                }
            }
            return false;
        }""")
        # เลือกไฟล์: ถ้ามี input ซ่อนในโมดอล set ตรง ไม่งั้นใช้ file chooser
        modal_has_input = bool(page.evaluate(r"""() => {
            const root = document.querySelector("[data-bt44-modal='1']") || document;
            const inp = root.querySelector("input[type=file]");
            if (inp) { inp.setAttribute('data-bt44-file', '1'); return true; }
            return false;
        }"""))
        if modal_has_input:
            page.set_input_files("[data-bt44-file='1']", str(p))
        else:
            choose = page.locator("[data-bt44-modal='1']").locator("button, a, label").filter(has_text="เลือกไฟล์").first
            if choose.count() == 0:
                choose = page.locator("button:visible, a:visible, label:visible").filter(has_text="เลือกไฟล์").last
            with page.expect_file_chooser(timeout=8000) as fc:
                choose.click(timeout=4000)
            fc.value.set_files(str(p))
        page.wait_for_timeout(900)
        # กรอกชื่อเอกสาร = ชื่อไฟล์ (ไม่รวมนามสกุล)
        try:
            page.locator("#document").fill(p.stem, timeout=3000)
        except Exception:
            page.evaluate(r"""(v) => {
                const t = document.getElementById('document');
                if (t) { t.value = v; t.dispatchEvent(new Event('input', { bubbles: true })); t.dispatchEvent(new Event('change', { bubbles: true })); }
            }""", p.stem)
        page.wait_for_timeout(300)
        # ยืนยัน
        try:
            cb = page.locator("[data-bt44-modal='1']").locator("button:has-text('ยืนยัน'), a:has-text('ยืนยัน')").first
            if cb.count() == 0:
                cb = page.locator("button:visible:has-text('ยืนยัน'), a:visible:has-text('ยืนยัน')").last
            cb.click(timeout=4000)
        except Exception:
            page.evaluate(r"""() => {
                const root = document.querySelector("[data-bt44-modal='1']") || document;
                const b = Array.from(root.querySelectorAll('button,a')).find(e => /ยืนยัน/.test((e.textContent||'').trim()));
                if (b) b.click();
            }""")
        # รอโมดอลปิด
        closed = False
        for _ in range(16):
            page.wait_for_timeout(500)
            try:
                closed = bool(page.evaluate(r"""() => { const d = document.getElementById('document'); return !d || !d.offsetParent; }"""))
            except Exception:
                closed = False
            if closed:
                break
        page.wait_for_timeout(400)
        if closed:
            log(f"        ✓ เอกสารอื่นๆ #{idx}: {p.name}")
            return True
        log(f"        ⚠ เอกสารอื่นๆ #{idx}: ไม่ยืนยันการปิดโมดอล ({p.name})")
        try:
            page.evaluate(r"""() => {
                const b = Array.from(document.querySelectorAll('button,a')).find(e => /ยกเลิก/.test((e.textContent||'').trim()) && e.offsetParent);
                if (b) b.click();
            }""")
        except Exception:
            pass
        return False
    except Exception as e:
        log(f"        ✗ เอกสารอื่นๆ #{idx} ผิดพลาด: {str(e)[:90]}")
        try:
            page.evaluate(r"""() => {
                const b = Array.from(document.querySelectorAll('button,a')).find(e => /ยกเลิก/.test((e.textContent||'').trim()) && e.offsetParent);
                if (b) b.click();
            }""")
        except Exception:
            pass
        return False


def _bt44_click_next(page: Page, log=print) -> bool:
    """กดปุ่ม 'ถัดไป' (#button_next) ในขั้นแนบเอกสาร"""
    try:
        nb = page.locator("#button_next:visible").first
        if nb.count() > 0:
            nb.scroll_into_view_if_needed(timeout=2000)
            nb.click(timeout=4000)
            return True
    except Exception:
        pass
    try:
        nb = page.locator(
            "button:visible.btn-next:has-text('ถัดไป'), button:visible:has-text('ถัดไป'), a:visible:has-text('ถัดไป')"
        ).last
        if nb.count() > 0:
            nb.scroll_into_view_if_needed(timeout=2000)
            nb.click(timeout=4000)
            return True
    except Exception:
        pass
    try:
        return bool(page.evaluate(r"""() => {
            const b = document.querySelector('#button_next')
              || Array.from(document.querySelectorAll('button,a')).find(e => /ถัดไป/.test((e.textContent||'').trim()) && e.offsetParent !== null);
            if (b) { b.scrollIntoView({ block: 'center' }); b.click(); return true; }
            return false;
        }"""))
    except Exception:
        return False


def _bt44_step4_attach_docs(page: Page, rec: dict[str, Any], screenshot_dir: Path, log=print) -> dict[str, Any]:
    """Step 4 — แนบเอกสาร: 4.1 เอกสารลูกจ้าง (จับคู่ด้วย data-document-th) + รูปถ่ายครอป +
    เอกสารอื่นๆ (โมดอล) → ถัดไป → 4.2 เอกสารนายจ้าง → ถัดไป
    คืน {status, note, screenshot}
    """
    res: dict[str, Any] = {"status": "", "note": "", "screenshot": ""}
    seq = rec.get("seq", "?")
    name = rec.get("name", "")
    base = _safe_filename(f"{seq}_{name}")
    try:
        # รอหน้าแนบเอกสารพร้อม (มี input[type=file] ที่มี data-document-th)
        try:
            page.wait_for_function(
                r"""() => Array.from(document.querySelectorAll('input[type=file]')).some(i => i.getAttribute('data-document-th'))""",
                timeout=15000,
            )
        except PWTimeoutError:
            try:
                page.screenshot(path=str(screenshot_dir / f"{base}_step4_fail.png"), full_page=True)
                res["screenshot"] = str(screenshot_dir / f"{base}_step4_fail.png")
            except Exception:
                pass
            res["status"] = "ALERT"
            res["note"] = "Step 4: ไม่พบช่องอัปโหลดเอกสาร"
            return res
        _wait_loading_disappeared(page, timeout_ms=8000, log=log, allow_modal=True)
        page.wait_for_timeout(800)

        # ---- 4.1 เอกสารหลัก ----
        n_ok = 0
        n_try = 0
        problems: list[str] = []
        for d in rec.get("docs", []):
            path = (d.get("path") or "").strip()
            if not path:
                continue
            n_try += 1
            input_id = _bt44_find_doc_input_id(page, d["th_name"])
            if not input_id:
                problems.append(f"ไม่พบช่อง:{d['th_name'][:20]}")
                log(f"        ✗ ไม่พบช่องอัปโหลด: {d['th_name'][:40]}")
                continue
            r = _bt44_upload_one_doc(page, input_id, path, bool(d.get("is_photo")), log=log)
            if r["ok"]:
                n_ok += 1
                log(f"        ✓ {d['th_name'][:36]}: {r['note']}")
            else:
                problems.append(f"{d['th_name'][:16]}:{r['note'][:24]}")
                log(f"        ✗ {d['th_name'][:36]}: {r['note']}")

        # ---- เอกสารอื่นๆที่เกี่ยวข้อง (โมดอล ทีละไฟล์) ----
        others = rec.get("other_docs", [])
        n_oth_ok = 0
        for i, o in enumerate(others, 1):
            if _bt44_attach_one_other(page, (o.get("path") or "").strip(), i, log=log):
                n_oth_ok += 1

        try:
            page.screenshot(path=str(screenshot_dir / f"{base}_step4_1_uploaded.png"), full_page=True)
        except Exception:
            pass

        # ---- 4.1 → ถัดไป ----
        if not _bt44_click_next(page, log=log):
            res["status"] = "ALERT"
            res["note"] = f"Step 4.1: ไม่พบปุ่มถัดไป" + (f" | {';'.join(problems)[:100]}" if problems else "")
            return res
        page.wait_for_timeout(1500)
        _wait_loading_disappeared(page, timeout_ms=10000, log=log, allow_modal=True)

        # ตรวจ swal แจ้งเตือน (validation บล็อก) หลังถัดไป
        alert_txt = ""
        try:
            alert_txt = page.evaluate(r"""() => {
                const norm = s => (s||'').replace(/\s+/g,' ').trim();
                const sw = document.querySelector('.swal2-popup');
                if (sw && sw.offsetParent !== null) {
                    const t = sw.querySelector('.swal2-title'); const h = sw.querySelector('.swal2-html-container');
                    const txt = norm((t ? t.textContent : '') + ' ' + (h ? h.textContent : ''));
                    if (txt && !/สำเร็จ|เรียบร้อย|complete/i.test(txt)) return txt;
                }
                return '';
            }""") or ""
        except Exception:
            alert_txt = ""
        if alert_txt:
            try:
                page.screenshot(path=str(screenshot_dir / f"{base}_step4_1_alert.png"), full_page=True)
                res["screenshot"] = str(screenshot_dir / f"{base}_step4_1_alert.png")
            except Exception:
                pass
            res["status"] = "ALERT"
            res["note"] = f"Step 4.1 ถัดไปถูกบล็อก: {alert_txt[:140]}"
            return res
        log(f"      [Step 4.1] ✅ แนบเอกสารลูกจ้าง {n_ok}/{n_try} + อื่นๆ {n_oth_ok}/{len(others)} → ถัดไป")

        # ---- 4.2 เอกสารนายจ้าง → ถัดไป (ถ้ายังอยู่ขั้นแนบเอกสาร) ----
        page.wait_for_timeout(800)
        on_42 = False
        try:
            on_42 = bool(page.evaluate(r"""() => {
                const t = (document.body.innerText || '');
                const hasNext = !!document.querySelector('#button_next');
                const hasUpload = Array.from(document.querySelectorAll('input[type=file]')).some(i => i.getAttribute('data-document-th'))
                  || !!document.querySelector("[data-action='popup-upload']");
                return hasNext && (/เอกสารนายจ้าง|เอกสารของนายจ้าง/.test(t) || hasUpload);
            }"""))
        except Exception:
            on_42 = False
        if on_42:
            if _bt44_click_next(page, log=log):
                page.wait_for_timeout(1500)
                _wait_loading_disappeared(page, timeout_ms=10000, log=log, allow_modal=True)
                log("      [Step 4.2] ✅ เอกสารนายจ้าง → ถัดไป")
            else:
                log("      [Step 4.2] ⚠ ไม่พบปุ่มถัดไป (อาจข้ามขั้นแล้ว)")
        else:
            log("      [Step 4.2] (ไม่พบขั้นเอกสารนายจ้าง — อาจรวมกับ 4.1 หรือข้ามไปสรุปแล้ว)")

        try:
            page.screenshot(path=str(screenshot_dir / f"{base}_step4_done.png"), full_page=True)
            res["screenshot"] = str(screenshot_dir / f"{base}_step4_done.png")
        except Exception:
            pass

        warn = (" | ปัญหา: " + ";".join(problems)) if problems else ""
        res["status"] = "SUCCESS"
        res["note"] = f"แนบเอกสาร {n_ok}/{n_try} + อื่นๆ {n_oth_ok}/{len(others)}{warn}"
        return res
    except Exception as e:
        res["status"] = "ERROR"
        res["note"] = f"Step 4 error: {str(e)[:160]}"
        return res


def _bt44_step5_summary(page: Page, rec: dict[str, Any], screenshot_dir: Path, log=print) -> dict[str, Any]:
    """Step 5 สรุปคำขอ (3 หน้า ใช้ปุ่ม 'ถัดไป' เหมือนกัน):
      5.1 ข้อมูลผู้ยื่นคำขอ → กดถัดไป
      5.2 ข้อมูลนายจ้างและประเภทงาน → กดถัดไป
      5.3 เอกสารผู้ยื่นคำขอ → ติ๊ก consent (name='consent') → กด #consentButton
    """
    res: dict[str, Any] = {"status": "", "note": "", "screenshot": ""}
    base = re.sub(r'[^A-Za-z0-9_-]+', '_', (rec.get("name") or "row"))[:40]

    def _click_next_5(stage: str) -> bool:
        try:
            nb = page.locator("#button_next:visible").first
            if nb.count() > 0:
                nb.scroll_into_view_if_needed(timeout=2000)
                nb.click(timeout=4000)
                return True
        except Exception:
            pass
        try:
            nb2 = page.locator("button:visible:has-text('ถัดไป'), a:visible:has-text('ถัดไป')").last
            if nb2.count() > 0:
                nb2.scroll_into_view_if_needed(timeout=2000)
                nb2.click(timeout=4000)
                return True
        except Exception:
            pass
        try:
            ok = page.evaluate(r"""() => {
                const b = Array.from(document.querySelectorAll('button,a')).find(
                  x => x.offsetParent !== null && /ถัดไป/.test((x.textContent||'').trim())
                );
                if (b) { b.click(); return true; }
                return false;
            }""")
            return bool(ok)
        except Exception:
            return False

    try:
        page.wait_for_timeout(1200)
        _wait_loading_disappeared(page, timeout_ms=10000, log=log, allow_modal=True)

        # 5.1 → ถัดไป
        if not _click_next_5("5.1"):
            res["status"] = "ALERT"
            res["note"] = "Step 5.1 กด 'ถัดไป' ไม่สำเร็จ"
            return res
        page.wait_for_timeout(1500)
        _wait_loading_disappeared(page, timeout_ms=10000, log=log, allow_modal=True)
        log("      [Step 5.1] ✅ กด 'ถัดไป' → หน้า 2/3")

        # 5.2 → ถัดไป
        if not _click_next_5("5.2"):
            res["status"] = "ALERT"
            res["note"] = "Step 5.2 กด 'ถัดไป' ไม่สำเร็จ"
            return res
        page.wait_for_timeout(1500)
        _wait_loading_disappeared(page, timeout_ms=10000, log=log, allow_modal=True)
        log("      [Step 5.2] ✅ กด 'ถัดไป' → หน้า 3/3")

        # 5.3 — ติ๊ก consent
        consent_result = page.evaluate(r"""() => {
            const cb = document.querySelector('input[name="consent"]');
            if (!cb) return 'NO_CB';
            if (cb.checked) return 'ALREADY';
            try { cb.click(); } catch(e) {}
            if (cb.checked) return 'OK_CLICK';
            cb.checked = true;
            cb.dispatchEvent(new Event('change', { bubbles: true }));
            cb.dispatchEvent(new Event('input', { bubbles: true }));
            return cb.checked ? 'OK_FORCE' : 'FAIL';
        }""")
        if consent_result in ('NO_CB', 'FAIL'):
            try:
                page.screenshot(path=str(screenshot_dir / f"{base}_step5_3_no_consent.png"), full_page=True)
            except Exception:
                pass
            res["status"] = "ALERT"
            res["note"] = f"Step 5.3 ติ๊ก consent ไม่สำเร็จ ({consent_result})"
            return res
        page.wait_for_timeout(600)
        log(f"      [Step 5.3] ✓ ติ๊ก consent ({consent_result})")

        # คลิก #consentButton
        cb_clicked = False
        try:
            btn = page.locator("#consentButton").first
            if btn.count() > 0:
                btn.scroll_into_view_if_needed(timeout=2000)
                btn.click(timeout=5000, force=True)
                cb_clicked = True
        except Exception as e:
            log(f"      [Step 5.3] ⚠ consentButton click err: {str(e)[:80]}")
        if not cb_clicked:
            try:
                page.evaluate(r"""() => { const b = document.getElementById('consentButton'); if (b) b.click(); }""")
                cb_clicked = True
            except Exception:
                pass
        if not cb_clicked:
            res["status"] = "ALERT"
            res["note"] = "Step 5.3 กด 'ถัดไป' (consentButton) ไม่สำเร็จ"
            return res

        page.wait_for_timeout(2500)
        _wait_loading_disappeared(page, timeout_ms=15000, log=log, allow_modal=True)
        page.wait_for_timeout(1500)
        log("      [Step 5.3] ✅ กด 'ถัดไป' → Step 6 (ยืนยันตัวตน)")

        try:
            shot = screenshot_dir / f"{base}_step5_done.png"
            page.screenshot(path=str(shot), full_page=True)
            res["screenshot"] = str(shot)
        except Exception:
            pass

        res["status"] = "SUCCESS"
        res["note"] = "สรุปคำขอผ่าน 3 หน้า + ติ๊กรับทราบ"
        return res
    except Exception as e:
        res["status"] = "ERROR"
        res["note"] = f"Step 5 error: {str(e)[:160]}"
        return res


def _bt44_step6_verify_identity(page: Page, rec: dict[str, Any], screenshot_dir: Path, log=print) -> dict[str, Any]:
    """Step 6 ยืนยันตัวตน:
      6.1 กด 'อัปโหลด' (data-action=picture-uploader) → set image #upload-popup → กด 'ยืนยัน'
          → อ่านสถานะ (ยืนยันตัวตนสำเร็จ / ไม่สำเร็จ) → กด 'ถัดไป' (ไม่ว่าผลใด)
    """
    res: dict[str, Any] = {"status": "", "note": "", "screenshot": ""}
    base = re.sub(r'[^A-Za-z0-9_-]+', '_', (rec.get("name") or "row"))[:40]

    # หาไฟล์ภาพ 3x4 จาก rec['docs']
    photo_path: Path | None = None
    for d in rec.get("docs", []) or []:
        th = (d.get("th_name") or "")
        if "รูปถ่าย" in th:
            p = d.get("path")
            if p:
                photo_path = Path(p)
            break
    if not photo_path or not photo_path.exists():
        res["status"] = "ALERT"
        res["note"] = f"Step 6 ไม่พบไฟล์ภาพ 3x4: {photo_path}"
        return res

    try:
        # คลิก 'อัปโหลด' (picture-uploader) เพื่อเปิดโมดัล
        opened = False
        try:
            up_btn = page.locator(
                "[data-action='picture-uploader']:visible, "
                "button.lang_btn_upload_photo:visible, "
                "a.lang_btn_upload_photo:visible"
            ).first
            if up_btn.count() > 0:
                up_btn.scroll_into_view_if_needed(timeout=2000)
                up_btn.click(timeout=4000)
                opened = True
        except Exception:
            pass
        if not opened:
            try:
                page.evaluate(r"""() => {
                    const b = Array.from(document.querySelectorAll('button,a')).find(x =>
                      x.offsetParent !== null && (
                        x.getAttribute('data-action') === 'picture-uploader' ||
                        /อัปโหลด|อัพโหลด/.test((x.textContent||'').trim())
                      )
                    );
                    if (b) b.click();
                }""")
                opened = True
            except Exception:
                pass
        page.wait_for_timeout(1200)

        # set ไฟล์ภาพเข้า #upload-popup (ซ่อนอยู่ก็ใช้ได้)
        try:
            page.set_input_files('#upload-popup', str(photo_path))
            log(f"      [Step 6] ✓ set ภาพ: {photo_path.name}")
        except Exception as e:
            # fallback ผ่าน file chooser
            try:
                with page.expect_file_chooser() as fc:
                    page.locator(
                        "button:visible:has-text('เลือกไฟล์'), "
                        "label:visible:has-text('เลือกไฟล์')"
                    ).first.click()
                fc.value.set_files(str(photo_path))
                log(f"      [Step 6] ✓ set ภาพ (file chooser): {photo_path.name}")
            except Exception as e2:
                res["status"] = "ALERT"
                res["note"] = f"Step 6 set ภาพไม่สำเร็จ: {str(e)[:60]} / {str(e2)[:60]}"
                return res
        page.wait_for_timeout(1500)

        # กด 'ยืนยัน' ในโมดัล (btn-primary float-right)
        confirmed = False
        try:
            ok_btn = page.locator(
                "button.btn-primary.float-right:visible:has-text('ยืนยัน'), "
                ".modal:visible button:has-text('ยืนยัน'), "
                ".swal2-popup:visible button:has-text('ยืนยัน'), "
                "[role='dialog']:visible button:has-text('ยืนยัน')"
            ).first
            if ok_btn.count() > 0:
                ok_btn.click(timeout=5000)
                confirmed = True
        except Exception:
            pass
        if not confirmed:
            try:
                page.evaluate(r"""() => {
                    const cands = Array.from(document.querySelectorAll('button')).filter(x =>
                      x.offsetParent !== null && /ยืนยัน/.test((x.textContent||'').trim())
                      && (x.className||'').includes('btn-primary')
                    );
                    if (cands.length > 0) cands[0].click();
                }""")
                confirmed = True
            except Exception:
                pass
        if not confirmed:
            res["status"] = "ALERT"
            res["note"] = "Step 6 กด 'ยืนยัน' ในโมดัลไม่สำเร็จ"
            return res

        page.wait_for_timeout(3000)
        _wait_loading_disappeared(page, timeout_ms=25000, log=log, allow_modal=True)
        page.wait_for_timeout(3000)  # รอให้ระบบประมวลผลผลลัพธ์

        # อ่านสถานะ
        body_txt = ""
        try:
            body_txt = page.evaluate("() => (document.body.innerText || '').replace(/\\s+/g, ' ')") or ""
        except Exception:
            pass
        identity_status = "ไม่ทราบ"
        if "ยืนยันตัวตนสำเร็จ" in body_txt:
            identity_status = "สำเร็จ"
        elif "ยืนยันตัวตนไม่สำเร็จ" in body_txt:
            identity_status = "ไม่สำเร็จ"
        elif "ไม่สำเร็จ" in body_txt and "ตัวตน" in body_txt:
            identity_status = "ไม่สำเร็จ"
        elif "สำเร็จ" in body_txt and "ตัวตน" in body_txt:
            identity_status = "สำเร็จ"
        log(f"      [Step 6] 🔎 ยืนยันตัวตน: {identity_status}")

        try:
            shot = screenshot_dir / f"{base}_step6_status_{identity_status}.png"
            page.screenshot(path=str(shot), full_page=True)
            res["screenshot"] = str(shot)
        except Exception:
            pass

        # กด 'ถัดไป' (ไม่ว่าจะสำเร็จหรือไม่)
        nxt = False
        try:
            nb = page.locator("button:visible:has-text('ถัดไป'), a:visible:has-text('ถัดไป')").last
            if nb.count() > 0:
                nb.scroll_into_view_if_needed(timeout=2000)
                nb.click(timeout=4000)
                nxt = True
        except Exception:
            pass
        if not nxt:
            try:
                page.evaluate(r"""() => {
                    const b = Array.from(document.querySelectorAll('button,a')).find(x =>
                      x.offsetParent !== null && /ถัดไป/.test((x.textContent||'').trim())
                    );
                    if (b) b.click();
                }""")
                nxt = True
            except Exception:
                pass
        page.wait_for_timeout(2000)
        _wait_loading_disappeared(page, timeout_ms=15000, log=log, allow_modal=True)
        if nxt:
            log("      [Step 6] ✅ กด 'ถัดไป' หลังยืนยันตัวตน")
        else:
            log("      [Step 6] ⚠ ไม่พบปุ่ม 'ถัดไป' หลังยืนยัน")

        res["status"] = "SUCCESS"
        res["note"] = f"ยืนยันตัวตน: {identity_status}"
        return res
    except Exception as e:
        res["status"] = "ERROR"
        res["note"] = f"Step 6 error: {str(e)[:160]}"
        return res


def _bt44_parse_payment_pdf(pdf_bytes: bytes) -> dict[str, str]:
    """แยกข้อมูลใบแจ้งชำระเงินของ บต.44 (โครงสร้างเดียวกับ บต.30)
    คืน dict: bill_no, request_no, amount, due, txn_date, ref1, ref2,
              alien_ref, biller_name, email, total_amount
    """
    out = _bt30_parse_payment_pdf(pdf_bytes)
    # เพิ่ม field ที่ user ขอ
    out.setdefault("biller_name", "")
    out.setdefault("email", "")
    out.setdefault("total_amount", "")
    if not pdf_bytes:
        return out
    try:
        import io
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(pdf_bytes))
        txt = "\n".join((pg.extract_text() or "") for pg in reader.pages)
    except Exception:
        return out
    import re as _re

    def g(pat: str) -> str:
        m = _re.search(pat, txt)
        return m.group(1).strip() if m else ""

    # Biller Name (ผู้นำส่ง = รายชื่อคนต่างด้าว) — บรรทัดถัดจาก '(Biller Name)'
    out["biller_name"] = g(r"\(Biller Name\)\s*([^\n]+?)\s*(?:\n|เลข)")
    # Email
    out["email"] = g(r"\(Email\)\s*([^\s\n]+@[^\s\n]+)")
    # Total Amount = "รวมเงินที่ต้องชำระทั้งสิ้น"
    out["total_amount"] = (
        g(r"Total Amount\s*([0-9,]+\.[0-9]{2})")
        or g(r"รวม.{0,30}ทั้งสิ้น[^0-9]*([0-9,]+\.[0-9]{2})")
        or out.get("amount", "")
    )
    return out


def _bt44_step7_payment(
    page: Page, rec: dict[str, Any], screenshot_dir: Path, log=print,
) -> dict[str, Any]:
    """Step 7 (การชำระเงิน) — *** ส่งคำขอจริง! ดาวน์โหลดได้ครั้งเดียว ***
    7.1 กด 'ถัดไป' (เลือกช่องทางถ้าจำเป็น) → ส่งคำขอจริง
    7.2 บนหน้าผลสำเร็จ: กด 'พิมพ์แบบฟอร์มการชำระเงิน' → ดาวน์โหลด PDF
        + parse: เลขที่ใบแจ้งชำระเงิน, เลขที่คำขอ, ภายในวัน, ยอดรวม,
                 วันที่ทำรายการ, ชื่อคนต่างด้าว, อีเมล
    """
    res: dict[str, Any] = {
        "status": "", "note": "", "screenshot": "",
        "bill_no": "", "request_no": "", "due": "", "total_amount": "",
        "txn_date": "", "biller_name": "", "email": "", "amount": "",
        "pdf_file": "", "data_file": "",
    }
    seq = rec.get("seq", "?")
    name = rec.get("name", "")
    submitted_dir = screenshot_dir.parent / "bt44_submitted"
    submitted_dir.mkdir(parents=True, exist_ok=True)

    try:
        # ── 7.1 SAFETY GUARD: ตรวจว่าอยู่หน้าชำระเงินจริง (ก่อนกดส่ง) ──
        page.wait_for_timeout(1500)
        _wait_loading_disappeared(page, timeout_ms=10000, log=log, allow_modal=True)
        page.wait_for_timeout(1500)

        on_payment = page.evaluate(
            r"""() => /วิธีการชำระเงิน|e-?Payment|รายการชำระเงิน|ค่ายื่นคำขอ|ค่าธรรมเนียม/i.test(document.body.innerText || '')"""
        )
        # ถ้าอยู่หน้าผลสำเร็จแล้ว (ไม่ต้องกด 7.1 ซ้ำ)
        on_success = page.evaluate(
            r"""() => /ส่งใบคำขอ.*เรียบร้อย|พิมพ์แบบฟอร์มการชำระเงิน|พิมพ์.{0,8}ชำระเงิน/i.test(document.body.innerText || '')"""
        )
        if on_success:
            log("      [Step 7.1] (อยู่หน้าผลสำเร็จแล้ว — ข้ามการกด 'ถัดไป')")
        elif not on_payment:
            try:
                page.screenshot(path=str(screenshot_dir / f"{_safe_filename(seq + '_' + name)}_step7_unknown.png"), full_page=True)
            except Exception:
                pass
            res["status"] = "ALERT"
            res["note"] = "Step 7: ไม่พบหน้าชำระเงิน — ยกเลิกการกดส่งเพื่อความปลอดภัย"
            log("      ⛔ Step 7: ไม่พบหน้าชำระเงิน — ไม่กดส่งคำขอ (กันส่งผิดหน้า)")
            return res
        else:
            # เลือกช่องทาง e-Payment ถ้ายังไม่ได้เลือก (เผื่อมี checkbox/radio)
            page.evaluate(
                r"""() => {
                const cks = Array.from(document.querySelectorAll('input[type=checkbox],input[type=radio]'));
                for (const c of cks) {
                    const lbl = ((c.closest('label') && c.closest('label').innerText) ||
                                 (c.parentElement && c.parentElement.innerText) || '');
                    if (/e-?Payment|ชำระเงินผ่าน|KTB|กรุงไทย/i.test(lbl) && !c.checked) {
                        try { c.click(); } catch (e) {}
                        if (!c.checked) { c.checked = true; c.dispatchEvent(new Event('change', {bubbles:true})); }
                    }
                }
            }"""
            )
            page.wait_for_timeout(500)

            # 7.1 กดถัดไป (ส่งคำขอ)
            clicked = page.evaluate(
                r"""() => {
                const vis = el => { const r = el.getBoundingClientRect(); const s = getComputedStyle(el);
                    return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden'; };
                const cand = Array.from(document.querySelectorAll('button,a'))
                    .filter(e => vis(e) && !e.disabled && /ถัดไป/.test((e.textContent || '').trim())
                        && (e.textContent || '').trim().length < 20
                        && !/ย้อนกลับ|ยกเลิก/.test((e.textContent || '').trim()));
                if (cand.length) { cand[cand.length - 1].click(); return true; }
                return false;
            }"""
            )
            if not clicked:
                res["status"] = "ALERT"
                res["note"] = "Step 7.1 กด 'ถัดไป' ไม่สำเร็จ"
                return res
            log("      [Step 7.1] · กด 'ถัดไป' → ส่งคำขอจริง")

            # รอหน้าผลสำเร็จ (เผื่อมี modal ยืนยัน → กดยืนยันให้)
            ok_success = False
            for _ in range(30):
                if page.evaluate(
                    r"""() => /เลขที่คำขอ|ส่งใบคำขอ.*เรียบร้อย|เรียบร้อยแล้ว|E-?Tracking|พิมพ์แบบฟอร์มการชำระเงิน|พิมพ์.{0,8}ชำระเงิน/i.test(document.body.innerText || '')"""
                ):
                    ok_success = True
                    break
                page.evaluate(
                    r"""() => {
                    const vis = el => { const r = el.getBoundingClientRect(); const s = getComputedStyle(el);
                        return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden'; };
                    const b = Array.from(document.querySelectorAll('.swal2-confirm,button,a'))
                        .find(e => vis(e) && /^(ยืนยัน|ตกลง|confirm|ใช่|ส่งคำขอ)/i.test((e.textContent || '').trim())
                            && (e.textContent || '').trim().length < 20);
                    if (b) b.click();
                }"""
                )
                page.wait_for_timeout(1000)
            if not ok_success:
                res["status"] = "ALERT"
                res["note"] = "Step 7.1: กดถัดไปแล้ว แต่ไม่พบหน้าผลสำเร็จ"
                try:
                    shot = screenshot_dir / f"{_safe_filename(seq + '_' + name)}_step7_no_success.png"
                    page.screenshot(path=str(shot), full_page=True)
                    res["screenshot"] = str(shot)
                except Exception:
                    pass
                return res
            log("      [Step 7.1] ✅ พบหน้าผลสำเร็จ")

        # ── 7.2 รอหน้าผลสำเร็จโหลดค่าจริง ──
        for _ in range(20):
            loaded = page.evaluate(
                r"""() => {
                const t = document.body.innerText || '';
                return /\d{10,}/.test(t) || /พิมพ์.{0,8}ชำระเงิน/.test(t);
            }"""
            )
            if loaded:
                break
            page.wait_for_timeout(1000)
        page.wait_for_timeout(1200)

        # 7.2 เก็บข้อมูลหน้าผลสำเร็จ (best-effort — แหล่งหลักคือ PDF)
        dom_data = page.evaluate(
            r"""() => {
            const lines = (document.body.innerText || '').split('\n').map(s => s.trim()).filter(s => s.length);
            const after = (re) => {
                for (let i = 0; i < lines.length; i++) {
                    if (re.test(lines[i])) {
                        const same = lines[i].replace(re, '').trim();
                        if (same) return same;
                        if (i + 1 < lines.length) return lines[i + 1];
                    }
                }
                return '';
            };
            const out = {};
            out.request_no = (after(/^เลขที่คำขอ/) || '').replace(/[^0-9]/g, '');
            out.subject = after(/^ระบบได้รับคำขอเรื่อง/);
            out.submit_date = after(/^วันที่ยื่นคำขอ/);
            out.alien = after(/^คนต่างด้าว/);
            out.pay_method = after(/^วิธีการชำระเงิน/);
            out.ref1 = (after(/^หมายเลขอ้างอิง\s*1/) || '').replace(/[^0-9]/g, '');
            out.ref2 = (after(/^หมายเลขอ้างอิง\s*2/) || '').replace(/[^0-9]/g, '');
            out.amount = after(/^ยอดชำระ/);
            const m = (document.body.innerText || '').match(/ภายในวันที่\s*([^\n]+)/);
            out.due = m ? m[1].trim() : '';
            out.full_text = (document.body.innerText || '');
            return out;
        }"""
        )
        dom_request_no = (dom_data.get("request_no") or "").strip()

        try:
            shot = screenshot_dir / f"{_safe_filename(seq + '_' + name)}_step7_success.png"
            page.screenshot(path=str(shot), full_page=True)
            res["screenshot"] = str(shot)
        except Exception:
            pass

        # ── 7.2 ตรวจปุ่ม 'พิมพ์แบบฟอร์มการชำระเงิน' (เก็บรายละเอียด เพื่อ diagnose + fallback URL) ──
        btn_info = page.evaluate(
            r"""() => {
            const vis = el => { const r = el.getBoundingClientRect(); const s = getComputedStyle(el);
                return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden'; };
            const b = Array.from(document.querySelectorAll('button,a,input[type=button],input[type=submit]'))
                .find(e => vis(e) && /พิมพ์.{0,8}(ฟอร์ม).{0,8}ชำระเงิน|พิมพ์.{0,8}ชำระเงิน/.test((e.textContent || e.value || '').trim()));
            if (!b) return null;
            // หาลิงก์ <a> ที่อยู่ใกล้ๆ ปุ่ม (บางครั้ง print เป็น <a target="_blank" href="...pdf">)
            const a = (b.tagName === 'A') ? b : b.closest('a') || b.querySelector('a');
            return {
                tag: b.tagName, id: b.id || '', cls: b.className || '',
                text: ((b.textContent || b.value) || '').trim().slice(0, 80),
                onclick: b.getAttribute('onclick') || '',
                href: (a && a.getAttribute('href')) || b.getAttribute('href') || '',
                target: (a && a.getAttribute('target')) || b.getAttribute('target') || '',
                formaction: b.getAttribute('formaction') || '',
                disabled: !!b.disabled,
            };
        }"""
        )
        log(f"      · 7.2 ปุ่มพิมพ์: {btn_info}")

        # extract URL จาก onclick/href ถ้ามี (เผื่อ popup ถูกบล็อก)
        direct_url = ""
        if btn_info:
            href = (btn_info.get("href") or "").strip()
            onclick = (btn_info.get("onclick") or "").strip()
            if href and href not in ("#", "javascript:void(0)", "javascript:;"):
                direct_url = href
            else:
                import re as _re
                # patterns: window.open('URL'), location.href='URL', GetDocumentConfirm('URL'), 'URL.pdf'
                for pat in (r"""window\.open\(\s*['"]([^'"]+)['"]""",
                            r"""location\.href\s*=\s*['"]([^'"]+)['"]""",
                            r"""['"]([^'"]+\.pdf[^'"]*)['"]""",
                            r"""['"]([^'"]+(?:GetBill|Print|Download|Payment)[^'"]*)['"]"""):
                    m = _re.search(pat, onclick, _re.IGNORECASE)
                    if m:
                        direct_url = m.group(1)
                        break
            if direct_url and direct_url.startswith("/"):
                # absolute path → prepend origin
                origin = page.evaluate("() => location.origin") or ""
                direct_url = origin + direct_url

        if direct_url:
            log(f"      · 7.2 พบ URL ตรง → จะลอง fetch โดยตรง: {direct_url[:120]}")

        # ── 7.2 ดาวน์โหลด 'พิมพ์แบบฟอร์มการชำระเงิน' (retry สูงสุด 3 ครั้ง) ──
        # ใช้ Playwright locator click (real user gesture — ไม่ถูก popup blocker)
        def _do_click():
            try:
                loc = page.locator(
                    r"button:has-text('พิมพ์แบบฟอร์มการชำระเงิน'), "
                    r"a:has-text('พิมพ์แบบฟอร์มการชำระเงิน'), "
                    r"button:has-text('พิมพ์ชำระเงิน'), "
                    r"a:has-text('พิมพ์ชำระเงิน')"
                ).first
                try:
                    loc.scroll_into_view_if_needed(timeout=3000)
                except Exception:
                    pass
                loc.click(timeout=8000, no_wait_after=True)
                return
            except Exception:
                pass
            # fallback: JS click
            page.evaluate(
                r"""() => {
                const vis = el => { const r = el.getBoundingClientRect(); const s = getComputedStyle(el);
                    return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden'; };
                const b = Array.from(document.querySelectorAll('button,a,input[type=button],input[type=submit]'))
                    .find(e => vis(e) && /พิมพ์.{0,8}(ฟอร์ม).{0,8}ชำระเงิน|พิมพ์.{0,8}ชำระเงิน/.test((e.textContent || e.value || '').trim()));
                if (b) b.click();
            }"""
            )

        body: bytes = b""
        err_last = ""
        for attempt in range(1, 4):
            b, err = _grab_pdf_after_click(page, _do_click, log=lambda *a: None,
                                           prefetch_url=direct_url if attempt == 1 else "")
            if b and len(b) >= 500 and b[:5] == b"%PDF-":
                body = b
                break
            err_last = err or "ไม่ได้ไฟล์"
            log(f"      · ⚠ 7.2 ดาวน์โหลดใบชำระเงินครั้งที่ {attempt} ไม่สำเร็จ ({err_last}) — ลองใหม่")
            page.wait_for_timeout(2500)

        # ── Fallback สุดท้าย: ถ้ายังไม่ได้ + รู้ direct_url → fetch ผ่าน page context ──
        if not body and direct_url:
            log(f"      · 7.2 fallback fetch URL ตรงผ่าน page context...")
            b, err = _popup_fetch_bytes(page, direct_url)
            if b and len(b) >= 500 and b[:5] == b"%PDF-":
                body = b
                log(f"      · ✓ 7.2 fetch URL ตรงสำเร็จ")
            else:
                log(f"      · ⚠ fallback fetch ก็ล้มเหลว: {err}")

        # 7.3 parse PDF (แหล่งข้อมูลหลัก)
        pdf_info = _bt44_parse_payment_pdf(body) if body else {}
        request_no = dom_request_no or pdf_info.get("request_no", "")
        bill_no = pdf_info.get("bill_no", "")
        due = pdf_info.get("due", "") or dom_data.get("due", "")
        txn_date = pdf_info.get("txn_date", "")
        amount = pdf_info.get("amount", "") or dom_data.get("amount", "")
        total_amount = pdf_info.get("total_amount", "") or amount
        biller_name = pdf_info.get("biller_name", "") or dom_data.get("alien", "") or rec.get("name", "")
        email = pdf_info.get("email", "")

        res["bill_no"] = bill_no
        res["request_no"] = request_no
        res["due"] = due
        res["txn_date"] = txn_date
        res["amount"] = amount
        res["total_amount"] = total_amount
        res["biller_name"] = biller_name
        res["email"] = email

        log(f"      · 7.2 เลขใบแจ้งชำระเงิน={bill_no or '-'} | เลขคำขอ={request_no or '-'}")
        log(f"      · ยอด={amount or '-'} (รวม {total_amount or '-'}) | ภายใน={due or '-'} | ทำรายการ={txn_date or '-'}")
        log(f"      · ผู้ชำระ={biller_name or '-'} | อีเมล={email or '-'}")

        # บันทึก JSON
        stem = _safe_filename(f"{request_no or seq}_{name}_payment")
        try:
            data_payload = {**dom_data, **pdf_info, "row": {
                "seq": seq, "name": name, "alien_ref": rec.get("alien_ref", ""),
                "workpermit_no": rec.get("workpermit_no", ""),
            }}
            (submitted_dir / f"{stem}.json").write_text(
                json.dumps(data_payload, ensure_ascii=False, indent=2), encoding="utf-8")
            res["data_file"] = f"{stem}.json"
        except Exception as e:
            log(f"      · ⚠ เขียน JSON ไม่สำเร็จ: {e}")

        # บันทึก PDF
        if body:
            permit = _safe_filename(rec.get("workpermit_no", "") or "")
            name_parts = [p for p in (request_no or seq, permit, name) if p]
            pdf_name = "_".join(_safe_filename(p) for p in name_parts) + "_payment.pdf"
            try:
                (submitted_dir / pdf_name).write_bytes(body)
                res["pdf_file"] = pdf_name
                log(f"      · ✓ 7.2 ดาวน์โหลดใบชำระเงินแล้ว → bt44_submitted/{pdf_name} ({len(body)//1024} KB)")
            except Exception as e:
                log(f"      · ⚠ เขียนไฟล์ PDF ไม่สำเร็จ: {e}")
        else:
            log(f"      · ⛔ 7.2 ดาวน์โหลดใบชำระเงินไม่สำเร็จหลังลอง 3 ครั้ง ({err_last}) — "
                f"กรุณาดาวน์โหลดเองทันที (เลขคำขอ {request_no or '-'}) เพราะดาวน์โหลดได้ครั้งเดียว")

        if request_no and res["pdf_file"]:
            res["status"] = "SUCCESS"
            res["note"] = (
                f"ส่งคำขอแล้ว | เลขใบแจ้ง={bill_no or '-'} | เลขคำขอ={request_no} | "
                f"ยอดรวม={total_amount or '-'} | ภายใน={due or '-'} | PDF={res['pdf_file']}"
            )
        elif request_no:
            res["status"] = "PARTIAL"
            res["note"] = f"ส่งคำขอแล้ว เลขคำขอ={request_no} แต่ดาวน์โหลด PDF ไม่สำเร็จ ({err_last})"
        else:
            res["status"] = "ALERT"
            res["note"] = f"ส่งคำขอแล้วแต่ไม่พบเลขที่คำขอ + PDF ล้มเหลว ({err_last})"
        return res
    except Exception as e:
        res["status"] = "ERROR"
        res["note"] = f"Step 7 error: {str(e)[:160]}"
        return res


def _bt44_step2_collect_details(page: Page, log=print) -> dict[str, str]:
    """เก็บรายละเอียดทั้งหมดของหน้า Step 2 (ก่อนแก้ไขที่อยู่) เพื่อทำรายงาน:
       - ข้อมูลผู้ยื่นคำขอ: ชื่อ Eng/ไทย, สัญชาติ, เกิดวันที่, อายุ
       - ข้อมูลสำหรับติดต่อ: โทรศัพท์, อีเมล
       - ที่อยู่ในประเทศไทย
       - ข้อมูลใบอนุญาตทำงานปัจจุบัน: เลขที่, ออกให้ที่, วันที่ออก, วันที่หมดอายุ
    คืน dict ของ key→value ตามที่อ่านได้ (ค่าว่างถ้าไม่พบ)
    """
    try:
        data = page.evaluate(
            r"""() => {
            const lines = (document.body.innerText || '')
                .split('\n').map(s => s.trim()).filter(s => s.length);

            // จับ "ค่าหลัง label" ในบรรทัดเดียวกัน หรือบรรทัดถัดไป (ถ้ายังว่าง)
            // labelRe: regex ของ label, stopRe: regex หยุด (ถ้าค่าหายไปแล้วเจอ label อื่น)
            const after = (labelRe, stopRe) => {
                for (let i = 0; i < lines.length; i++) {
                    const m = lines[i].match(labelRe);
                    if (!m) continue;
                    // ลองตัด label ออกในบรรทัดเดียวกัน
                    const same = lines[i].replace(labelRe, '').trim();
                    if (same && !/^[:：\-]?$/.test(same)) return same;
                    // มิเช่นนั้นใช้บรรทัดถัดไป (skip บรรทัดว่าง/มี label อื่น)
                    for (let j = i + 1; j < Math.min(i + 4, lines.length); j++) {
                        const v = lines[j].trim();
                        if (!v) continue;
                        if (stopRe && stopRe.test(v)) break;
                        return v;
                    }
                    return '';
                }
                return '';
            };

            // ที่อยู่: รับหลายบรรทัดต่อกันจนเจอ section ใหม่
            const addressGet = () => {
                const startRe = /^ที่อยู่ในประเทศไทย/;
                const stopRe = /^(ข้อมูลใบอนุญาตทำงานปัจจุบัน|ข้อมูลสำหรับติดต่อ|ข้อมูลผู้ยื่นคำขอ|ที่อยู่ตามทะเบียน|สถานที่ทำงาน)/;
                const addrRe = /(เลขที่|ซอย|ถนน|แขวง|ตำบล|เขต|อำเภอ|จังหวัด|รหัสไปรษณีย์|\d{5})/;
                let i0 = -1;
                for (let i = 0; i < lines.length; i++) {
                    if (startRe.test(lines[i])) { i0 = i; break; }
                }
                if (i0 < 0) return '';
                const parts = [];
                for (let i = i0 + 1; i < Math.min(i0 + 10, lines.length); i++) {
                    const v = lines[i].trim();
                    if (!v) continue;
                    if (stopRe.test(v)) break;
                    if (/^(ที่อยู่|Address)[\s:：]*$/.test(v)) continue;  // ข้าม label เปล่า
                    if (/^แก้ไขข้อมูล/.test(v)) continue;
                    // ค่าอาจอยู่ในบรรทัดเดียวกับ label
                    const m = v.match(/^ที่อยู่[\s:：]+(.+)$/);
                    if (m) { parts.push(m[1].trim()); break; }
                    if (addrRe.test(v)) { parts.push(v); break; }
                }
                return parts.join(' ').trim();
            };

            const stopGen = /^(ข้อมูล|ที่อยู่|แก้ไขข้อมูล|ข้อมูลใบอนุญาต)/;
            return {
                name_en: after(/^ชื่อคนต่างด้าว\s*\(\s*Eng\s*\)/i, stopGen),
                name_th: after(/^ชื่อคนต่างด้าว\s*\(\s*ไทย\s*\)/i, stopGen),
                nationality: after(/^สัญชาติ$/, stopGen),
                birthdate: after(/^เกิดวันที่$/, stopGen),
                age: after(/^อายุ$/, stopGen),
                phone: after(/^โทรศัพท์$/, stopGen),
                email: after(/^อีเมล$/, stopGen),
                address: addressGet(),
                permit_no: after(/^เลขที่$/, stopGen),
                permit_province: after(/^ออกให้ที่\s*\(?\s*จังหวัด\s*\)?\s*$/, stopGen),
                permit_issue: after(/^วันที่ออกเอกสาร$/, stopGen),
                permit_expire: after(/^วันที่เอกสารหมดอายุ$/, stopGen),
            };
        }"""
        )
        return {k: (str(v) if v is not None else "") for k, v in (data or {}).items()}
    except Exception as e:
        log(f"      ⚠ collect details error: {str(e)[:160]}")
        return {}


def _bt44_process_record(
    page: Page, rec: dict[str, Any], screenshot_dir: Path,
    log=print, dry_run: bool = False, dry_stop_at: str = "step2",
    check_docs: bool = True,
) -> dict[str, str]:

    """ประมวลผลประวัติหนึ่งแถวผ่านทั้ง 3 ขั้นตอน
    Step 1: ค้นหาคนต่างด้าว (บันทึก)
    Step 2.1: เช็ก checkbox ข้อมูลใจสำสัญญา + กด ถัดไป
    Step 2.2: แก้ไขที่อยู่ + บันทึก + ยืนยัน
    Step 2.3: ตรวจสอบเลขที่ใบอนุญาตตรงกันไหม

    dry_run=True : โหมด **ทดลองยื่น** — ทำตาม `dry_stop_at`:
        * "step2" : หยุดที่ Step 2.3 (หลังตรวจเลขใบอนุญาต — ไม่แก้ไขที่อยู่)
        * "step3" : ทำ Step 2.2 (แก้ที่อยู่+บันทึก) + Step 3.1 (เปลี่ยนนายจ้าง)
                    + Step 3.3-3.4 (สถานที่ทำงาน/ประเภทกิจการ/งาน) จากนั้นหยุดก่อน Step 4
                    (ไม่แนบเอกสาร/ไม่ไปสรุป/ไม่ส่งคำขอ)
    check_docs=False : ข้ามการตรวจไฟล์แนบ (Step 4) ก่อนเริ่ม — สำหรับทดสอบ flow
    """
    dry_stop_at = (dry_stop_at or "step2").lower()
    if dry_stop_at not in ("step2", "step3"):
        dry_stop_at = "step2"
    result = {"status": "", "note": ""}

    # ===== ตรวจไฟล์แนบ (Step 4) ก่อนเริ่ม — ถ้าไม่ผ่านให้ข้าม (ไม่รัน) =====
    # check_docs=False หรือ ENV BT44_SKIP_DOC_CHECK=1 → ข้ามการตรวจไฟล์
    import os as _os
    _skip_env = _os.environ.get("BT44_SKIP_DOC_CHECK") == "1"
    if (not check_docs) or _skip_env:
        _reason = "ผู้ใช้ปิด 'ตรวจไฟล์แนบ'" if not check_docs else "BT44_SKIP_DOC_CHECK=1"
        log(f"      [Step 4-ตรวจไฟล์] ⚠ ข้าม ({_reason}) — สำหรับทดสอบ flow เท่านั้น")
    else:
        doc_problems = _bt44_validate_row_docs(rec)
        if doc_problems:
            result["status"] = "SKIP"
            result["note"] = "เอกสารแนบไม่ผ่าน (ไม่รัน): " + " | ".join(doc_problems[:6])
            log(f"      [Step 4-ตรวจไฟล์] ⛔ ข้าม (ไม่รัน): {result['note']}")
            return result

    # ===== Step 1: ค้นหาคนต่างด้าว =====
    step1_res = _bt44_fill_search_one(page, rec, screenshot_dir, log=log)
    if step1_res.get("status") != "SUCCESS":
        log(f"      [Step 1] ❌ ไม่สำเร็จ: {step1_res.get('status')}")
        return step1_res

    log(f"      [Step 1] ✅ เสร็จ — ดำเนินการ Step 2")
    
    # Close any modal that Step 1 might have left open
    page.wait_for_timeout(500)
    try:
        page.evaluate(r"""() => {
            const modals = document.querySelectorAll('.modal, .swal2-popup, [role="dialog"]');
            for (const m of modals) {
                if (m.offsetParent !== null) {
                    // Prefer explicit action buttons first
                    const actionBtn = Array.from(m.querySelectorAll('button, a'))
                      .find(b => /ปิด|ยืนยัน|ตกลง|ok/i.test((b.textContent || '').trim()) && b.offsetParent !== null);
                    if (actionBtn) {
                        actionBtn.click();
                        continue;
                    }
                    const closeBtn = m.querySelector('.close, .btn-close, .swal2-close, .btn[onclick*="close"]');
                    if (closeBtn) closeBtn.click();
                }
            }
        }""")
        page.wait_for_timeout(500)
    except Exception:
        pass

    page.wait_for_timeout(800)

    # ===== Step 2.1: เช็ก Consent Checkbox =====
    consent_ok = _bt44_step2_consent(page, log=log)
    if not consent_ok:
        log(f"      [Step 2.1] ❌ ไม่สำเร็จ")
        result["status"] = "ALERT"
        result["note"] = "Step 2.1: consent/next transition failed"
        return result
    
    log(f"      [Step 2.1] ✅ เสร็จ")

    # ===== Step 2.3: ตรวจสอบเลขที่ใบอนุญาต (ก่อน Step 2.2) =====
    log(f"      [Step 2.3-pre] ตรวจสอบเลขที่ใบอนุญาตก่อนแก้ไขที่อยู่...")
    permit_match, current_permit = _bt44_step2_verify_permit(page, rec, log=log)
    if not current_permit:
        result["status"] = "ALERT"
        result["note"] = "Step 2.3-pre: permit number not found"
        return result
    if not permit_match:
        result["status"] = "SKIP"
        result["note"] = f"Permit mismatch: expected {rec.get('workpermit_no','')}, got {current_permit}"
        return result

    log(f"      [Step 2.3-pre] ✅ ผ่าน — Permit: {current_permit}")

    # เก็บรายละเอียดหน้า Step 2 (ใช้ใน report ทุกโหมด dry-run)
    _dry_details = None
    if dry_run:
        _dry_details = _bt44_step2_collect_details(page, log=log)
        for k, v in (_dry_details or {}).items():
            result[f"dry_{k}"] = v
        result["dry_permit_match"] = "ตรง ✓" if permit_match else "ไม่ตรง ✗"
        result["dry_expected_permit"] = rec.get("workpermit_no", "")
        result["dry_current_permit"] = current_permit

    # ===== โหมด 'ทดลองยื่น' หยุดที่ Step 2.3 — เก็บข้อมูล + ข้ามไป record ถัดไป =====
    if dry_run and dry_stop_at == "step2":
        details = _dry_details or {}

        log(f"      [Dry-run] 📋 รายละเอียดที่อ่านได้:")
        log(f"        · ชื่อ (Eng): {details.get('name_en','-')}")
        log(f"        · ชื่อ (ไทย): {details.get('name_th','-')}")
        log(f"        · สัญชาติ: {details.get('nationality','-')} | เกิด: {details.get('birthdate','-')} | อายุ: {details.get('age','-')}")
        log(f"        · โทรศัพท์: {details.get('phone','-')} | อีเมล: {details.get('email','-')}")
        log(f"        · ที่อยู่: {details.get('address','-')}")
        log(f"        · ใบอนุญาตเลขที่: {details.get('permit_no','-')} | จังหวัด: {details.get('permit_province','-')}")
        log(f"        · ออกเอกสาร: {details.get('permit_issue','-')} | หมดอายุ: {details.get('permit_expire','-')}")

        try:
            shot = screenshot_dir / f"{_safe_filename(str(rec.get('seq','?')) + '_' + rec.get('name',''))}_dryrun_step2_3.png"
            page.screenshot(path=str(shot), full_page=True)
            result["screenshot"] = str(shot)
        except Exception:
            pass
        result["status"] = "DRY_OK"
        result["note"] = (
            f"🧪 ทดลองยื่น ผ่านถึง Step 2.3 | Permit: {current_permit} ({result['dry_permit_match']}) | "
            f"ใบอนุญาตหมดอายุ: {details.get('permit_expire','-')} | "
            f"อีเมล: {details.get('email','-')}"
        )
        # รีเซ็ตกลับไปหน้าเริ่มต้น บต.44 เพื่อให้ record ถัดไปเริ่มได้
        try:
            log("      [Dry-run] ↺ รีเซ็ตกลับหน้าเริ่มต้น บต.44 สำหรับ record ถัดไป...")
            _open_bt44_form(page, log=lambda *a: None)
        except Exception as e:
            log(f"      [Dry-run] ⚠ รีเซ็ตฟอร์มไม่สำเร็จ: {str(e)[:120]}")
        return result

    # ===== Step 2.2: แก้ไขที่อยู่ =====
    addr_res = _bt44_step2_edit_address(page, rec, screenshot_dir, log=log)
    if addr_res.get("status") != "SUCCESS":
        log(f"      [Step 2.2] ❌ มีปัญหา: {addr_res.get('note')}")
        result["status"] = "ALERT"
        result["note"] = f"Step 2.2: {addr_res.get('note')}"
        return result

    log(f"      [Step 2.2] ✅ เสร็จ")

    # ===== หลัง Step 2.2 ต้องกด 'ถัดไป' =====
    next_ok = False
    try:
        next_btn = page.locator("#gonextSubmit").first
        next_btn.click(timeout=5000, force=True)
        next_ok = True
    except Exception:
        try:
            next_btn2 = page.locator("button, a").filter(has_text="ถัดไป").first
            if next_btn2.count() > 0:
                next_btn2.click(timeout=5000, force=True)
                next_ok = True
        except Exception:
            next_ok = False

    if not next_ok:
        result["status"] = "ALERT"
        result["note"] = "Step 2.2: next button click failed after save confirm"
        return result

    log("      [Step 2.2→ถัดไป] ✅ กด 'ถัดไป' หลังบันทึกสำเร็จแล้ว")
    page.wait_for_timeout(1200)
    _wait_loading_disappeared(page, timeout_ms=10000, log=log)

    # ===== Step 3.1: เปลี่ยนนายจ้าง =====
    step3_1_res = _bt44_step3_change_employer(page, rec, log=log)
    if step3_1_res.get("status") == "SKIP":
        result["status"] = "SKIP"
        result["note"] = step3_1_res.get("note", "Step 3.1 skipped")
        return result
    if step3_1_res.get("status") != "SUCCESS":
        result["status"] = "ALERT"
        result["note"] = step3_1_res.get("note", "Step 3.1 failed")
        return result
    log("      [Step 3.1] ✅ เสร็จ")

    # ===== Step 3.3-3.4: เลือกสถานที่ทำงาน + ประเภทกิจการ + ตรวจประเภทงาน + ลักษณะงาน + ถัดไป =====
    step3_3_res = _bt44_step3_workplace(page, rec, log=log)
    if step3_3_res.get("status") == "SKIP":
        result["status"] = "SKIP"
        result["note"] = step3_3_res.get("note", "Step 3.3 skipped")
        return result
    if step3_3_res.get("status") != "SUCCESS":
        result["status"] = "ALERT"
        result["note"] = step3_3_res.get("note", "Step 3.3 failed")
        return result
    log("      [Step 3.3-3.4] ✅ เสร็จ")

    # ===== โหมด 'ทดลองยื่น' หยุดที่ Step 3 — ไม่แนบเอกสาร/ไม่สรุป/ไม่ส่งคำขอ =====
    if dry_run and dry_stop_at == "step3":
        details = _dry_details or {}
        log(f"      [Dry-run/Step3] ✅ ผ่าน Step 1-3 ครบ — หยุดก่อน Step 4")
        log(f"        · นายจ้างใหม่: {rec.get('change_emp_keyword','-')}")
        log(f"        · สถานที่ทำงาน: {rec.get('workplace_branch','-')}")
        log(f"        · ประเภทกิจการ: {rec.get('work_biz','-')}")
        log(f"        · ประเภทงาน: {rec.get('work_permit_job','-')} | ลักษณะงาน: {rec.get('work_detail','-')}")
        try:
            shot = screenshot_dir / f"{_safe_filename(str(rec.get('seq','?')) + '_' + rec.get('name',''))}_dryrun_step3.png"
            page.screenshot(path=str(shot), full_page=True)
            result["screenshot"] = str(shot)
        except Exception:
            pass
        result["status"] = "DRY_OK"
        result["note"] = (
            f"🧪 ทดลองยื่น ผ่านถึง Step 3.3-3.4 | Permit: {current_permit} "
            f"({result.get('dry_permit_match','-')}) | "
            f"นายจ้าง: {rec.get('change_emp_keyword','-')} | "
            f"สถานที่ทำงาน: {rec.get('workplace_branch','-')} | "
            f"ประเภทกิจการ: {rec.get('work_biz','-')}"
        )
        try:
            log("      [Dry-run/Step3] ↺ รีเซ็ตกลับหน้าเริ่มต้น บต.44 สำหรับ record ถัดไป...")
            _open_bt44_form(page, log=lambda *a: None)
        except Exception as e:
            log(f"      [Dry-run/Step3] ⚠ รีเซ็ตฟอร์มไม่สำเร็จ: {str(e)[:120]}")
        return result

    # ===== Step 4: แนบเอกสาร (4.1 ลูกจ้าง + 4.2 นายจ้าง) =====
    step4_res = _bt44_step4_attach_docs(page, rec, screenshot_dir, log=log)
    if step4_res.get("screenshot"):
        result["screenshot"] = step4_res["screenshot"]
    if step4_res.get("status") != "SUCCESS":
        result["status"] = step4_res.get("status") or "ALERT"
        result["note"] = f"Step 4: {step4_res.get('note','')}"
        return result
    log("      [Step 4] ✅ เสร็จ")

    # ===== Step 5: สรุปคำขอ (3 หน้า + ติ๊กรับทราบ) =====
    step5_res = _bt44_step5_summary(page, rec, screenshot_dir, log=log)
    if step5_res.get("screenshot"):
        result["screenshot"] = step5_res["screenshot"]
    if step5_res.get("status") != "SUCCESS":
        result["status"] = step5_res.get("status") or "ALERT"
        result["note"] = f"Step 5: {step5_res.get('note','')}"
        return result
    log("      [Step 5] ✅ เสร็จ")

    # ===== Step 6: ยืนยันตัวตน (อัปโหลดภาพ + อ่านสถานะ + ถัดไป) =====
    step6_res = _bt44_step6_verify_identity(page, rec, screenshot_dir, log=log)
    if step6_res.get("screenshot"):
        result["screenshot"] = step6_res["screenshot"]
    if step6_res.get("status") not in ("SUCCESS",):
        # ยังไม่ fail ทั้งหมด — แค่บันทึกเป็นหมายเหตุ (per spec: ทั้งสำเร็จ/ไม่สำเร็จ ก็ผ่านไปได้)
        log(f"      [Step 6] ⚠ {step6_res.get('note','')}")

    identity_note = step6_res.get("note", "") or "Step 6 ไม่ทราบผล"

    # ===== Step 7: การชำระเงิน (กดถัดไป → ส่งคำขอจริง → ดาวน์โหลด PDF + parse) =====
    step7_res = _bt44_step7_payment(page, rec, screenshot_dir, log=log)
    if step7_res.get("screenshot"):
        result["screenshot"] = step7_res["screenshot"]
    # ถ่ายโอนค่าทั้งหมดของ Step 7 เข้า result (ใช้ใน Excel report)
    for k in ("bill_no", "request_no", "due", "txn_date", "amount",
              "total_amount", "biller_name", "email", "pdf_file", "data_file"):
        result[k] = step7_res.get(k, "")

    if step7_res.get("status") == "SUCCESS":
        log("      [Step 7] ✅ เสร็จ")
        result["status"] = "SUCCESS"
        result["note"] = (
            f"✅ บต.44 Step 1-7 เสร็จ | Permit: {current_permit} | "
            f"เลขใบแจ้ง: {step7_res.get('bill_no','-')} | "
            f"เลขคำขอ: {step7_res.get('request_no','-')} | "
            f"ยอดรวม: {step7_res.get('total_amount') or step7_res.get('amount','-')} | "
            f"ภายใน: {step7_res.get('due','-')} | "
            f"PDF: {step7_res.get('pdf_file','-')}"
        )
    else:
        log(f"      [Step 7] ⚠ {step7_res.get('note','')}")
        result["status"] = step7_res.get("status") or "ALERT"
        result["note"] = (
            f"Step 1-6 เสร็จ + Step 7: {step7_res.get('note','')} | "
            f"Permit: {current_permit} | {identity_note}"
        )
    return result


def _save_bt44_report(rows: list[dict[str, Any]], out_path: Path, log=print) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "บต.44"
    headers = ["ลำดับ", "คำนำหน้า", "หมายเลขอ้างอิง", "ชื่อ", "สัญชาติ", "เพศ", "วันเกิด",
               "สถานะ",
               # ── ข้อมูลที่อ่านได้จาก Step 2.3 (โหมดทดลองยื่น) ──
               "ชื่อ (Eng)", "ชื่อ (ไทย)", "สัญชาติ (จากระบบ)",
               "เกิดวันที่ (ระบบ)", "อายุ", "โทรศัพท์", "อีเมล (ผู้ยื่น)",
               "ที่อยู่ในประเทศไทย",
               "เลขที่ใบอนุญาตปัจจุบัน", "จังหวัดที่ออก",
               "วันที่ออกเอกสาร", "วันที่เอกสารหมดอายุ",
               "ผลตรวจเลขใบอนุญาต",
               # ── ข้อมูลการชำระเงิน (โหมดยื่นจริง Step 7) ──
               "เลขที่ใบแจ้งชำระเงิน", "เลขที่คำขอ", "ชำระภายในวันที่",
               "จำนวนเงินรวม", "วันที่ทำรายการ",
               "รายชื่อคนต่างด้าว (ผู้ชำระ)", "อีเมล (ใบชำระเงิน)",
               "ไฟล์ PDF (ใบชำระเงิน)",
               "หมายเหตุ (ข้อความแจ้งเตือน)", "Screenshot"]
    ws.append(headers)
    head_fill = PatternFill("solid", fgColor="305496")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # คอลัมน์ที่ต้องทำเป็น hyperlink (1-based index)
    PDF_COL = 29       # ไฟล์ PDF (ใบชำระเงิน)
    SHOT_COL = 31      # Screenshot
    link_font = Font(color="0563C1", underline="single")

    def _to_hyperlink(cell, raw_path: str) -> None:
        """แปลง cell ให้คลิกแล้วเปิดไฟล์/โฟลเดอร์ได้ (Excel hyperlink)"""
        if not raw_path:
            return
        p = str(raw_path).strip()
        if not p:
            return
        try:
            abs_path = str(Path(p).resolve())
        except Exception:
            abs_path = p
        # Excel รองรับ file:/// บน Windows ดี + แสดงชื่อไฟล์เป็น display text
        try:
            display = Path(abs_path).name or abs_path
        except Exception:
            display = abs_path
        cell.value = display
        cell.hyperlink = abs_path
        cell.font = link_font

    for r in rows:
        ws.append([
            r.get("seq", ""), r.get("prefix", ""), r.get("alien_ref", ""), r.get("name", ""),
            r.get("nationality", ""), r.get("sex", ""), r.get("birthdate", ""),
            r.get("status", ""),
            r.get("dry_name_en", ""), r.get("dry_name_th", ""), r.get("dry_nationality", ""),
            r.get("dry_birthdate", ""), r.get("dry_age", ""),
            r.get("dry_phone", ""), r.get("dry_email", ""),
            r.get("dry_address", ""),
            r.get("dry_permit_no", ""), r.get("dry_permit_province", ""),
            r.get("dry_permit_issue", ""), r.get("dry_permit_expire", ""),
            r.get("dry_permit_match", ""),
            r.get("bill_no", ""), r.get("request_no", ""), r.get("due", ""),
            (r.get("total_amount") or r.get("amount", "")),
            r.get("txn_date", ""),
            r.get("biller_name", ""), r.get("email", ""),
            r.get("pdf_file", ""),
            r.get("note", ""), r.get("screenshot", ""),
        ])
        # แปลง path → hyperlink สำหรับแถวที่เพิ่งใส่
        row_idx = ws.max_row
        _to_hyperlink(ws.cell(row=row_idx, column=PDF_COL), r.get("pdf_file", ""))
        _to_hyperlink(ws.cell(row=row_idx, column=SHOT_COL), r.get("screenshot", ""))
    widths = [8, 12, 20, 28, 16, 8, 14, 12,
              28, 26, 16, 16, 8, 16, 28, 60,
              22, 18, 16, 16, 16,
              22, 18, 22, 14, 22, 28, 28, 36, 60, 28]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    log(f"บันทึกรายงาน: {out_path}")


def run_bt44(
    cfg: dict,
    excel_input: Path,
    login_excel: Path,
    out_path: Path,
    row_range: str | None = None,
    log=print,
    progress=None,
    is_cancelled=None,
    dry_run: bool = False,
    dry_stop_at: str = "step2",
    check_docs: bool = True,
) -> tuple[int, Path]:
    """โหมด บต.44 — การแจ้งการทำงาน และการยื่นคำขอเปลี่ยนรายการในใบอนุญาตทำงาน
    ซึ่งไม่กระทบในใบอนุญาต (Step 1-3)
    
    Step 1: ค้นหาข้อมูลคนต่างด้าว แล้วบันทึก
    Step 2.1: เช็ก checkbox ข้อมูลใจสำสัญญา แล้วกด "ถัดไป"
    Step 2.2: แก้ไขที่อยู่ + บันทึก + ยืนยัน
    Step 2.3: ตรวจสอบเลขที่ใบอนุญาตตรงกับ Excel ถ้าไม่ตรงให้ข้าม
    """
    out_path = _timestamped_path(out_path)
    records = _read_bt44_excel(excel_input)
    accounts = (_read_login_accounts(login_excel)
                if login_excel and Path(login_excel).exists() else {})
    total = len(records)
    indices = _parse_row_range(row_range, total)
    selected = [records[i - 1] for i in indices]
    log(f"[1/3] อ่าน {Path(excel_input).name}: {total} แถว → จะทำ {len(selected)} แถว "
        f"({row_range or 'ทั้งหมด'})")
    if dry_run:
        _stop_label = "Step 3.3-3.4 (เลือกนายจ้าง+ประเภทกิจการ+งาน)" if dry_stop_at == "step3" else "Step 2.3 (ตรวจเลขใบอนุญาต)"
        log(f"      🧪 โหมด 'ทดลองยื่น' (DRY-RUN) — จะหยุดที่ {_stop_label} ทุก record และไม่ส่งคำขอจริง")
    if not check_docs:
        log("      ⚠ ปิด 'ตรวจไฟล์แนบ' — จะข้ามการตรวจสอบไฟล์ทุก record (สำหรับทดสอบ flow เท่านั้น)")
    log(f"      ไฟล์รายงาน: {out_path.name}")

    # ตรวจไฟล์แนบ Step 4 ล่วงหน้า (สรุปแถวที่จะถูกข้าม) — ข้ามถ้าผู้ใช้ปิดไว้
    if check_docs:
        _bt44_preflight_docs(selected, log=log)
    else:
        log("      [Preflight] ⏭ ข้ามการตรวจไฟล์แนบล่วงหน้า")

    if accounts:
        acct = next(iter(accounts.values()))
        login_cfg = {
            "username": acct["username"], "password": acct["password"],
            "user_type": acct["type"],
            "method": acct.get("method") or cfg.get("method", "E-Workpermit"),
        }
    else:
        login_cfg = {k: cfg.get(k, "") for k in ("username", "password", "user_type", "method")}
    if not login_cfg.get("username") or not login_cfg.get("password"):
        raise ValueError("ไม่พบบัญชี login — กรุณาระบุ UsernameLogin.xlsx หรือกรอก Username/Password")
    log(f"      บัญชี login: {login_cfg['username']} ({login_cfg.get('user_type','')})")

    screenshot_dir = out_path.parent / "bt44_screenshots"
    screenshot_dir.mkdir(parents=True, exist_ok=True)

    results: list[dict[str, Any]] = []
    success = 0
    if progress:
        try: progress(0, len(selected))
        except Exception: pass

    with sync_playwright() as pw:
        browser = _launch_chromium(pw, cfg, ["--ignore-certificate-errors", "--start-maximized"])
        ctx = browser.new_context(
            locale="th-TH", ignore_https_errors=True,
            viewport={"width": 1920, "height": 1080},
        )
        page = ctx.new_page()
        try:
            log("[2/3] เข้าสู่ระบบ...")
            login(page, login_cfg)
            page.wait_for_timeout(1500)
            if not _open_bt44_form(page, log=log):
                log("[!] เปิดฟอร์ม บต.44 ไม่สำเร็จ — ยุติ")
                for rec in selected:
                    results.append({**rec, "status": "FORM_FAIL",
                                    "note": "เปิดฟอร์ม บต.44 ไม่สำเร็จ", "screenshot": ""})
                _save_bt44_report(results, out_path, log=log)
                return 0, out_path
            log("      ✓ เปิดฟอร์ม บต.44 สำเร็จ — เริ่มกรอกข้อมูลคนต่างด้าว")
            for k, rec in enumerate(selected, start=1):
                if is_cancelled and is_cancelled():
                    log("[!] ผู้ใช้ยกเลิก — หยุด")
                    break
                # ระหว่าง record ที่ 2 เป็นต้นไป — ต้อง reset กลับหน้า บต.44
                # (record แรก browser อยู่หน้า form แล้วจาก _open_bt44_form ข้างบน,
                # record ถัดมา page ค้างที่ Step 7 / payment / dry-run reset → ต้องเปิดฟอร์มใหม่)
                # หมายเหตุ: ไม่ต้อง logout/login ใหม่ — session ยังอยู่ใน browser context
                if k > 1:
                    log(f"      ↺ รีเซ็ตกลับหน้า บต.44 สำหรับ record ถัดไป...")
                    # เผื่อกรณี popup ใบชำระเงินจาก Step 7 ทำให้ main page ปิดไปด้วย
                    # → สร้าง page ใหม่จาก context (session/cookies ยังอยู่) แล้วเข้าใหม่
                    page_dead = False
                    try:
                        page_dead = page.is_closed()
                    except Exception:
                        page_dead = True
                    if page_dead:
                        log(f"      ⚠ Main page ถูกปิด (อาจจาก popup Step 7) — สร้าง page ใหม่")
                        try:
                            page = ctx.new_page()
                        except Exception as e:
                            log(f"      ✗ สร้าง page ใหม่ไม่สำเร็จ: {e}")
                            results.append({**rec, "status": "FORM_FAIL",
                                            "note": f"Browser context ปิด: {str(e)[:120]}",
                                            "screenshot": ""})
                            break
                    # ปิด popup/dialog ที่อาจค้าง (กันสถานะรกของหน้า)
                    try:
                        page.evaluate(r"""() => {
                            for (const m of document.querySelectorAll('.modal.show, .swal2-popup, [role="dialog"]')) {
                                try { m.style.display = 'none'; } catch(e){}
                            }
                            try { document.body.classList.remove('modal-open'); } catch(e){}
                            for (const bd of document.querySelectorAll('.modal-backdrop, .swal2-container')) {
                                try { bd.remove(); } catch(e){}
                            }
                        }""")
                    except Exception:
                        pass
                    if not _open_bt44_form(page, log=lambda *a: None):
                        log(f"      ⚠ เปิดฟอร์ม บต.44 ใหม่ไม่สำเร็จ — ลองทาง login ใหม่")
                        # ทางสำรอง: login ใหม่ในหน้า page เดิม (เผื่อ session timeout)
                        try:
                            login(page, login_cfg)
                            page.wait_for_timeout(1500)
                        except Exception as e:
                            log(f"      ✗ Login ซ้ำไม่สำเร็จ: {str(e)[:120]}")
                        if not _open_bt44_form(page, log=lambda *a: None):
                            log(f"      ⛔ เปิดฟอร์ม บต.44 ไม่ได้แม้ login ใหม่ — ข้าม record นี้")
                            results.append({**rec, "status": "FORM_FAIL",
                                            "note": "เปิดฟอร์ม บต.44 ใหม่ไม่สำเร็จ (ระหว่าง record)",
                                            "screenshot": ""})
                            if progress:
                                try: progress(k, len(selected))
                                except Exception: pass
                            continue
                    page.wait_for_timeout(800)
                log(f"  ({k}/{len(selected)}) แถว {rec['row_index']}: "
                    f"{rec.get('prefix','')} {rec.get('name','')} | "
                    f"{rec.get('nationality','')} | {rec.get('sex','')} | {rec.get('birthdate','')}"
                    + (f" | อ้างอิง {rec['alien_ref']}" if rec.get("alien_ref") else ""))
                res = _bt44_process_record(page, rec, screenshot_dir, log=log, dry_run=dry_run, dry_stop_at=dry_stop_at, check_docs=check_docs)
                row1 = {**rec, **res}
                results.append(row1)
                if res.get("status") in ("SUCCESS", "DRY_OK"):
                    success += 1
                log(f"      → {res.get('status')}" + (f" | {res['note']}" if res.get("note") else ""))
                if progress:
                    try: progress(k, len(selected))
                    except Exception: pass
                _save_bt44_report(results, out_path, log=lambda *a: None)
        finally:
            # หน่วงก่อนปิด browser เพื่อให้ตรวจสอบหน้าจอได้
            import os
            pause_sec = int(os.environ.get("BT44_PAUSE_BEFORE_CLOSE", "0"))
            if pause_sec > 0:
                log(f"      ⏸ หยุด {pause_sec} วินาทีก่อนปิด browser (กด Ctrl+C เพื่อยกเลิก)")
                try:
                    import time as _time
                    _time.sleep(pause_sec)
                except KeyboardInterrupt:
                    pass
            ctx.close(); browser.close()

    _save_bt44_report(results, out_path, log=log)
    log(f"[3/3] สรุป: สำเร็จ {success} / {len(selected)} รายการ "
        f"(ดู screenshots ใน {screenshot_dir.name}/)")
    return success, out_path


# ════════════════════════════════════════════════════════════════════════════
#  โหมด Bill Payment — ดาวน์โหลด 'ใบแจ้งชำระเงิน' ของคำขอที่สถานะ
#  'รอชำระเงิน → รอจ่ายเงินค่าธรรมเนียม' (ดาวน์โหลดอย่างเดียว ไม่ได้จ่ายเงินจริง:
#  ระบบสร้างใบแจ้งหนี้ + QR/เลขอ้างอิงให้ไปชำระผ่านแอปธนาคารภายหลัง)
# ════════════════════════════════════════════════════════════════════════════
_BILLPAY_TARGET_STATUS = "รอจ่ายเงินค่าธรรมเนียม"


def _billpay_parse_aliens(pdf_bytes: bytes) -> list[dict[str, str]]:
    """ดึงรายชื่อคนต่างด้าว + เลขอ้างอิง จากหน้า 'เอกสารแนบท้าย' ของใบแจ้งชำระเงิน
    แต่ละแถวอยู่ในรูป  '<ลำดับ>. <ชื่อคนต่างด้าว> <เลขอ้างอิง/เลขประจำตัว> <จำนวนเงิน>'
    คืน list ของ {name, ref} (กรองเฉพาะส่วนหลังหัวข้อ 'รายชื่อคนต่างด้าว / Name List')
    """
    out: list[dict[str, str]] = []
    if not pdf_bytes:
        return out
    try:
        import io
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(pdf_bytes))
        txt = "\n".join((pg.extract_text() or "") for pg in reader.pages)
    except Exception:
        return out
    idx = txt.find("Name List")
    if idx < 0:
        idx = txt.find("รายชื่อคนต่างด้าว")
    section = txt[idx:] if idx >= 0 else txt
    seen: set[tuple[str, str]] = set()
    for raw in section.splitlines():
        line = raw.strip()
        # ref = passport / RA-เลข / เลข 13 หลัก (มีตัวเลขอย่างน้อย 1 ตัว, ไม่มีช่องว่าง)
        m = re.match(r"^\d+\.\s+(.+?)\s+([A-Za-z]{0,3}\d[A-Za-z0-9]{5,})\s+[\d,]+\.\d{2}$", line)
        if not m:
            continue
        name = m.group(1).strip()
        ref = m.group(2).strip()
        if not re.search(r"[A-Za-zก-๙]", name):
            continue
        key = (name, ref)
        if key in seen:
            continue
        seen.add(key)
        out.append({"name": name, "ref": ref})
    return out


def _billpay_filename_stem(req_no: str, name: str, fallback: str = "") -> str:
    """ตั้งชื่อไฟล์ตามรูปแบบ {เลขที่คำขอ}_{รายชื่อคนต่างด้าว}"""
    parts = [p for p in (req_no, name) if str(p or "").strip()]
    stem = "_".join(_safe_filename(p) for p in parts)
    return stem or (_safe_filename(fallback) if str(fallback or "").strip() else "bill_payment")


def _process_one_bill_payment(
    page: Page, login_cfg: dict, row: dict[str, Any], bills_dir: Path, log=print,
) -> dict[str, Any]:
    """เปิดหน้า detail → แท็บ 'การชำระเงิน' → ปุ่ม 'ชำระเงิน' (ค่าธรรมเนียมขออนุญาตทำงาน)
    → modal → 'ชำระเงิน' → modal → 'พิมพ์แบบฟอร์มการชำระเงิน' → ดาวน์โหลด PDF + อ่านข้อมูล

    ขั้นตอนนี้ "สร้างใบแจ้งหนี้ + ดาวน์โหลด" เท่านั้น — ไม่ได้ตัดเงินจริง
    (เงินจะถูกชำระเมื่อผู้ใช้สแกน QR / ใช้เลขอ้างอิงผ่านแอปธนาคารภายหลัง)
    """
    req_no = str(row.get("group_id", "") or "")
    res: dict[str, Any] = {
        "req_no": req_no, "bill_no": "", "alien_name": "", "alien_ref": "",
        "amount": "", "due": "", "status": "", "pdf_file": "", "error": "",
        "statusText": row.get("statusText", ""), "requester": row.get("requester", ""),
    }

    # resume: ข้ามถ้ามีไฟล์ของเลขคำขอนี้อยู่แล้ว (ชื่อไฟล์ขึ้นต้นด้วย {req_no}_ หรือ {req_no}.pdf)
    try:
        existing = (sorted(bills_dir.glob(f"{req_no}_*.pdf")) + sorted(bills_dir.glob(f"{req_no}.pdf"))) if req_no else []
    except Exception:
        existing = []
    if existing:
        res["status"] = "SKIP_EXISTS"
        res["pdf_file"] = existing[0].name
        log(f"      ↷ {req_no} มีไฟล์แล้ว — ข้าม ({existing[0].name})")
        return res

    try:
        url = build_detail_url(row)
        page.goto(url, wait_until="domcontentloaded", timeout=40_000)
        page.wait_for_timeout(3500)
        # ฟื้น session ถ้าหลุด
        if _is_logged_out(page):
            log("      ⚠ session หมดอายุ — login ใหม่แล้วลองอีกครั้ง")
            login(page, login_cfg)
            page.goto(url, wait_until="domcontentloaded", timeout=40_000)
            page.wait_for_timeout(3500)
        # ตั้งภาษาไทย + ปิด popup ข่าว/โมดัลที่ค้าง
        try:
            page.locator("#sltLang").first.select_option(value="th")
            page.wait_for_timeout(800)
        except Exception:
            pass
        page.evaluate(
            r"""() => { if (window.jQuery){ try{ jQuery('.modal').modal('hide'); }catch(e){} }
              document.querySelectorAll('.modal.show .btn-close,.modal.show [data-bs-dismiss],.modal.show [data-dismiss]')
                .forEach(b => { try { b.click(); } catch(e){} }); }"""
        )
        page.wait_for_timeout(500)

        # เปิดแท็บ 'การชำระเงิน'
        page.evaluate(
            r"""() => { const a = Array.from(document.querySelectorAll('a[href^="#"]'))
                .find(e => /การชำระเงิน/.test((e.textContent||'').trim())); if (a) a.click(); }"""
        )
        page.wait_for_timeout(1500)

        # คลิกปุ่ม 'ชำระเงิน' (OpenModalPay1) ของแถว 'ค่าธรรมเนียมขออนุญาตทำงาน'
        clicked = page.evaluate(
            r"""() => {
              const vis = e => { const r = e.getBoundingClientRect(); return r.width>0 && r.height>0; };
              const btns = Array.from(document.querySelectorAll('button,a'))
                .filter(e => /OpenModalPay1/.test(e.getAttribute('onclick')||''));
              let pick = btns.find(b => {
                let txt = '', el = b;
                for (let i=0; i<5 && el; i++){ el = el.parentElement; if (el) txt += ' ' + (el.innerText||''); }
                return /ค่าธรรมเนียม/.test(txt) && vis(b);
              });
              if (!pick) pick = btns.find(vis) || btns[0];
              if (pick) { pick.click(); return true; }
              return false;
            }"""
        )
        if not clicked:
            res["status"] = "FAIL"
            res["error"] = "ไม่พบปุ่ม 'ชำระเงิน' (ค่าธรรมเนียมขออนุญาตทำงาน)"
            log(f"      ✗ {req_no}: {res['error']}")
            return res

        # modal 1 (#exampleModal) → ปุ่ม 'ชำระเงิน' (#openpaymentdetail = OpentWP)
        try:
            page.wait_for_selector("#exampleModal.show", timeout=8000)
        except Exception:
            pass
        page.wait_for_timeout(1000)
        page.evaluate(r"""() => { const b = document.getElementById('openpaymentdetail'); if (b) b.click(); }""")

        # modal 2 (#WP_Payment) แสดง QR + เลขอ้างอิง
        try:
            page.wait_for_selector("#WP_Payment.show", timeout=12_000)
        except Exception:
            pass
        page.wait_for_timeout(1500)
        refs = page.evaluate(
            r"""() => ({
              ref1:((document.getElementById('ref1')||{}).innerText||'').trim(),
              ref2:((document.getElementById('ref2')||{}).innerText||'').trim(),
              amount:((document.getElementById('lbl_price_payment_amount')||{}).innerText||'').trim(),
              due:((document.getElementById('payment_expired_time')||{}).innerText||'').trim(),
            })"""
        )

        # ปุ่ม 'พิมพ์แบบฟอร์มการชำระเงิน' (#button_payment = UpdatePayment) → เปิด PDF ใน tab ใหม่
        def _do_click():
            page.evaluate(r"""() => { const b = document.getElementById('button_payment'); if (b) b.click(); }""")

        body: bytes = b""
        err_last = ""
        for attempt in range(1, 4):
            b, err = _grab_pdf_after_click(page, _do_click, log=lambda *a: None)
            if b and len(b) >= 500 and b[:5] == b"%PDF-":
                body = b
                break
            err_last = err or "ไม่ได้ไฟล์"
            log(f"      · ⚠ {req_no} ดาวน์โหลดใบแจ้งชำระเงินครั้งที่ {attempt} ไม่สำเร็จ ({err_last}) — ลองใหม่")
            page.wait_for_timeout(2000)

        if not (body and body[:5] == b"%PDF-"):
            res["status"] = "FAIL"
            res["error"] = f"ดาวน์โหลด PDF ไม่สำเร็จ: {err_last}"
            res["bill_no"] = refs.get("ref2", "")
            res["amount"] = refs.get("amount", "")
            res["due"] = refs.get("due", "")
            log(f"      ✗ {req_no}: {res['error']}")
            return res

        # อ่านข้อมูลจาก PDF (แหล่งข้อมูลหลัก) + รายชื่อคนต่างด้าวจากหน้าแนบท้าย
        parsed = _bt30_parse_payment_pdf(body)
        aliens = _billpay_parse_aliens(body)
        bill_no = parsed.get("bill_no") or refs.get("ref2", "")
        req_no_final = parsed.get("request_no") or req_no
        if aliens:
            name = aliens[0]["name"]
            ref = aliens[0]["ref"]
            res["alien_name"] = "; ".join(a["name"] for a in aliens)
            res["alien_ref"] = "; ".join(a["ref"] for a in aliens)
        else:
            name = ""
            ref = parsed.get("alien_ref", "")
            res["alien_name"] = ""
            res["alien_ref"] = ref

        stem = _billpay_filename_stem(req_no_final, name, fallback=req_no)
        fname = f"{stem}.pdf"
        (bills_dir / fname).write_bytes(body)

        res["status"] = "SUCCESS"
        res["bill_no"] = bill_no
        res["amount"] = parsed.get("amount") or refs.get("amount", "")
        res["due"] = parsed.get("due") or refs.get("due", "")
        res["pdf_file"] = fname
        extra = f" (+{len(aliens) - 1} คน)" if len(aliens) > 1 else ""
        log(f"      ✓ {req_no}: ใบแจ้งชำระเงิน {bill_no} | {name}{extra} | ยอด {res['amount']} → {fname}")
        return res

    except Exception as e:
        res["status"] = "ERROR"
        res["error"] = str(e).splitlines()[0][:200]
        log(f"      ✗ {req_no}: ผิดพลาด — {res['error']}")
        return res


def _save_billpay_report(rows: list[dict[str, Any]], out_path: Path, log=print) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "ใบแจ้งชำระเงิน"
    headers = ["ลำดับ", "Username", "เลขที่คำขอ", "เลขที่ใบแจ้งชำระเงิน",
               "รายชื่อคนต่างด้าว", "เลขอ้างอิง/เลขประจำตัว", "ยอดชำระ", "ชำระภายใน",
               "สถานะคำขอ", "Status", "ไฟล์", "Error"]
    ws.append(headers)
    head_fill = PatternFill("solid", fgColor="305496")
    link_font = Font(color="0563C1", underline="single")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    text_cols = {"เลขที่คำขอ", "เลขที่ใบแจ้งชำระเงิน", "เลขอ้างอิง/เลขประจำตัว"}
    for i, r in enumerate(rows, start=1):
        ws.append([
            i, r.get("username", ""), r.get("req_no", ""), r.get("bill_no", ""),
            r.get("alien_name", ""), r.get("alien_ref", ""), r.get("amount", ""),
            r.get("due", ""), r.get("statusText", ""), r.get("status", ""),
            r.get("pdf_file", ""), r.get("error", ""),
        ])
        row_idx = ws.max_row
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=row_idx, column=c_idx)
            if h in text_cols and cell.value not in (None, ""):
                cell.value = str(cell.value)
                cell.number_format = "@"
            if h == "ไฟล์" and cell.value:
                fname = str(cell.value)
                cell.value = f'=HYPERLINK("bill_payment/{fname}","{fname}")'
                cell.font = link_font

    widths = [8, 30, 20, 22, 30, 24, 12, 20, 28, 12, 40, 45]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    log(f"บันทึกรายงาน: {out_path}")


def run_bill_payment(
    cfg: dict,
    login_excel: Path,
    out_path: Path,
    row_range: str | None = None,
    request_types: list[str] | None = None,
    log=print,
    progress=None,
    is_cancelled=None,
) -> tuple[int, Path]:
    """ดาวน์โหลด 'ใบแจ้งชำระเงิน' ของคำขอที่สถานะ 'รอจ่ายเงินค่าธรรมเนียม' (ทุกบัญชีใน login_excel)

    - วนทุก Username ใน UsernameLogin.xlsx → login → เปิด e-Tracking → filter 'รอชำระเงิน' (WP)
      → เลือกเฉพาะแถว 'รอจ่ายเงินค่าธรรมเนียม'
    - request_types: รหัส 'รายการคำขอ' ที่จะกรอง (เลือกได้หลายรายการ) — None/ว่าง = ทุกรายการคำขอ
    - แต่ละคำขอ: เข้าหน้า detail → แท็บการชำระเงิน → กดชำระเงิน → modal → ชำระเงิน → modal
      → 'พิมพ์แบบฟอร์มการชำระเงิน' → ดาวน์โหลด PDF (สร้างใบแจ้งหนี้ ไม่ได้ตัดเงินจริง)
    - ตั้งชื่อไฟล์ {เลขที่คำขอ}_{รายชื่อคนต่างด้าว}
    - resume ได้: ข้ามคำขอที่มีไฟล์อยู่แล้ว
    """
    out_path = _timestamped_path(out_path)
    accounts = _read_login_accounts(login_excel)
    # เอาเฉพาะรหัสที่ไม่ว่าง (ตัด '' = ทั้งหมด ออก) — ถ้าว่างทั้งหมดถือว่าไม่กรอง
    req_types = [rt for rt in (request_types or []) if rt]
    log(f"[1/3] บัญชี login: {login_excel} ({len(accounts)} บัญชี)")
    log(f"      เป้าหมาย: คำขอสถานะ 'รอชำระเงิน → {_BILLPAY_TARGET_STATUS}'")
    if req_types:
        log(f"      กรองเฉพาะรายการคำขอ: {', '.join(req_types)} ({len(req_types)} รายการ)")
    else:
        log("      รายการคำขอ: ทั้งหมด (ไม่กรอง)")

    bills_dir = out_path.parent / "bill_payment"
    bills_dir.mkdir(parents=True, exist_ok=True)

    results: list[dict[str, Any]] = []
    success = 0
    total_seen = 0

    with sync_playwright() as pw:
        browser = _launch_chromium(pw, cfg, ["--ignore-certificate-errors", "--start-maximized"])
        try:
            for ai, (ukey, acct) in enumerate(accounts.items(), start=1):
                if is_cancelled and is_cancelled():
                    log("[!] ผู้ใช้ยกเลิก — หยุด")
                    break
                login_cfg = {
                    "username": acct["username"],
                    "password": acct["password"],
                    "user_type": acct["type"],
                    "method": acct.get("method") or cfg.get("method", "E-Workpermit"),
                }
                ctx = browser.new_context(
                    locale="th-TH",
                    ignore_https_errors=True,
                    viewport={"width": 1920, "height": 1080},
                    accept_downloads=True,
                )
                page = ctx.new_page()
                try:
                    log(f"[บัญชี {ai}/{len(accounts)}] เข้าสู่ระบบ: {acct['username']} ({acct['type']})")
                    try:
                        login(page, login_cfg)
                    except Exception as e:
                        log(f"      ✗ login ไม่สำเร็จ: {e} — ข้ามบัญชีนี้")
                        results.append({"username": acct["username"], "req_no": "", "bill_no": "",
                                        "alien_name": "", "alien_ref": "", "amount": "", "due": "",
                                        "statusText": "", "status": "LOGIN_FAIL", "pdf_file": "",
                                        "error": str(e).splitlines()[0][:200]})
                        continue

                    goto_tracking(page)
                    if req_types:
                        # กรองทีละรายการคำขอ (ใช้ filter ของเว็บเอง) แล้วรวมผล + ตัดซ้ำตาม group_id
                        targets = []
                        seen_ids: set[str] = set()
                        for rt in req_types:
                            apply_wa_filter(page, rt, status_ids=["WP"])
                            page.wait_for_timeout(1200)
                            rows = collect_all_wa_rows(page, log=log)
                            for r in rows:
                                gid = r.get("group_id")
                                if (_BILLPAY_TARGET_STATUS in (r.get("statusText", "") or "")
                                        and gid not in seen_ids):
                                    seen_ids.add(gid)
                                    targets.append(r)
                    else:
                        apply_wa_filter(page, "", status_ids=["WP"])
                        page.wait_for_timeout(1500)
                        rows = collect_all_wa_rows(page, log=log)
                        targets = [r for r in rows if _BILLPAY_TARGET_STATUS in (r.get("statusText", "") or "")]
                    n_idx = _parse_row_range(row_range, len(targets))
                    selected = [targets[i - 1] for i in n_idx]
                    total_seen += len(selected)
                    log(f"      พบ {len(targets)} คำขอรอจ่ายค่าธรรมเนียม → จะทำ {len(selected)} "
                        f"({row_range or 'ทั้งหมด'})")

                    for ri, row in enumerate(selected, start=1):
                        if is_cancelled and is_cancelled():
                            log("[!] ผู้ใช้ยกเลิก — หยุด")
                            break
                        log(f"  ({ri}/{len(selected)}) คำขอ {row.get('group_id','')} | {row.get('requester','')}")
                        res = _process_one_bill_payment(page, login_cfg, row, bills_dir, log=log)
                        res["username"] = acct["username"]
                        results.append(res)
                        if res.get("status") == "SUCCESS":
                            success += 1
                        if progress:
                            try: progress(len(results), total_seen)
                            except Exception: pass
                except Exception as e:
                    log(f"      ✗ บัญชี {acct['username']} ผิดพลาด: {str(e).splitlines()[0][:200]}")
                finally:
                    ctx.close()
                _save_billpay_report(results, out_path, log=lambda *_: None)
        finally:
            browser.close()

    _save_billpay_report(results, out_path, log=log)
    fail = sum(1 for r in results if r.get("status") not in ("SUCCESS", "SKIP_EXISTS"))
    log(f"[3/3] สรุป: ดาวน์โหลดสำเร็จ {success} / {len(results)} (ล้มเหลว {fail})")
    return success, out_path


# ════════════════════════════════════════════════════════════════════════════
#  โหมด Payment Receipts — ดาวน์โหลด 'ใบเสร็จรับเงิน' ทั้งชุดของคำขอ
#  (ใบเสร็จ 900 / 100 และอื่นๆ) จากหน้า /Center/MultiplePayments
#  - แต่ละไฟล์ดาวน์โหลดจากระบบมีหลายคน → แยกหน้าเป็น 1 ใบเสร็จ/1 คน/1 ไฟล์
#  - ชื่อไฟล์: {เลขที่คำขอ}_{ชื่อคนต่างด้าว}_{ชื่อนายจ้าง}_{Amount}_{เลขอ้างอิงคนต่างด้าว}.pdf
# ════════════════════════════════════════════════════════════════════════════

def _read_request_receipt_excel(path: Path) -> list[dict[str, Any]]:
    """อ่าน Request_Receipt.xlsx → [{seq, req_no, username, row_index}]
    คอลัมน์ (ยืดหยุ่น): ลำดับ | เลขที่คำขอ | Username
    """
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    hdr = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    def col(*names: str) -> int:
        for nm in names:
            for i, h in enumerate(hdr):
                if h == nm:
                    return i
        for nm in names:
            for i, h in enumerate(hdr):
                if nm.lower() in h.lower():
                    return i
        return -1

    i_seq = col("ลำดับ", "No.", "Seq")
    i_req = col("เลขที่คำขอ", "Request No.", "RequestNo", "request_no", "group_id")
    i_user = col("Username", "username", "อีเมล")

    out: list[dict[str, Any]] = []
    for r_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not any(v not in (None, "") for v in row):
            continue

        def g(i: int) -> str:
            return str(row[i]).strip() if 0 <= i < len(row) and row[i] is not None else ""

        req = "".join(ch for ch in g(i_req) if ch.isdigit())
        user = g(i_user)
        if not req or not user:
            continue
        out.append({
            "seq": g(i_seq) or str(r_idx - 1),
            "req_no": req,
            "username": user,
            "row_index": r_idx,
        })
    return out


def _payreceipt_parse_page(text: str) -> dict[str, str]:
    """แยกข้อมูล 1 หน้าใบเสร็จ (1 ใบ / 1 คน) จาก text ของ PDF page
    Labels ภาษาอังกฤษเป็น anchor หลัก (เสถียร) คืน dict ของ field ที่อ่านได้
    """
    out = {
        "request_no": "", "receipt_no": "", "bill_payment_no": "",
        "foreigner_name": "", "foreigner_ref": "", "nationality": "",
        "employer_name": "", "employer_id": "",
        "amount": "", "payment_date": "", "biller_name": "",
    }
    if not text:
        return out
    import re as _re

    def g(pat: str, flags: int = 0) -> str:
        m = _re.search(pat, text, flags)
        return m.group(1).strip() if m else ""

    # label → next non-blank line (ค่าจริงอยู่ถัดจาก label English ที่อยู่ในวงเล็บ)
    def after_label(label: str) -> str:
        # หา label ที่อยู่บนบรรทัด — แล้วคืนบรรทัดถัดไปที่ไม่ว่าง
        lines = text.split("\n")
        for i, ln in enumerate(lines):
            if label in ln:
                for j in range(i + 1, min(i + 6, len(lines))):
                    v = lines[j].strip()
                    if v and "(" not in v[:3]:  # ข้ามบรรทัด label ถัดไป
                        return v
                break
        return ""

    out["foreigner_name"] = after_label("(Foreigner Name)").rstrip()
    out["nationality"] = after_label("(Nationality)")
    out["foreigner_ref"] = after_label("(Foreigner Reference no./ Foreigner ID)")
    out["payment_date"] = after_label("(Payment Date)")
    out["receipt_no"] = after_label("(Receipt No.)")
    out["request_no"] = after_label("(Request No.)")
    out["bill_payment_no"] = after_label("(Bill Payment No.)")
    out["employer_id"] = after_label("(Employer ID)")
    out["biller_name"] = after_label("(Biller Name)")

    # Employer Name อาจมีหลายบรรทัด (ตัวอย่างมี linebreak กลางชื่อ) — รวม 1-3 บรรทัดถัดจาก label
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        if "(Employer Name/Company Name)" in ln:
            parts: list[str] = []
            for j in range(i + 1, min(i + 5, len(lines))):
                v = lines[j].strip()
                if not v:
                    continue
                if v.startswith("(") and ")" in v[:25]:
                    break
                if v.startswith("เลขประจำตัว") or "(Employer ID)" in v:
                    break
                parts.append(v)
            out["employer_name"] = " ".join(parts).strip()
            break

    # Amount — รูปแบบในใบเสร็จ: "...(Amount)\n900.00\n..." (อยู่หลัง label Amount ที่เป็นยอดรวม)
    # หา occurrence สุดท้ายของ "(Amount)" (ตำแหน่งของยอดรวม) แล้วคืนค่าตัวเลข
    m = list(_re.finditer(r"\(Amount\)\s*\n([0-9,]+\.\d{2})", text))
    if m:
        out["amount"] = m[-1].group(1)
    else:
        # fallback: บรรทัด "900.00 1 900.00" → เอาตัวสุดท้าย
        m2 = _re.search(r"([0-9,]+\.\d{2})\s+\d+\s+([0-9,]+\.\d{2})", text)
        if m2:
            out["amount"] = m2.group(2)
    return out


def _payreceipt_split_pages(body: bytes) -> list[tuple[bytes, str]]:
    """แยก PDF หลายหน้า → [(single_page_pdf_bytes, page_text), ...]"""
    out: list[tuple[bytes, str]] = []
    if not body or body[:5] != b"%PDF-":
        return out
    try:
        import io
        from pypdf import PdfReader, PdfWriter
        reader = PdfReader(io.BytesIO(body))
        for pg in reader.pages:
            writer = PdfWriter()
            writer.add_page(pg)
            buf = io.BytesIO()
            writer.write(buf)
            out.append((buf.getvalue(), pg.extract_text() or ""))
    except Exception:
        return out
    return out


def _payreceipt_filename(req_no: str, info: dict[str, str], page_idx: int) -> str:
    """สร้างชื่อไฟล์ {request_no}_{foreigner_name}_{employer_name}_{foreigner_ref}_RECEIPT{amount_int}.pdf
    Amount เป็นจำนวนเต็ม (ตัดทศนิยม) เช่น 900.00 → 900
    """
    req = info.get("request_no") or req_no
    name = info.get("foreigner_name") or ""
    emp = info.get("employer_name") or ""
    ref = info.get("foreigner_ref") or ""
    amt_raw = (info.get("amount") or "").replace(",", "").strip()
    amt_int = ""
    if amt_raw:
        try:
            amt_int = str(int(float(amt_raw)))
        except Exception:
            amt_int = amt_raw.split(".")[0]
    parts = [p for p in (req, name, emp, ref) if str(p).strip()]
    stem = "_".join(_safe_filename(p) for p in parts)
    if not stem:
        stem = f"{_safe_filename(req_no)}_page{page_idx}"
    suffix = f"_RECEIPT{amt_int}" if amt_int else "_RECEIPT"
    return f"{stem}{suffix}.pdf"


def _payreceipt_open_list_page(page: Page, req_no: str, log=print) -> tuple[bool, str]:
    """เปิดหน้า /Center/MultiplePayments ของคำขอ — return (ok, error)
    flow: Tracking → ค้นหา req_no → คลิกแถวแรก → tab การชำระเงิน → ปุ่ม ดูใบเสร็จรับเงิน
    """
    try:
        page.goto(TRACKING_URL, wait_until="domcontentloaded", timeout=40_000)
        page.wait_for_timeout(2500)
        if _is_logged_out(page):
            return False, "session หมดอายุ"
        _search_request(page, req_no)
        page.wait_for_timeout(1500)

        # ตรวจว่ามีผลค้นหา
        has = page.evaluate(
            r"""() => !!document.querySelector('a[onclick*="openDetail"]')"""
        )
        if not has:
            return False, "ไม่พบเลขคำขอในระบบ"

        # คลิกเข้า detail
        try:
            with page.expect_navigation(timeout=20_000, wait_until="domcontentloaded"):
                page.evaluate(
                    r"""() => { const a = document.querySelector('a[onclick*="openDetail"]'); if (a) a.click(); }"""
                )
        except Exception:
            pass
        page.wait_for_timeout(2500)

        # เปลี่ยนภาษา ปิด popup
        try:
            page.locator("#sltLang").first.select_option(value="th")
            page.wait_for_timeout(600)
        except Exception:
            pass
        page.evaluate(
            r"""() => {
            if (window.jQuery) { try { jQuery('.modal').modal('hide'); } catch(e){} }
            document.querySelectorAll('.modal.show .btn-close,.modal.show [data-bs-dismiss]')
                .forEach(b => { try { b.click(); } catch(e){} });
        }"""
        )
        page.wait_for_timeout(500)

        # เปิดแท็บ การชำระเงิน
        page.evaluate(
            r"""() => {
            const a = Array.from(document.querySelectorAll('a[href^="#"], .nav-link, .nav a'))
                .find(e => /^การชำระเงิน$/.test((e.textContent||'').trim()));
            if (a) a.click();
        }"""
        )
        page.wait_for_timeout(1800)

        # คลิก 'ดูใบเสร็จรับเงิน' — มันเป็น window.location.href navigation (ไม่ใช่ popup)
        try:
            with page.expect_navigation(timeout=20_000, wait_until="domcontentloaded"):
                page.evaluate(
                    r"""() => {
                    const vis = el => { const r = el.getBoundingClientRect(); return r.width>0 && r.height>0; };
                    const b = Array.from(document.querySelectorAll('button,a'))
                        .find(e => vis(e) && /ดูใบเสร็จรับเงิน/.test((e.textContent||'').trim()));
                    if (b) b.click();
                }"""
                )
        except Exception:
            pass
        page.wait_for_timeout(3500)

        if "/Center/MultiplePayments" not in (page.url or ""):
            return False, f"ไม่ได้เข้าหน้า MultiplePayments (อยู่ที่ {page.url[:80]})"
        return True, ""
    except Exception as e:
        return False, str(e).splitlines()[0][:200]


def _payreceipt_total_rows(page: Page) -> int:
    """อ่านจำนวนรายการทั้งหมดจาก 'รายการคำขอ N รายการ'"""
    try:
        return page.evaluate(
            r"""() => {
            const m = (document.body.innerText || '').match(/รายการคำขอ\s+(\d+)\s+รายการ/);
            return m ? parseInt(m[1], 10) : 0;
        }"""
        )
    except Exception:
        return 0


def _payreceipt_set_page_size(page: Page, size: int = 100) -> bool:
    """พยายามเปลี่ยน page size dropdown เป็นค่าสูงสุด (ลด pagination)"""
    try:
        return bool(page.evaluate(
            r"""(sz) => {
            const selects = Array.from(document.querySelectorAll('select'));
            for (const s of selects) {
                const opts = Array.from(s.options).map(o => parseInt(o.value || o.textContent, 10) || 0);
                if (opts.some(v => v === 10) && opts.some(v => v >= 20)) {
                    // เป็น dropdown 'รายการต่อหน้า'
                    const target = opts.filter(v => v >= sz).sort((a,b)=>a-b)[0]
                                || Math.max(...opts);
                    s.value = String(target);
                    s.dispatchEvent(new Event('change', {bubbles: true}));
                    return true;
                }
            }
            return false;
        }""", size))
    except Exception:
        return False


def _payreceipt_go_next_page(page: Page) -> bool:
    """กดปุ่ม 'ถัดไป' / Next ใน pagination — คืน True ถ้ากดสำเร็จและเปลี่ยนหน้าได้"""
    try:
        return bool(page.evaluate(
            r"""() => {
            const vis = el => { const r = el.getBoundingClientRect(); const s = getComputedStyle(el);
                return r.width>0 && r.height>0 && s.display !== 'none' && s.visibility !== 'hidden'; };
            const btns = Array.from(document.querySelectorAll('a, button, li'));
            // ปุ่มถัดไป (ไม่ถูก disable)
            const next = btns.find(b => vis(b)
                && /^(ถัดไป|next|»|›)/i.test((b.textContent||'').trim())
                && !b.disabled
                && !/disabled/i.test(b.className || '')
                && !(b.parentElement && /disabled/i.test(b.parentElement.className || '')));
            if (next) {
                const a = next.tagName === 'LI' ? next.querySelector('a, button') : next;
                if (a) { a.click(); return true; }
            }
            return false;
        }"""
        ))
    except Exception:
        return False


def _payreceipt_iter_and_download(
    page: Page, req_no: str, save_dir: Path, log=print,
) -> list[dict[str, Any]]:
    """หน้า /Center/MultiplePayments: วนทุกแถว ทุก page → คลิก 'ใบเสร็จรับเงิน' → grab PDF → split per page → save
    คืน list ของ result records
    """
    save_dir.mkdir(parents=True, exist_ok=True)
    out_records: list[dict[str, Any]] = []

    page.wait_for_timeout(1500)
    total = _payreceipt_total_rows(page)
    log(f"      · พบ {total or '?'} รายการใบแจ้งชำระเงิน")
    _payreceipt_set_page_size(page, 100)
    page.wait_for_timeout(2000)

    processed_rows = 0
    page_idx = 0
    while True:
        page_idx += 1
        # หาแถวทั้งหมดในหน้านี้ — ใช้ตำแหน่ง bill_no เป็นกุญแจกัน duplicate
        row_count = page.evaluate(
            r"""() => {
            const trs = Array.from(document.querySelectorAll('table tbody tr'));
            return trs.filter(tr => tr.querySelector('button, a')
                && /ใบเสร็จรับเงิน/.test(tr.innerText || '')).length;
        }"""
        )
        log(f"      · หน้า {page_idx}: {row_count} แถวที่มีปุ่มใบเสร็จรับเงิน")
        if not row_count:
            break

        for i in range(row_count):
            # ดึงข้อมูลแถวก่อนคลิก (เลขใบแจ้ง, สถานะ)
            row_info = page.evaluate(
                r"""(idx) => {
                const trs = Array.from(document.querySelectorAll('table tbody tr'))
                    .filter(tr => /ใบเสร็จรับเงิน/.test(tr.innerText || ''));
                const tr = trs[idx];
                if (!tr) return null;
                const tds = tr.querySelectorAll('td');
                const txt = (n) => tds[n] ? (tds[n].innerText||'').trim() : '';
                return {
                    seq: txt(0),
                    bill_no: (txt(1).match(/\d{10,}/) || [''])[0],
                    due: txt(2),
                    paid_date: txt(3),
                    status: txt(4),
                };
            }""", i,
            )
            if not row_info:
                continue
            bill_no = row_info.get("bill_no", "")
            log(f"        ({i + 1}/{row_count}) ใบแจ้ง {bill_no} | {row_info.get('status','-')} | จ่าย {row_info.get('paid_date','-')}")

            # คลิกปุ่มใบเสร็จรับเงินของแถวนี้
            def _do_click(idx: int = i):
                page.evaluate(
                    r"""(idx) => {
                    const trs = Array.from(document.querySelectorAll('table tbody tr'))
                        .filter(tr => /ใบเสร็จรับเงิน/.test(tr.innerText || ''));
                    const tr = trs[idx];
                    if (!tr) return;
                    const btn = Array.from(tr.querySelectorAll('button, a'))
                        .find(b => /ใบเสร็จรับเงิน/.test((b.textContent||'').trim()));
                    if (btn) btn.click();
                }""", idx,
                )

            body: bytes = b""
            err_last = ""
            for attempt in range(1, 4):
                b, err = _grab_pdf_after_click(page, _do_click, log=lambda *a: None)
                if b and len(b) >= 500 and b[:5] == b"%PDF-":
                    body = b
                    break
                err_last = err or "ไม่ได้ไฟล์"
                log(f"          · ⚠ ดาวน์โหลดครั้งที่ {attempt} ไม่สำเร็จ ({err_last})")
                page.wait_for_timeout(1500)

            if not body:
                out_records.append({
                    "req_no": req_no, "bill_no": bill_no, "status": "FAIL",
                    "error": f"ดาวน์โหลดล้มเหลว: {err_last}",
                    **{k: row_info.get(k, "") for k in ("paid_date", "due")},
                })
                processed_rows += 1
                continue

            # split per page → save
            pages = _payreceipt_split_pages(body)
            log(f"          · ดาวน์โหลด PDF {len(body)//1024} KB ({len(pages)} หน้า) — แยกเป็น 1 ใบ/1 คน")
            for p_idx, (single_body, txt) in enumerate(pages, start=1):
                info = _payreceipt_parse_page(txt)
                fname = _payreceipt_filename(req_no, info, p_idx)
                # resume: ข้ามถ้ามีไฟล์อยู่แล้ว
                if (save_dir / fname).exists():
                    log(f"            ↷ {fname} (มีอยู่แล้ว — ข้าม)")
                else:
                    try:
                        (save_dir / fname).write_bytes(single_body)
                        log(f"            ✓ {fname} ({len(single_body)//1024} KB)")
                    except Exception as e:
                        log(f"            ✗ เขียนไฟล์ไม่สำเร็จ: {str(e)[:120]}")
                out_records.append({
                    "req_no": req_no, "bill_no": bill_no,
                    "status": "SUCCESS",
                    "page": p_idx, "pages_total": len(pages),
                    "foreigner_name": info.get("foreigner_name", ""),
                    "foreigner_ref": info.get("foreigner_ref", ""),
                    "nationality": info.get("nationality", ""),
                    "employer_name": info.get("employer_name", ""),
                    "amount": info.get("amount", ""),
                    "payment_date": info.get("payment_date", "") or row_info.get("paid_date", ""),
                    "receipt_no": info.get("receipt_no", ""),
                    "bill_payment_no": info.get("bill_payment_no", "") or bill_no,
                    "pdf_file": fname,
                    "paid_date": row_info.get("paid_date", ""),
                    "due": row_info.get("due", ""),
                    "row_status": row_info.get("status", ""),
                })
            processed_rows += 1
            page.wait_for_timeout(400)

        # ไปหน้าถัดไป
        if not _payreceipt_go_next_page(page):
            break
        page.wait_for_timeout(2000)

    log(f"      · รวมประมวลผล {processed_rows} แถว → {len(out_records)} ใบเสร็จ")
    return out_records


def _save_payreceipt_report(rows: list[dict[str, Any]], out_path: Path, log=print) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "ใบเสร็จรับเงิน"
    headers = ["ลำดับ", "Username", "เลขที่คำขอ",
               "เลขที่ใบแจ้งชำระเงิน", "เลขที่ใบเสร็จ",
               "ชื่อคนต่างด้าว", "สัญชาติ", "เลขอ้างอิงคนต่างด้าว",
               "ชื่อนายจ้าง/สถานประกอบการ", "ยอด (Amount)",
               "วันที่ชำระเงิน", "กำหนดชำระ", "สถานะ", "หน้าที่",
               "ไฟล์ PDF", "Error"]
    ws.append(headers)
    head_fill = PatternFill("solid", fgColor="305496")
    link_font = Font(color="0563C1", underline="single")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    text_cols = {"เลขที่คำขอ", "เลขที่ใบแจ้งชำระเงิน", "เลขที่ใบเสร็จ", "เลขอ้างอิงคนต่างด้าว"}
    for i, r in enumerate(rows, start=1):
        ws.append([
            i, r.get("username", ""), r.get("req_no", ""),
            r.get("bill_payment_no", "") or r.get("bill_no", ""),
            r.get("receipt_no", ""),
            r.get("foreigner_name", ""), r.get("nationality", ""), r.get("foreigner_ref", ""),
            r.get("employer_name", ""), r.get("amount", ""),
            r.get("payment_date", "") or r.get("paid_date", ""), r.get("due", ""),
            r.get("status", ""),
            (f"{r.get('page','')}/{r.get('pages_total','')}" if r.get("page") else ""),
            r.get("pdf_file", ""), r.get("error", ""),
        ])
        row_idx = ws.max_row
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=row_idx, column=c_idx)
            if h in text_cols and cell.value not in (None, ""):
                cell.value = str(cell.value)
                cell.number_format = "@"
            if h == "ไฟล์ PDF" and cell.value:
                fname = str(cell.value)
                cell.value = f'=HYPERLINK("payment_receipts/{fname}","{fname}")'
                cell.font = link_font

    widths = [6, 30, 20, 22, 20, 28, 14, 24, 36, 12, 22, 22, 12, 10, 60, 40]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    log(f"บันทึกรายงาน: {out_path}")


def run_payment_receipts(
    cfg: dict,
    request_excel: Path,
    login_excel: Path,
    out_path: Path,
    row_range: str | None = None,
    log=print,
    progress=None,
    is_cancelled=None,
) -> tuple[int, Path]:
    """ดาวน์โหลด 'ใบเสร็จรับเงิน' ทั้งชุด (900 + 100 + อื่นๆ) ตามเลขคำขอใน Request_Receipt.xlsx
    - Loop ทีละ Username (sequential) — login/logout เปลี่ยน user อัตโนมัติ
    - แต่ละคำขอ: เปิด /Center/MultiplePayments → วนทุกแถว ทุก page → ดาวน์โหลด → split per page → ตั้งชื่อ
    - resume: ข้ามไฟล์ที่มีอยู่แล้ว
    """
    out_path = _timestamped_path(out_path)
    records = _read_request_receipt_excel(request_excel)
    accounts = _read_login_accounts(login_excel)
    total = len(records)
    indices = _parse_row_range(row_range, total)
    selected = [records[i - 1] for i in indices]
    log(f"[1/3] อ่าน {Path(request_excel).name}: {total} แถว → จะทำ {len(selected)} แถว "
        f"({row_range or 'ทั้งหมด'})")
    log(f"      บัญชี login จาก {Path(login_excel).name}: {len(accounts)} บัญชี")

    save_dir = out_path.parent / "payment_receipts"
    save_dir.mkdir(parents=True, exist_ok=True)

    # จัดกลุ่มตาม Username (รักษาลำดับ)
    by_user: dict[str, list[dict[str, Any]]] = {}
    for rec in selected:
        by_user.setdefault(rec["username"], []).append(rec)

    results: list[dict[str, Any]] = []
    success = 0
    done_rows = 0

    with sync_playwright() as pw:
        browser = _launch_chromium(pw, cfg, ["--ignore-certificate-errors", "--start-maximized"])
        ctx = browser.new_context(
            locale="th-TH", ignore_https_errors=True,
            viewport={"width": 1920, "height": 1080},
        )
        page = ctx.new_page()
        try:
            for u_idx, (username, recs) in enumerate(by_user.items(), start=1):
                if is_cancelled and is_cancelled():
                    log("[!] ผู้ใช้ยกเลิก — หยุด")
                    break
                acct = accounts.get(username.lower())
                if not acct:
                    log(f"  [{u_idx}/{len(by_user)}] ⛔ ไม่พบ {username} ใน UsernameLogin.xlsx — ข้าม "
                        f"({len(recs)} คำขอ)")
                    for r in recs:
                        results.append({**r, "status": "NO_LOGIN",
                                        "error": "ไม่พบบัญชีใน UsernameLogin.xlsx",
                                        "username": username})
                        done_rows += 1
                    continue

                log(f"  [{u_idx}/{len(by_user)}] === เปลี่ยน user: {username} ({len(recs)} คำขอ) ===")
                _logout_safely(page)
                page.wait_for_timeout(1500)
                try:
                    login(page, {
                        "username": acct["username"], "password": acct["password"],
                        "user_type": acct["type"],
                        "method": acct.get("method") or cfg.get("method", "E-Workpermit"),
                    })
                except Exception as e:
                    log(f"      ⛔ login ล้มเหลว: {str(e)[:160]}")
                    for r in recs:
                        results.append({**r, "status": "LOGIN_FAIL",
                                        "error": str(e)[:160], "username": username})
                        done_rows += 1
                    continue
                page.wait_for_timeout(1500)

                for k, rec in enumerate(recs, start=1):
                    if is_cancelled and is_cancelled():
                        break
                    req_no = rec["req_no"]
                    log(f"    ({k}/{len(recs)}) คำขอ {req_no}")
                    try:
                        ok, err = _payreceipt_open_list_page(page, req_no, log=log)
                        if not ok:
                            log(f"      ⛔ {err}")
                            results.append({**rec, "status": "FAIL", "error": err,
                                            "username": username})
                            done_rows += 1
                            continue
                        recs_one = _payreceipt_iter_and_download(page, req_no, save_dir, log=log)
                        if not recs_one:
                            results.append({**rec, "status": "EMPTY",
                                            "error": "ไม่พบรายการใบเสร็จในระบบ",
                                            "username": username})
                        else:
                            for rr in recs_one:
                                results.append({**rec, **rr, "username": username})
                                if rr.get("status") == "SUCCESS":
                                    success += 1
                    except Exception as e:
                        results.append({**rec, "status": "ERROR",
                                        "error": str(e).splitlines()[0][:200],
                                        "username": username})
                        log(f"      ✗ ผิดพลาด: {str(e).splitlines()[0][:160]}")
                    done_rows += 1
                    if progress:
                        try: progress(done_rows, total)
                        except Exception: pass
                    _save_payreceipt_report(results, out_path, log=lambda *a: None)

                # logout ก่อนเปลี่ยน user
                _logout_safely(page)
                page.wait_for_timeout(1500)
        finally:
            ctx.close(); browser.close()

    _save_payreceipt_report(results, out_path, log=log)
    fail = sum(1 for r in results if r.get("status") not in ("SUCCESS",))
    log(f"[3/3] สรุป: ดาวน์โหลดใบเสร็จสำเร็จ {success} ใบ (ล้มเหลว/empty {fail})")
    return success, out_path


# =====================================================================
# MODE: appointment (นัดหมายถ่ายบัตร) — เก็บที่อยู่จากใบเสร็จค่าธรรมเนียมใบอนุญาตทำงาน
# =====================================================================

def _appt_open_detail(page: Page, req_no: str, log=print) -> tuple[bool, str]:
    """เปิดหน้า detail ของคำขอจาก Tracking. flow: goto Tracking → search req_no → click openDetail
    คืน (ok, error)
    """
    try:
        page.goto(TRACKING_URL, wait_until="domcontentloaded", timeout=40_000)
        page.wait_for_timeout(2200)
        if _is_logged_out(page):
            return False, "session หมดอายุ"
        _search_request(page, req_no)
        page.wait_for_timeout(1500)

        has = page.evaluate(
            r"""() => !!document.querySelector('a[onclick*="openDetail"]')"""
        )
        if not has:
            return False, "ไม่พบเลขคำขอในระบบ"

        try:
            with page.expect_navigation(timeout=20_000, wait_until="domcontentloaded"):
                page.evaluate(
                    r"""() => { const a = document.querySelector('a[onclick*="openDetail"]'); if (a) a.click(); }"""
                )
        except Exception:
            pass
        page.wait_for_timeout(2200)

        # เปลี่ยนเป็นภาษาไทย + ปิด modal
        try:
            page.locator("#sltLang").first.select_option(value="th")
            page.wait_for_timeout(500)
        except Exception:
            pass
        page.evaluate(
            r"""() => {
            if (window.jQuery) { try { jQuery('.modal').modal('hide'); } catch(e){} }
            document.querySelectorAll('.modal.show .btn-close,.modal.show [data-bs-dismiss]')
                .forEach(b => { try { b.click(); } catch(e){} });
        }"""
        )
        page.wait_for_timeout(400)
        return True, ""
    except Exception as e:
        return False, str(e).splitlines()[0][:200]


def _appt_get_company_th(page: Page) -> str:
    """คลิก Tab 'ข้อมูลคนต่างด้าว' แล้วอ่านค่า 'ชื่อสถานประกอบการ(ไทย)' — retry สูงสุด 4 ครั้ง

    ใช้ pattern เดียวกับ extract_label_value_pairs:
    หา <p class="label-form-info">ชื่อสถานประกอบการ(ไทย)</p> แล้วเดินหา
    <p class="form-info ...">value</p> ตัวถัดไปใน DOM order
    """
    read_js = r"""() => {
        const norm = s => (s||'').replace(/\s+/g,' ').trim();
        const root = document.querySelector('.tab-content, .container, main, body') || document.body;
        const all = Array.from(root.querySelectorAll('.label-form-info, .form-info'));
        for (let i = 0; i < all.length; i++) {
            const el = all[i];
            const cls = el.className || '';
            const text = norm(el.innerText || el.textContent || '');
            if (!cls.includes('label-form-info')) continue;
            if (!/ชื่อสถานประกอบการ\s*\(\s*ไทย\s*\)/.test(text)) continue;
            for (let j = i + 1; j < all.length; j++) {
                const next = all[j];
                const ncls = next.className || '';
                if (ncls.includes('label-form-info')) break;  // เจอ label ถัดไป → เลิก
                if (ncls.includes('form-info')) {
                    const v = norm(next.innerText || next.textContent || '');
                    if (v) return v;
                }
            }
        }
        return '';
    }"""
    for attempt in range(4):
        page.evaluate(
            r"""() => {
                const f = [...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a, button')]
                    .find(a => /ข้อมูลคนต่างด้าว/.test(a.innerText || a.textContent || ''));
                if (f) f.click();
            }"""
        )
        page.wait_for_timeout(1500 if attempt == 0 else 1100)
        try:
            v = page.evaluate(read_js)
        except Exception:
            v = ""
        v = (v or "").strip()
        if v:
            return v
    return ""


def _appt_open_payment_tab(page: Page) -> None:
    """คลิก Tab 'การชำระเงิน'"""
    page.evaluate(
        r"""() => {
            const a = Array.from(document.querySelectorAll('a[href^="#"], .nav-link, .nav a'))
                .find(e => /^การชำระเงิน$/.test((e.textContent||'').trim()));
            if (a) a.click();
        }"""
    )
    page.wait_for_timeout(1800)


def _appt_get_appointment_tab_text(page: Page) -> str:
    """คลิก Tab 'นัดหมาย' / 'การนัดหมาย' / 'รอนัดหมาย' (ถ้ามี) แล้วเก็บข้อความ
    ที่แสดงใน pane นั้น — สำหรับใส่ในรายงานเป็น Note

    คืน: ข้อความ (ตัดให้ไม่เกิน ~500 ตัวอักษร) หรือ "" ถ้าไม่พบ tab
    """
    clicked = page.evaluate(
        r"""() => {
            const tabs = [...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a, button')];
            // หา tab ที่ text มีคำว่า "นัดหมาย" — เรียงลำดับ priority
            const patterns = [/^นัดหมาย$/, /^การนัดหมาย$/, /^รอนัดหมาย$/, /นัดหมาย/];
            for (const re of patterns) {
                const t = tabs.find(a => re.test((a.textContent||a.innerText||'').trim()));
                if (t) { t.click(); return (t.textContent||'').trim(); }
            }
            return '';
        }"""
    )
    if not clicked:
        return ""
    page.wait_for_timeout(1500)
    text = page.evaluate(
        r"""() => {
            const norm = s => (s||'').replace(/[\t\r]+/g,' ').replace(/[ ]{2,}/g,' ').trim();
            // หาก pane ที่ active แล้วมีคำว่า "นัดหมาย" — ใช้ pane นั้น
            const panes = [...document.querySelectorAll('.tab-pane')];
            let target = panes.find(p => {
                if (!(p.classList.contains('active') || p.classList.contains('show'))) return false;
                const t = p.innerText || '';
                return /นัดหมาย/.test(t);
            });
            if (!target) {
                target = panes.find(p => (p.classList.contains('active') || p.classList.contains('show')));
            }
            const raw = target ? (target.innerText || '') : (document.body.innerText || '');
            const lines = raw.split(/\n+/).map(s => s.trim()).filter(Boolean);
            // ตัดบรรทัดเปล่า + บรรทัดที่ซ้ำติดกัน
            const out = [];
            for (const ln of lines) {
                if (out.length && out[out.length-1] === ln) continue;
                out.push(ln);
            }
            return norm(out.join(' | '));
        }"""
    )
    text = (text or "").strip()
    if len(text) > 500:
        text = text[:497] + "..."
    return text


_APPT_LABELS = (
    "การนัดหมาย", "สถานที่", "วันที่", "เวลา",
    "หมายเลขนัดหมาย", "ชื่อคนต่างด้าว", "เลขที่หนังสือเดินทาง",
)


def _parse_appointment_body(body: str) -> dict:
    """แยก สถานที่/วันที่/เวลา/ชื่อ/passport จากข้อความ iframe การนัดหมาย
    (ข้อความต่อกันไม่มี space): '...สถานที่<PLACE>วันที่<DATE>เวลา<TIME>พิมพ์แบบฟอร์ม...'
    ตัวอย่างจริง: 'การนัดหมายพิมพ์แบบฟอร์มนัดหมายทั้งหมดสถานที่ศูนย์บริการใบอนุญาต
    ทำงานของคนต่างด้าว จังหวัดฉะเชิงเทราวันที่16 ธันวาคม 2569เวลาพิมพ์แบบฟอร์มนัดหมาย...
    หมายเลขนัดหมายชื่อคนต่างด้าวเลขที่หนังสือเดินทาง2-CCO001122601108Miss KYI KYI NAING MH190359'
    (ในตัวอย่างนี้ 'เวลา' ว่าง — production ที่จองแล้วจะมีเวลา เช่น '8.30 น.')
    """
    body = re.sub(r"\s+", " ", body or "").strip()
    out = {"place": "", "date": "", "time": "", "worker_name": "", "passport": ""}
    _strip = " :·-\u2022\u2018\u2019"
    # สถานที่: ระหว่าง 'สถานที่' กับ 'วันที่'
    m = re.search(r"สถานที่\s*(.+?)\s*วันที่", body)
    if m:
        out["place"] = m.group(1).strip(_strip)
    # วันที่: ระหว่าง 'วันที่' กับ 'เวลา'
    m = re.search(r"วันที่\s*(.+?)\s*เวลา", body)
    if m:
        out["date"] = m.group(1).strip(_strip)
    # เวลา: ระหว่าง 'เวลา' กับ marker ถัดไป (อาจว่างได้)
    m = re.search(
        r"เวลา\s*(.*?)\s*(?:พิมพ์แบบฟอร์ม|ขอยกเลิก|หมายเลขนัดหมาย|ชื่อคนต่างด้าว|$)",
        body,
    )
    if m:
        t = m.group(1).strip(_strip)
        # ต้องมีตัวเลข (เวลา เช่น 08.30/8:30) — กันจับ label/ปุ่มภาษาไทยเมื่อไม่มีเวลาจริง
        if t and len(t) <= 40 and re.search(r"\d", t):
            out["time"] = t
    # ชื่อคนต่างด้าว (Latin): Mr/Mrs/Miss/Ms + ตัวอักษร
    nm = re.search(
        r"(Mr\.?|Mrs\.?|Miss|Ms\.?|Mss\.?)\s+([A-Z][A-Za-z\.\s]{1,80}?)"
        r"(?=\s*[A-Z]{1,2}\d|\s{2,}|ยกเลิก|$)",
        body,
    )
    if nm:
        out["worker_name"] = (nm.group(1) + " " + nm.group(2)).strip()
    # passport: Myanmar M[A-Z]\d, Cambodia C[A-Z]\d, generic 1-2 ตัวอักษร + 6-9 หลัก
    # ใช้ขอบเขต Latin/digit เอง (ไม่ใช้ \b) เพราะอักษรไทยเป็น word-char ใน Python regex
    # → \b หลังเลข passport ที่ตามด้วยไทยทันที (เช่น 'MH190359ยกเลิก') จะไม่เกิด
    for pat in (
        r"(?<![A-Za-z0-9])(M[A-Z]\d{6,8})(?![A-Za-z0-9])",
        r"(?<![A-Za-z0-9])(C[A-Z]\d{6,8})(?![A-Za-z0-9])",
        r"(?<![A-Za-z0-9])([A-Z]{1,2}\d{6,9})(?![A-Za-z0-9])",
    ):
        pm = re.search(pat, body)
        if pm:
            out["passport"] = pm.group(1)
            break
    return out


def _extract_appointment_details(page: Page, log=print) -> dict:
    """เปิดแท็บ 'การนัดหมาย' (#tab_default_5) → อ่าน iframe (queue-fe*.doe.go.th)
    → คืน {place, date, time, worker_name, passport, raw, error}

    ถ้ายังไม่ได้จอง / ไม่มีการนัดหมาย → คืนค่าว่างพร้อม error อธิบายเหตุ
    (record ยังคงถูกใส่ในรายงาน — ผู้ใช้กรองสถานะเองใน UI)
    """
    blank = {"place": "", "date": "", "time": "", "worker_name": "",
             "passport": "", "raw": "", "error": ""}
    # 1) คลิกแท็บการนัดหมาย
    try:
        page.evaluate(r"""() => {
            if (window.jQuery) { try { jQuery('a[href="#tab_default_5"]').tab('show'); } catch(e){} }
            const a = document.querySelector('a[href="#tab_default_5"]');
            if (a) a.click();
            const els = [...document.querySelectorAll('a,.nav-link,.nav-tabs a,.nav a,[role="tab"]')];
            const t = els.find(x => /^\s*การนัดหมาย\s*$/.test((x.textContent||'').trim()));
            if (t) t.click();
        }""")
    except Exception:
        pass
    page.wait_for_timeout(2200)
    # 2) หา iframe src
    try:
        iframe_src = page.evaluate(
            "() => (document.querySelector('#link_appointment') "
            "|| document.querySelector('iframe.iframe_show') "
            "|| document.querySelector('iframe[src*=\"queue\"]') || {}).src || ''"
        )
    except Exception:
        iframe_src = ""
    if not iframe_src:
        try:
            empty_txt = page.evaluate(
                "() => { const p=document.querySelector('#tab_default_5'); "
                "return p ? (p.innerText||'').trim().slice(0,120) : ''; }"
            )
        except Exception:
            empty_txt = ""
        b = dict(blank)
        b["error"] = ("record นี้ไม่มีการนัดหมาย"
                      if "ไม่มีการนัดหมาย" in (empty_txt or "")
                      else "ไม่พบ iframe การนัดหมาย")
        return b
    # 3) หา frame ที่ตรง + รอโหลด
    appt_frame = None
    for _ in range(15):
        for fr in page.frames:
            if fr.name == "link_appointment" \
                    or "bookingdate" in (fr.url or "") \
                    or "bookingdetail" in (fr.url or ""):
                appt_frame = fr
                break
        if appt_frame:
            break
        page.wait_for_timeout(500)
    if not appt_frame:
        b = dict(blank)
        b["error"] = "iframe การนัดหมายโหลดไม่สำเร็จ"
        return b
    try:
        appt_frame.wait_for_load_state("domcontentloaded", timeout=15_000)
    except Exception:
        pass
    page.wait_for_timeout(2000)
    # 4) อ่าน body text จาก iframe
    try:
        body = appt_frame.evaluate(
            "() => document.body ? document.body.innerText : ''"
        ) or ""
    except Exception as e:
        b = dict(blank)
        b["error"] = f"อ่าน iframe การนัดหมายไม่สำเร็จ: {str(e)[:80]}"
        return b
    norm = re.sub(r"\s+", " ", body).strip()
    # ยังไม่ได้จอง (booking form) → ไม่มีวันที่/เวลา/สถานที่
    if re.search(r"ยังไม่มีการนัดหมาย|เพิ่มวันนัดหมาย|กรุณาเลือก", norm):
        b = dict(blank)
        b["raw"] = norm[:500]
        b["error"] = "ยังไม่ได้จองการนัดหมาย"
        return b
    parsed = _parse_appointment_body(norm)
    parsed["raw"] = norm[:500]
    parsed["error"] = ""
    return parsed


def _appt_detect_payment_case(page: Page) -> str:
    """ตรวจ Case บนแท็บการชำระเงิน:
       - 'case1': มี Record การชำระเงิน + ปุ่ม 'หลักฐานการชำระเงิน'
       - 'case2': มี banner 'ท่านได้ชำระเงินสำเร็จแล้ว' + ปุ่ม 'ดูใบเสร็จรับเงิน'
       - 'empty': ไม่พบทั้งคู่ → ข้าม
    """
    info = page.evaluate(
        r"""() => {
            const vis = el => { try { const r = el.getBoundingClientRect(); return r.width>0 && r.height>0; } catch(e){ return false; } };
            const allText = document.body.innerText || '';
            const hasSuccess = /ท่านได้ชำระเงินสำเร็จแล้ว/.test(allText);
            const hasReceiptBtn = Array.from(document.querySelectorAll('button,a'))
                .some(e => vis(e) && /ดูใบเสร็จรับเงิน/.test((e.textContent||'').trim()));
            const hasProofBtn = Array.from(document.querySelectorAll('button,a'))
                .some(e => vis(e) && /หลักฐานการชำระเงิน/.test((e.textContent||'').trim()));
            return { hasSuccess, hasReceiptBtn, hasProofBtn };
        }"""
    )
    if info.get("hasProofBtn"):
        return "case1"
    if info.get("hasSuccess") or info.get("hasReceiptBtn"):
        return "case2"
    return "empty"


def _appt_parse_address(text: str) -> tuple[str, str]:
    """แยก (จังหวัด, อำเภอ/เขต) จากข้อความใบเสร็จ
    Format ตัวอย่าง:
        ที่อยู่
        (Address)
        เลขที่ 1/65 หมู่ 5 แขวง/ตำบล คานหาม เขต/อำเภอ อุทัย จังหวัด
        พระนครศรีอยุธยา รหัสไปรษณีย์ 13210
    ค่า 'จังหวัด' อาจอยู่บรรทัดถัดไป — รวมหลายบรรทัดก่อน regex
    """
    if not text:
        return "", ""
    import re as _re
    # ตัดเฉพาะส่วนใกล้ ๆ label ที่อยู่ — กันชนะเอาจังหวัดจากนายจ้าง/ที่อื่น
    addr_block = text
    m_addr = _re.search(r"ที่อยู่\s*\n\s*\(Address\)([\s\S]{0,400})", text)
    if m_addr:
        addr_block = m_addr.group(1)
    # ยุบ newline เป็น space เพื่อให้ regex จับข้าม linebreak ได้
    flat = _re.sub(r"\s+", " ", addr_block).strip()

    province = ""
    district = ""
    # จังหวัดมีได้สูงสุด 2 คำ (เช่น "นครราชสีมา") — หยุดก่อนเห็นคำ 'รหัสไปรษณีย์' หรือตัวเลข 5 หลัก
    m_prov = _re.search(
        r"จังหวัด\s+((?:(?!รหัสไปรษณีย์|\d{5})\S+)(?:\s+(?:(?!รหัสไปรษณีย์|\d{5})\S+))?)\s+(?:รหัสไปรษณีย์|\d{5})",
        flat,
    )
    if not m_prov:
        m_prov = _re.search(r"จังหวัด\s+(\S+)", flat)
    if m_prov:
        province = m_prov.group(1).strip().rstrip(",")

    m_dist = _re.search(r"เขต/อำเภอ\s+([^\s]+)", flat)
    if m_dist:
        district = m_dist.group(1).strip().rstrip(",")
    return province, district


def _appt_case1_download(
    page: Page, req_no: str, save_dir: Path, log=print,
) -> tuple[bytes, str, str]:
    """Case 1: คลิกปุ่ม 'หลักฐานการชำระเงิน'
       - พยายามเลือกปุ่มบน "แถวเดียวกับ" ข้อความ 'ค่าธรรมเนียมใบอนุญาตทำงาน' ก่อน
       - ถ้าไม่เจอ → fallback คลิกปุ่ม 'หลักฐานการชำระเงิน' ตัวแรกที่มองเห็น
    """
    info = page.evaluate(
        r"""() => {
            const vis = el => { try { const r = el.getBoundingClientRect(); return r.width>0 && r.height>0; } catch(e){ return false; } };
            const trs = Array.from(document.querySelectorAll('table tbody tr')).filter(vis);
            let preferred = -1;
            for (let i = 0; i < trs.length; i++) {
                const tx = trs[i].innerText || '';
                if (/ค่าธรรมเนียมใบอนุญาตทำงาน/.test(tx)
                        && Array.from(trs[i].querySelectorAll('button,a'))
                            .some(b => /หลักฐานการชำระเงิน/.test((b.textContent||'').trim()))) {
                    preferred = i; break;
                }
            }
            const allBtns = Array.from(document.querySelectorAll('button,a'))
                .filter(b => vis(b) && /หลักฐานการชำระเงิน/.test((b.textContent||'').trim()));
            return { preferred, totalBtns: allBtns.length };
        }"""
    )
    preferred = int(info.get("preferred", -1))
    total_btns = int(info.get("totalBtns", 0))
    if total_btns <= 0:
        return b"", "", "ไม่พบปุ่ม 'หลักฐานการชำระเงิน' บนหน้า"
    use_row = preferred >= 0
    log(f"      · case1 row='ค่าธรรมเนียมใบอนุญาตทำงาน' {'พบ' if use_row else 'ไม่พบ'} — ปุ่มทั้งหมด {total_btns}")

    def _do_click():
        if use_row:
            page.evaluate(
                r"""(idx) => {
                    const vis = el => { try { const r = el.getBoundingClientRect(); return r.width>0 && r.height>0; } catch(e){ return false; } };
                    const trs = Array.from(document.querySelectorAll('table tbody tr')).filter(vis);
                    const tr = trs[idx];
                    if (!tr) return;
                    const btn = Array.from(tr.querySelectorAll('button,a'))
                        .find(b => /หลักฐานการชำระเงิน/.test((b.textContent||'').trim()));
                    if (btn) btn.click();
                }""", preferred,
            )
        else:
            page.evaluate(
                r"""() => {
                    const vis = el => { try { const r = el.getBoundingClientRect(); return r.width>0 && r.height>0; } catch(e){ return false; } };
                    const btn = Array.from(document.querySelectorAll('button,a'))
                        .find(b => vis(b) && /หลักฐานการชำระเงิน/.test((b.textContent||'').trim()));
                    if (btn) btn.click();
                }"""
            )

    body = b""; err_last = ""
    for attempt in range(1, 4):
        b, err = _grab_pdf_after_click(page, _do_click, log=lambda *a: None)
        if b and len(b) >= 500 and b[:5] == b"%PDF-":
            body = b; break
        err_last = err or "ไม่ได้ไฟล์"
        log(f"      · ⚠ ดาวน์โหลดครั้งที่ {attempt} ไม่สำเร็จ ({err_last})")
        page.wait_for_timeout(1500)
    if not body:
        return b"", "", f"ดาวน์โหลดล้มเหลว: {err_last}"

    save_dir.mkdir(parents=True, exist_ok=True)
    fname = f"{_safe_filename(req_no)}_case1_receipt.pdf"
    try:
        (save_dir / fname).write_bytes(body)
    except Exception as e:
        return body, "", f"เขียนไฟล์ไม่สำเร็จ: {str(e)[:120]}"
    return body, fname, ""


def _appt_case2_download(
    page: Page, req_no: str, save_dir: Path, log=print,
) -> tuple[bytes, str, str]:
    """Case 2: คลิก 'ดูใบเสร็จรับเงิน' → /Center/MultiplePayments → search ด้วย req_no → click record #2
    → ดาวน์โหลด PDF (อาจมีหลายหน้า) → คืน (body, filename, error)
    """
    # คลิกปุ่ม "ดูใบเสร็จรับเงิน" → navigate
    try:
        with page.expect_navigation(timeout=20_000, wait_until="domcontentloaded"):
            page.evaluate(
                r"""() => {
                    const vis = el => { try { const r = el.getBoundingClientRect(); return r.width>0 && r.height>0; } catch(e){ return false; } };
                    const b = Array.from(document.querySelectorAll('button,a'))
                        .find(e => vis(e) && /ดูใบเสร็จรับเงิน/.test((e.textContent||'').trim()));
                    if (b) b.click();
                }"""
            )
    except Exception:
        pass
    page.wait_for_timeout(3000)

    if "/Center/MultiplePayments" not in (page.url or ""):
        return b"", "", f"ไม่ได้เข้าหน้า MultiplePayments (อยู่ที่ {page.url[:80]})"

    # ค้นหาด้วยเลขคำขอ — ใช้ช่องค้นหาบนหน้า MultiplePayments
    # หน้านี้มักมี input ค้นหาที่ผูกกับ tablesorter/datatable — ใส่ค่าและ trigger keyup
    page.evaluate(
        r"""(reqNo) => {
            const inputs = Array.from(document.querySelectorAll('input[type="text"], input[type="search"]'))
                .filter(i => { try { const r = i.getBoundingClientRect(); return r.width>0 && r.height>0; } catch(e){ return false; } });
            // หาช่องค้นหาที่มี placeholder/label ใกล้คำว่า "ค้นหา" หรือ "เลขคำขอ"
            const target = inputs.find(i =>
                /ค้นหา|เลขที่คำขอ|เลขคำขอ|Request/i.test((i.placeholder||'') + ' ' + (i.getAttribute('aria-label')||'')))
                || inputs[0];
            if (target) {
                target.focus(); target.value = reqNo;
                target.dispatchEvent(new Event('input', {bubbles:true}));
                target.dispatchEvent(new Event('keyup', {bubbles:true}));
                target.dispatchEvent(new Event('change', {bubbles:true}));
            }
        }""", req_no,
    )
    page.wait_for_timeout(1500)
    # ปุ่มค้นหา (ถ้ามี)
    page.evaluate(
        r"""() => {
            const btns = Array.from(document.querySelectorAll('button, a.btn, input[type=button], input[type=submit]'))
                .filter(b => { try { const r = b.getBoundingClientRect(); return r.width>0 && r.height>0; } catch(e){ return false; } });
            const f = btns.find(b => /^ค้นหา$/.test(((b.innerText||b.value||'')).trim()));
            if (f) f.click();
        }"""
    )
    page.wait_for_timeout(2500)

    # นับจำนวนแถวที่มีปุ่ม 'ใบเสร็จรับเงิน'
    row_count = page.evaluate(
        r"""() => {
            const trs = Array.from(document.querySelectorAll('table tbody tr'));
            return trs.filter(tr =>
                /ใบเสร็จรับเงิน/.test(tr.innerText || '') &&
                Array.from(tr.querySelectorAll('button,a')).some(b => /ใบเสร็จรับเงิน/.test((b.textContent||'').trim()))
            ).length;
        }"""
    )
    if not row_count:
        return b"", "", "ไม่พบแถวใบเสร็จในหน้า MultiplePayments"
    log(f"      · พบ {row_count} แถวใบเสร็จ — เลือก record ที่ 2 (ตามสเปก)")
    target_idx = 1 if row_count >= 2 else 0  # ผู้ใช้ระบุ record ที่ 2 เสมอ; ถ้ามีแถวเดียว fallback ไป 0

    def _do_click(idx: int = int(target_idx)):
        page.evaluate(
            r"""(idx) => {
                const trs = Array.from(document.querySelectorAll('table tbody tr'))
                    .filter(tr => /ใบเสร็จรับเงิน/.test(tr.innerText || ''));
                const tr = trs[idx];
                if (!tr) return;
                const btn = Array.from(tr.querySelectorAll('button,a'))
                    .find(b => /ใบเสร็จรับเงิน/.test((b.textContent||'').trim()));
                if (btn) btn.click();
            }""", idx,
        )

    body = b""; err_last = ""
    for attempt in range(1, 4):
        b, err = _grab_pdf_after_click(page, _do_click, log=lambda *a: None)
        if b and len(b) >= 500 and b[:5] == b"%PDF-":
            body = b; break
        err_last = err or "ไม่ได้ไฟล์"
        log(f"      · ⚠ ดาวน์โหลดครั้งที่ {attempt} ไม่สำเร็จ ({err_last})")
        page.wait_for_timeout(1500)
    if not body:
        return b"", "", f"ดาวน์โหลดล้มเหลว: {err_last}"

    save_dir.mkdir(parents=True, exist_ok=True)
    fname = f"{_safe_filename(req_no)}_case2_receipt.pdf"
    try:
        (save_dir / fname).write_bytes(body)
    except Exception as e:
        return body, "", f"เขียนไฟล์ไม่สำเร็จ: {str(e)[:120]}"
    return body, fname, ""


def _bt30_ctn_receipt_fee_kind(pdf_bytes: bytes) -> str:
    """ตรวจประเภทค่าธรรมเนียมจากเนื้อ PDF ใบเสร็จ
    → 'workpermit'  = ค่าธรรมเนียมใบอนุญาตทำงาน (ใบที่ต้องการ)
    → 'submission'  = ค่ายื่นคำขอ (Submission Fee)
    → 'unknown'     = ระบุไม่ได้
    """
    try:
        import io
        from pypdf import PdfReader
        text = "\n".join((pg.extract_text() or "") for pg in PdfReader(io.BytesIO(pdf_bytes)).pages)
    except Exception:
        return "unknown"
    if re.search(r"ค่าธรรมเนียม.{0,12}ใบอนุญาตทำงาน|Work\s*Permit\s*Fee", text):
        return "workpermit"
    if re.search(r"ค่ายื่นคำขอ|Submission\s*Fee", text):
        return "submission"
    return "unknown"


def _bt30_ctn_grab_receipt_case1(
    page: Page, log=print,
) -> tuple[bytes, str, str, str]:
    """Case 1: วนคลิกปุ่ม 'หลักฐานการชำระเงิน' ทุกปุ่มบนหน้ารายละเอียด → ดาวน์โหลดใบเสร็จแต่ละใบ
    → อ่านยอดเงิน + ประเภทค่าธรรมเนียมจากเนื้อ PDF → เลือกใบ 'ค่าธรรมเนียมใบอนุญาตทำงาน'
    (ถ้าไม่พบประเภทนี้ → fallback ใช้ใบที่ยอดเงินมากสุด — ค่าธรรมเนียมใบอนุญาตมักสูงกว่าค่ายื่นคำขอ)

    คืน (pdf_bytes, amount, fee_kind, error)
      - amount   : สตริงยอดเงิน เช่น '1800' (ว่าง = อ่านไม่ได้)
      - fee_kind : 'workpermit' | 'submission' | 'unknown'
    """
    n = int(page.evaluate(
        r"""() => {
            const vis = el => { try { const r = el.getBoundingClientRect(); return r.width>0 && r.height>0; } catch(e){ return false; } };
            return Array.from(document.querySelectorAll('button,a'))
                .filter(b => vis(b) && /หลักฐานการชำระเงิน/.test((b.textContent||'').trim())).length;
        }"""
    ) or 0)
    if n <= 0:
        return b"", "", "", "ไม่พบปุ่ม 'หลักฐานการชำระเงิน' บนหน้ารายละเอียด"

    def _click_idx(idx: int):
        page.evaluate(
            r"""(idx) => {
                const vis = el => { try { const r = el.getBoundingClientRect(); return r.width>0 && r.height>0; } catch(e){ return false; } };
                const btns = Array.from(document.querySelectorAll('button,a'))
                    .filter(b => vis(b) && /หลักฐานการชำระเงิน/.test((b.textContent||'').trim()));
                if (btns[idx]) btns[idx].click();
            }""", idx,
        )

    found: list[tuple[bytes, str, str]] = []  # (pdf_bytes, amount, fee_kind)
    for i in range(n):
        body = b""
        for _attempt in range(1, 3):
            b, _err = _grab_pdf_after_click(page, lambda i=i: _click_idx(i), log=lambda *a: None)
            if b and len(b) >= 500 and b[:5] == b"%PDF-":
                body = b
                break
            page.wait_for_timeout(1200)
        if not body:
            log(f"      · ⚠ ใบเสร็จปุ่มที่ {i + 1}/{n} โหลดไม่สำเร็จ")
            continue
        amount = _extract_pdf_amount(body)
        kind = _bt30_ctn_receipt_fee_kind(body)
        found.append((body, amount, kind))
        log(f"      · ใบเสร็จ {i + 1}/{n}: ยอด={amount or '-'} ประเภท={kind}")

    if not found:
        return b"", "", "", "ดาวน์โหลดใบเสร็จไม่สำเร็จทุกปุ่ม"

    wp = [f for f in found if f[2] == "workpermit"]
    if wp:
        chosen = wp[0]
    else:
        def _amt(a: str) -> float:
            try:
                return float((a or "0").replace("_", "."))
            except Exception:
                return 0.0
        chosen = max(found, key=lambda f: _amt(f[1]))
        log("      · ⚠ ไม่พบใบ 'ค่าธรรมเนียมใบอนุญาตทำงาน' → ใช้ใบยอดเงินสูงสุดแทน")
    return chosen[0], chosen[1], chosen[2], ""


def _appt_extract_address_for_req(body: bytes, req_no: str) -> tuple[str, str, str]:
    """อ่าน PDF (อาจหลายหน้า) → หา page ที่ตรงกับ req_no → คืน (จังหวัด, อำเภอ, raw_text_of_matched_page)
    ถ้าไม่พบ exact match → ใช้ page แรก
    """
    pages = _payreceipt_split_pages(body)
    if not pages:
        return "", "", ""
    matched_text = ""
    for _bytes, txt in pages:
        if req_no and req_no in (txt or ""):
            matched_text = txt
            break
    if not matched_text:
        matched_text = pages[0][1]
    prov, dist = _appt_parse_address(matched_text)
    return prov, dist, matched_text


def _appt_process_record(
    page: Page, rec: dict, save_dir: Path, log=print,
) -> dict[str, Any]:
    """ประมวลผล 1 คำขอ — คืน dict ที่มี: company_th, province, district, pdf_file, case, status, error"""
    req_no = rec.get("req_no", "")
    out = {
        "req_no": req_no,
        "username": rec.get("username", ""),
        "company_th": "",
        "province": "",
        "district": "",
        "pdf_file": "",
        "case": "",
        "appointment_note": "",
        "status": "FAIL",
        "error": "",
    }
    ok, err = _appt_open_detail(page, req_no, log=log)
    if not ok:
        out["status"] = "FAIL"
        out["error"] = err
        return out

    # ขั้น 2.4: ชื่อสถานประกอบการ(ไทย)
    company = _appt_get_company_th(page)
    out["company_th"] = company
    if not company:
        log("      · ⚠ อ่าน 'ชื่อสถานประกอบการ(ไทย)' ไม่ได้ (ดำเนินการต่อ)")

    # ขั้น 2.4.5: tab 'นัดหมาย' (ถ้ามี) — เก็บข้อความที่แสดงเป็น Note
    try:
        appt_note = _appt_get_appointment_tab_text(page)
    except Exception as e:
        appt_note = ""
        log(f"      · ⚠ อ่าน tab นัดหมายไม่ได้: {str(e).splitlines()[0][:120]}")
    out["appointment_note"] = appt_note
    if appt_note:
        log(f"      · tab นัดหมาย: {appt_note[:120]}")

    # ขั้น 2.5: เปิดแท็บการชำระเงิน
    _appt_open_payment_tab(page)
    case = _appt_detect_payment_case(page)
    out["case"] = case
    log(f"      · ตรวจการชำระเงิน: {case}")
    if case == "empty":
        out["status"] = "EMPTY"
        out["error"] = "ไม่มี Record การชำระเงิน — ข้าม"
        return out

    if case == "case1":
        body, fname, derr = _appt_case1_download(page, req_no, save_dir, log=log)
    else:  # case2
        body, fname, derr = _appt_case2_download(page, req_no, save_dir, log=log)

    if not body:
        out["status"] = "FAIL"
        out["error"] = derr or "ดาวน์โหลด PDF ล้มเหลว"
        return out

    out["pdf_file"] = fname
    prov, dist, _ = _appt_extract_address_for_req(body, req_no)
    out["province"] = prov
    out["district"] = dist
    if not (prov or dist):
        out["status"] = "PARTIAL"
        out["error"] = "ดาวน์โหลด PDF ได้ แต่หาที่อยู่ในไฟล์ไม่พบ"
    else:
        out["status"] = "SUCCESS"
    log(f"      · จังหวัด={prov or '-'} | อำเภอ/เขต={dist or '-'}")
    return out


def _save_appointment_report(rows: list[dict[str, Any]], out_path: Path, log=print) -> None:
    """เซฟรายงาน — คอลัมน์ PDF เป็น hyperlink (file://) แสดงชื่อไฟล์ น้ำเงินขีดเส้นใต้"""
    wb = Workbook()
    ws = wb.active
    ws.title = "นัดหมายถ่ายบัตร"
    headers = [
        "ลำดับ", "Username", "เลขที่คำขอ",
        "ชื่อสถานประกอบการ(ไทย)", "จังหวัด", "อำเภอ/เขต",
        "Case", "สถานะ", "ไฟล์ PDF", "ข้อความ Tab นัดหมาย", "หมายเหตุ",
    ]
    ws.append(headers)
    head_fill = PatternFill("solid", fgColor="305496")
    link_font = Font(color="0563C1", underline="single")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    text_cols = {"เลขที่คำขอ"}
    save_subdir = "appointment_receipts"
    for i, r in enumerate(rows, start=1):
        ws.append([
            i,
            r.get("username", ""),
            r.get("req_no", ""),
            r.get("company_th", ""),
            r.get("province", ""),
            r.get("district", ""),
            r.get("case", ""),
            r.get("status", ""),
            r.get("pdf_file", ""),
            r.get("appointment_note", ""),
            r.get("error", ""),
        ])
        row_idx = ws.max_row
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=row_idx, column=c_idx)
            if h in text_cols and cell.value not in (None, ""):
                cell.value = str(cell.value)
                cell.number_format = "@"
            if h == "ไฟล์ PDF" and cell.value:
                fname = str(cell.value)
                abs_path = (out_path.parent / save_subdir / fname).resolve()
                try:
                    cell.hyperlink = abs_path.as_uri()
                except Exception:
                    cell.hyperlink = f"{save_subdir}/{fname}"
                cell.value = fname
                cell.font = link_font
            if h == "ข้อความ Tab นัดหมาย" and cell.value:
                cell.alignment = Alignment(wrap_text=True, vertical="top")

    widths = [6, 28, 20, 40, 22, 22, 10, 12, 50, 60, 40]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions

    out_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(out_path)
        log(f"      · เซฟรายงาน → {out_path}")
    except Exception as e:
        log(f"      · ⚠ เซฟรายงานไม่สำเร็จ: {str(e)[:160]}")


# ─────────────────────────────────────────────────────────────────
# โหมดใหม่ — ดาวน์โหลด 'แบบ บต.30' จาก tab 'เอกสารตอบรับจากระบบ' (สถานะ WP)
#   วน login ทุกบัญชีใน UsernameLogin.xlsx → Tracking → filter WP →
#   ทุก req_no: เปิด detail → tab 'เอกสารตอบรับ' → หา row 'แบบ บต.XX' → download
#   → อ่าน PDF text (pypdf) → ดึง passport + form number → rename เป็น
#      {PASSPORT}_BT{XX}_CTN.pdf  (เช่น MH788309_BT30_CTN.pdf)
# ─────────────────────────────────────────────────────────────────
_THAI_DIGIT_TR = str.maketrans("๐๑๒๓๔๕๖๗๘๙", "0123456789")


def _thai_to_arabic(s: str) -> str:
    return (s or "").translate(_THAI_DIGIT_TR)


def _extract_bt_form_number_from_pdf(pdf_text: str) -> str:
    """หา form number จาก PDF text:
    - 'แบบ บต. ๓๐'    → '30'
    - 'FORM WP. 30'   → '30'
    ค้นเฉพาะช่วงต้นไฟล์ (~2000 ตัวอักษร) ที่เป็น header
    คืน '' ถ้าหาไม่พบ
    """
    if not pdf_text:
        return ""
    head = pdf_text[:2500]
    # ภาษาไทย: 'แบบ บต. ๓๐' / 'แบบ บต.30' / 'แบบบต.30'
    m = re.search(r"แบบ\s*บต\s*\.?\s*([0-9๐-๙]{1,3})", head)
    if m:
        num = _thai_to_arabic(m.group(1)).strip()
        if num.isdigit():
            return num
    # อังกฤษ: 'FORM WP. 30' / 'FORM WP.30'
    m = re.search(r"FORM\s+WP\s*\.?\s*(\d{1,3})", head, flags=re.IGNORECASE)
    if m:
        return m.group(1)
    return ""


def _extract_passport_from_pdf(pdf_text: str) -> str:
    """หา passport จาก PDF text ของแบบ บต.30
    ที่ระบบดาวน์โหลดมา ตัวแบบฟอร์มเป็น placeholder (…) — ค่าที่กรอกจริงจะถูก
    เรียงเป็นบล็อคข้อความท้ายหน้า 1 (คั่นด้วยขึ้นบรรทัดใหม่) เช่น:
        อ้างอิงจากเอกสารเวอร์ชัน 1.0.1
        MISS SANIN  NOURN
        กัมพูชา
        ...
        ✔                    ← checkmark หน้ากล่อง 'หนังสือเดินทาง'
        T0892785              ← passport ที่กรอก
        MIN PHNOM PENH        ← 'ออกให้โดย'
        กัมพูชา                ← ประเทศ
        ...

    รูปแบบ passport: 1-2 ตัวอักษร + 6-9 ตัวเลข/อักษร (เช่น T0892785, MH788309, A12345678)
    คืน '' ถ้าหาไม่พบ
    """
    if not pdf_text:
        return ""
    # กันคำที่ดูเหมือน passport แต่ไม่ใช่ (เกิดจาก visa type / prefix)
    _reject = {"PASSPORT", "PASSPO", "MISS", "MRS", "MSTR", "NON", "NONE"}
    # 1) หลังเครื่องหมายถูก (✔/☑/✅) ทันที = passport ที่ระบบกรอกไว้
    for pat in (
        r"[✔☑✅√✓]\s*\n\s*([A-Z][A-Z0-9]{5,13})\b",
        r"[✔☑✅√✓]\s*([A-Z][A-Z0-9]{5,13})\b",
    ):
        m = re.search(pat, pdf_text)
        if m:
            pp = m.group(1).strip().upper()
            if pp not in _reject:
                return pp
    # 2) หลัง 'หนังสือเดินทาง' → หา token passport-like ตัวแรก
    #    (บล็อกข้อมูลจะอยู่ท้ายหน้า 1 หลังคำ 'หนังสือเดินทาง' หลายบรรทัด)
    m = re.search(
        r"หนังสือเดินทาง[\s\S]{0,1500}?\b([A-Z]{1,2}\d{6,9})\b",
        pdf_text,
    )
    if m:
        pp = m.group(1).strip().upper()
        if pp not in _reject:
            return pp
    # 3) fallback แบบเก่า — หลัง 'Passport เลขที่'
    for pat in (
        r"(?:Passport|หนังสือเดินทาง)[\s\S]{0,80}?เลขที่[.\s]*([A-Z][A-Z0-9]{5,13})",
        r"(?:Passport)[\s\S]{0,60}?([A-Z]{1,2}\d{6,9})",
    ):
        m = re.search(pat, pdf_text, flags=re.IGNORECASE)
        if m:
            pp = m.group(1).strip().upper()
            if pp not in _reject:
                return pp
    return ""


def _uniq_pdf_path(path: Path) -> Path:
    """ถ้าไฟล์ชนกัน → เพิ่ม _2, _3, ... ท้ายชื่อ (ก่อนนามสกุล)"""
    if not path.exists():
        return path
    n = 2
    while True:
        alt = path.parent / f"{path.stem}_{n}{path.suffix}"
        if not alt.exists():
            return alt
        n += 1


def _bt30_ctn_download(
    page: Page, log=print, pat: str = r"บต\.?\s*30", doc_name: str = "บต.30",
) -> tuple[bytes, str, str]:
    r"""ดาวน์โหลดเอกสารในแท็บ 'เอกสารตอบรับจากระบบ' บนหน้า detail (default: แบบ บต.30)
    (สมมติ page อยู่บนหน้า detail ของคำขอแล้ว)
    Args:
        pat: regex ของป้ายเอกสารที่ต้องการ (เช่น r"บต\.?\s*30" หรือ r"บต\.?\s*25")
        doc_name: ชื่อเอกสารสำหรับข้อความ log/error (เช่น "บต.30", "บต.25")
    คืน (pdf_bytes, matched_label, error_msg)
      - pdf_bytes: bytes ของ PDF (ว่าง = ล้มเหลว)
      - matched_label: ข้อความ label ของแถวที่คลิก (เช่น 'แบบ บต.30 คำขอต่ออายุ...')
      - error_msg: '' ถ้าสำเร็จ, มิฉะนั้นข้อความ error
    """
    # 1) คลิกแท็บเอกสารตอบรับ
    clicked = page.evaluate(
        r"""() => {
            const re = /เอกสารตอบรับ/;
            const f = [...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a')]
                .find(a => re.test(a.innerText || ''));
            if (f) { f.click(); return true; }
            return false;
        }"""
    )
    if not clicked:
        return b"", "", "ไม่พบแท็บเอกสารตอบรับ"
    page.wait_for_timeout(1800)

    # 2) หาลิงก์ GetDocumentConfirm ที่อยู่ใต้แถวป้าย 'บต.30' (ใช้ pattern เดียวกับ _bt55)
    #    โครงสร้าง: <div>label</div>...<div>...<a onclick="GetDocumentConfirm(...)">...</a></div>
    pick_js = r"""(args) => {
        const { pat, mode } = args;
        const re = new RegExp(pat);
        const seen = new Set(); const list = [];
        let bestLabel = '';
        document.querySelectorAll('[onclick*="GetDocumentConfirm"]').forEach(a => {
            const oc = a.getAttribute('onclick') || '';
            if (!oc) return;
            const grp = a.parentElement;
            const lbl = (grp && grp.previousElementSibling)
                ? ((grp.previousElementSibling.innerText || '').trim()) : '';
            if (!re.test(lbl)) return;
            if (seen.has(oc)) return;
            seen.add(oc); list.push({a, lbl});
        });
        if (!list.length) return mode === 'url' ? '' : (mode === 'label' ? '' : false);
        // เลือกป้ายที่สั้นสุด (ตรงกับ บต.30 มากกว่า)
        list.sort((x, y) => (x.lbl.length - y.lbl.length));
        const first = list[0];
        bestLabel = first.lbl;
        if (mode === 'label') return bestLabel;
        if (mode === 'url')   return (first.a.getAttribute('onclick') || '') + ' | ' + (first.a.getAttribute('href') || '');
        first.a.click(); return true;
    }"""

    # เก็บ label ที่ match ไว้ก่อน (สำหรับ log/report)
    matched_label = ""
    try:
        matched_label = page.evaluate(pick_js, {"pat": pat, "mode": "label"}) or ""
    except Exception:
        pass
    if not matched_label:
        return b"", "", f"ไม่พบเอกสาร {doc_name} ในแท็บเอกสารตอบรับ (อาจยังไม่ถูกสร้าง)"

    # 3) ลอง prefetch URL — เร็วกว่าเปิด popup
    prefetch_url = ""
    try:
        raw = page.evaluate(pick_js, {"pat": pat, "mode": "url"}) or ""
        prefetch_url = _extract_doc_url(raw, page.url)
    except Exception:
        prefetch_url = ""

    # 4) คลิก + คว้า PDF
    body, err = _grab_pdf_after_click(
        page,
        lambda: page.evaluate(pick_js, {"pat": pat, "mode": "click"}),
        log=log,
        prefetch_url=prefetch_url,
    )
    if err:
        return b"", matched_label, f"ดาวน์โหลด PDF ไม่สำเร็จ: {err}"
    if not body or len(body) < 500:
        return b"", matched_label, f"ไฟล์ PDF เล็กผิดปกติ ({len(body) if body else 0} bytes)"
    return body, matched_label, ""


def _bt30_ctn_read_pdf_text(pdf_bytes: bytes) -> str:
    """อ่านข้อความจาก PDF (pypdf) — คืน '' ถ้าอ่านไม่ได้"""
    if not pdf_bytes:
        return ""
    try:
        import io as _io
        from pypdf import PdfReader
        reader = PdfReader(_io.BytesIO(pdf_bytes))
        # อ่านเฉพาะ 2 หน้าแรก — พอสำหรับ header + passport
        parts = []
        for p in reader.pages[:2]:
            try:
                parts.append(p.extract_text() or "")
            except Exception:
                pass
        return "\n".join(parts)
    except Exception:
        return ""


def _bt30_ctn_download_appointment(
    page: Page, log=print,
) -> tuple[bytes, str, str, str]:
    """ดาวน์โหลด 'ใบนัดหมาย' จาก tab การนัดหมาย (สมมติ page อยู่บน detail แล้ว)

    Flow:
      1. คลิก tab การนัดหมาย (#tab_default_5)
      2. หา iframe #link_appointment (จาก queue-fe*.doe.go.th)
      3. อ่าน passport + ชื่อ จาก iframe DOM (ไม่ต้อง OCR)
      4. เปิด iframe.src ใน new tab (share context/cookies) → page.pdf() → bytes

    คืน (pdf_bytes, passport, name, error_msg)
      - pdf_bytes: bytes ของ PDF (ว่าง = ล้มเหลว)
      - passport: เลขที่หนังสือเดินทาง (จาก DOM)
      - name: ชื่อคนต่างด้าว (Eng)
      - error_msg: '' ถ้าสำเร็จ, มิฉะนั้นข้อความ error
    """
    # 1) คลิก tab การนัดหมาย
    try:
        page.evaluate(r"""() => {
            if (window.jQuery) {
                try { jQuery('a[href="#tab_default_5"]').tab('show'); } catch(e){}
            }
            const a = document.querySelector('a[href="#tab_default_5"]');
            if (a) a.click();
            // fallback: หา anchor ที่ข้อความ 'การนัดหมาย'
            const els = [...document.querySelectorAll('a,.nav-link,.nav-tabs a,.nav a,[role="tab"]')];
            const t = els.find(x => /^\s*การนัดหมาย\s*$/.test((x.textContent||'').trim()));
            if (t) t.click();
        }""")
    except Exception:
        pass
    page.wait_for_timeout(2200)

    # 2) หา iframe #link_appointment / iframe src
    iframe_src = ""
    try:
        iframe_src = page.evaluate(
            "() => (document.querySelector('#link_appointment') "
            "|| document.querySelector('iframe.iframe_show') "
            "|| document.querySelector('iframe[src*=\"queue\"]') || {}).src || ''"
        )
    except Exception:
        iframe_src = ""
    if not iframe_src:
        # ตรวจว่า pane บอก "ไม่มีการนัดหมาย" ไหม
        try:
            empty_txt = page.evaluate(
                r"""() => {
                    const p = document.querySelector('#tab_default_5');
                    if (!p) return '';
                    return (p.innerText || '').trim().slice(0, 100);
                }"""
            )
        except Exception:
            empty_txt = ""
        if "ไม่มีการนัดหมาย" in (empty_txt or ""):
            return b"", "", "", "record นี้ไม่มีการนัดหมาย (ยังไม่ได้จอง / ยกเลิกไปแล้ว)"
        return b"", "", "", "ไม่พบ iframe ใบนัดหมายบน tab การนัดหมาย"

    # 3) หา frame ที่ตรงกัน + รอ load
    appt_frame = None
    for _ in range(15):
        for fr in page.frames:
            if fr.name == "link_appointment" \
                    or "bookingdate" in (fr.url or "") \
                    or "bookingdetail" in (fr.url or ""):
                appt_frame = fr
                break
        if appt_frame:
            break
        page.wait_for_timeout(500)
    if not appt_frame:
        return b"", "", "", "iframe โหลดไม่สำเร็จ (queue system อาจล่มชั่วคราว)"
    try:
        appt_frame.wait_for_load_state("domcontentloaded", timeout=15_000)
    except Exception:
        pass
    page.wait_for_timeout(2500)

    # 4) อ่าน passport + name จาก iframe DOM
    #    Note: iframe body มี text ต่อกันไม่มี space:
    #    "...หมายเลขนัดหมายชื่อคนต่างด้าวเลขที่หนังสือเดินทาง2-CCO001122601108Miss KYI KYI NAINGMH190359ยกเลิก..."
    #    → หา passport แบบ Myanmar (M[A-Z]\d) หรือ 1-2 letters + 6-8 digits
    #    → ตรวจ 'ยังไม่มีการนัดหมาย' ก่อน (record status=AP ยังไม่จอง → หน้า booking form)
    passport = ""
    name_eng = ""
    is_booking_form = False  # หน้าแบบยังไม่จอง (ต้อง skip download)
    try:
        info = appt_frame.evaluate(r"""() => {
            const norm = s => (s || '').replace(/\s+/g, ' ').trim();
            const text = norm(document.body ? document.body.innerText : '');
            // ตรวจว่ายังไม่มีการนัดหมาย (booking form, ยังไม่ได้จอง)
            const isBooking = /ยังไม่มีการนัดหมาย|กรุณาเลือก\s*['\u2018\u2019]?สถานที่/.test(text)
                           || /เพิ่มวันนัดหมาย/.test(text);
            // Passport: หา pattern ที่ตรง — Myanmar M[A-Z]\d{6-8} มากที่สุด
            let passport = '';
            const patterns = [
                /\b(M[A-Z]\d{6,8})\b/,       // Myanmar: MH190359, MI123456
                /\b(C[A-Z]\d{6,8})\b/,       // Cambodia: CB123456
                /\b([A-Z]{1,2}\d{6,9})\b/,   // generic 1-2 letters + 6-9 digits
            ];
            for (const p of patterns) {
                const m = text.match(p);
                if (m) { passport = m[1]; break; }
            }
            // Name: pattern 'Miss/Mr/Mrs' + words in Latin
            let name = '';
            const nm = text.match(/(Mr\.?|Mrs\.?|Miss|Ms\.?|Mss\.?)\s+([A-Z][A-Za-z\.\s]{1,80}?)(?=\s*[A-Z]{1,2}\d|\s{2,}|$)/);
            if (nm) name = (nm[1] + ' ' + nm[2]).trim();
            return { passport, name, isBooking, bodyLen: text.length,
                     bodyStart: text.slice(0, 200) };
        }""")
        passport = (info or {}).get("passport", "") or ""
        name_eng = (info or {}).get("name", "") or ""
        is_booking_form = bool((info or {}).get("isBooking", False))
    except Exception as e:
        log(f"     ⚠ อ่าน passport/name จาก iframe ไม่สำเร็จ: {str(e)[:100]}")

    # ถ้าเป็นหน้า booking (ยังไม่ได้จอง) → skip ไม่ต้องดาวน์โหลด
    if is_booking_form:
        return b"", passport, name_eng, "record นี้ยังไม่ได้จองการนัดหมาย (หน้า booking form)"

    # 5) เปิด iframe URL ใน new tab → คลิก 'พิมพ์แบบฟอร์มนัดหมาย' (React จะ swap DOM
    #    เป็น 'ใบนัดหมาย' ตัวจริง — ตราครุฑ + QR + แบบฟอร์ม A4 ตาม @media print/tailwind)
    #    → page.pdf() → bytes
    ctx = page.context
    pdf_bytes = b""
    err = ""
    new_page = None
    try:
        new_page = ctx.new_page()
        # hook window.print เพื่อไม่ให้ blocking print dialog เปิดใน headed mode
        try:
            new_page.add_init_script(
                "window.print = function(){ window.__printCalled = true; };"
            )
        except Exception:
            pass
        new_page.goto(iframe_src, wait_until="networkidle", timeout=30_000)
        new_page.wait_for_timeout(3000)

        # คลิกปุ่ม 'พิมพ์แบบฟอร์มนัดหมาย' (ตัวแรก — single ไม่ใช่ 'ทั้งหมด')
        # ปุ่มนี้ทำให้ React swap DOM เป็น print template (~5x เพิ่มขึ้น)
        try:
            new_page.locator(
                "button:has-text('พิมพ์แบบฟอร์มนัดหมาย')"
            ).nth(0).click(timeout=6000)
        except Exception as _e:
            # fallback: หาปุ่มด้วย JS
            new_page.evaluate(r"""() => {
                const b = [...document.querySelectorAll('button')].find(x =>
                    /^\s*พิมพ์แบบฟอร์มนัดหมาย\s*$/.test((x.textContent||'').trim())
                );
                if (b) b.click();
            }""")
        # รอ React render + window.print ถูกเรียก
        new_page.wait_for_timeout(3500)

        try:
            new_page.emulate_media(media="print")
        except Exception:
            pass
        pdf_bytes = new_page.pdf(
            format="A4",
            print_background=True,
            margin={"top": "10mm", "bottom": "10mm", "left": "10mm", "right": "10mm"},
        )
        if not pdf_bytes or len(pdf_bytes) < 5000:
            err = f"PDF เล็กผิดปกติ ({len(pdf_bytes) if pdf_bytes else 0} bytes) — DOM swap อาจไม่สำเร็จ"
    except Exception as e:
        err = f"page.pdf() ล้มเหลว: {str(e).splitlines()[0][:200]}"
    finally:
        if new_page is not None:
            try: new_page.close()
            except Exception: pass

    return pdf_bytes, passport, name_eng, err


# ─────────────────────────────────────────────────────────────────
# โหมด: ตรวจวันว่างจอง (คิวถ่ายบัตร)
# - login → เปิด iframe ของ 1 record ที่มี booking (AP/APSS) → ดึง Bearer token
# - เรียก /branch/by-codes → รายชื่อทุกสาขา
# - เรียก /calendar/get-data/{branch}/{date}?month=X&year=Y → วันว่าง+slot เหลือ
# - รวบ Excel: (สาขา, วัน, เปิด, เต็มวัน, เหลือ, %) — พร้อม sheet สรุป
# ─────────────────────────────────────────────────────────────────
QUEUE_FE_ORIGIN = "https://queue-fe-uat.doe.go.th"
QUEUE_BE_ORIGIN = "https://queue-be-uat.doe.go.th"


def _booking_extract_token_from_iframe_src(iframe_src: str) -> str:
    """สกัด Bearer token จาก iframe src ของ #link_appointment
    รูปแบบ: https://queue-fe-uat.doe.go.th/doe/bookingdate/{TOKEN}?t=...
    """
    if not iframe_src:
        return ""
    # ตัด query string
    src = iframe_src.split("?", 1)[0]
    # เอาส่วนสุดท้ายของ path
    return src.rsplit("/", 1)[-1] if "/" in src else ""


def _booking_get_token(page: Page, log=print) -> str:
    """ค้นหา 1 record ที่มี booking iframe (AP/APSS ที่จองแล้ว) → ดึง token
    return "" ถ้าหาไม่เจอ (ต้องมี record ที่จองการนัดหมายไว้แล้วอย่างน้อย 1)
    """
    goto_tracking(page)
    apply_wa_filter(page, "", status_ids=["AP", "SS"])
    rows = collect_all_wa_rows(page, log=lambda *a: None)
    log(f"      [token] มี {len(rows)} คำขอสถานะ AP/SS")

    # priority: APSS ก่อน (จองแล้ว) → AP (อาจยังไม่จอง)
    prio = sorted(rows, key=lambda r: 0 if str(r.get("status") or "").upper() == "APSS" else 1)
    for i, row in enumerate(prio[:20], 1):  # ลอง 20 records แรก
        req_no = row.get("reqNo") or ""
        if not req_no:
            continue
        try:
            url = build_detail_url(row)
            if not url:
                continue
            page.goto(url, wait_until="domcontentloaded", timeout=25_000)
            page.wait_for_timeout(2200)
            # trigger tab นัดหมาย
            page.evaluate(r"""() => {
                if (window.jQuery) { try { jQuery('a[href="#tab_default_5"]').tab('show'); } catch(e){} }
                const a = document.querySelector('a[href="#tab_default_5"]');
                if (a) a.click();
            }""")
            page.wait_for_timeout(2200)

            iframe_src = page.evaluate(
                "() => (document.querySelector('#link_appointment') || {}).src || ''"
            ) or ""
            token = _booking_extract_token_from_iframe_src(iframe_src)
            if token and len(token) > 40:
                log(f"      [token] ✓ ได้ token จาก {req_no} (status={row.get('status')})")
                return token
        except Exception as _e:
            continue
    log("      [token] ✗ หา token ไม่เจอใน 20 records แรก")
    return ""


# สาขาที่รู้จัก — จาก probe ก่อนหน้า (fallback ถ้า API /branch/by-codes ล่ม)
_KNOWN_BRANCHES: list[dict[str, str]] = [
    {"code": "MDH-OB-M-001", "name": "ศูนย์แรกรับเข้าทำงานและสิ้นสุดการจ้าง จังหวัดมุกดาหาร"},
    {"code": "NKI-OB-L-001", "name": "ศูนย์แรกรับเข้าทำงานและสิ้นสุดการจ้าง จังหวัดหนองคาย"},
    {"code": "TAK-OB-L-001", "name": "ศูนย์แรกรับเข้าทำงานและสิ้นสุดการจ้าง จังหวัดตาก"},
    {"code": "RNG-SC-S-001", "name": "ศูนย์บริการใบอนุญาตทำงานของคนต่างด้าว จังหวัดระนอง"},
    {"code": "CPN-SC-S-001", "name": "ศูนย์บริการใบอนุญาตทำงานของคนต่างด้าว จังหวัดชุมพร"},
    {"code": "CCO-SC-S-001", "name": "ศูนย์บริการใบอนุญาตทำงานของคนต่างด้าว จังหวัดฉะเชิงเทรา"},
    {"code": "PTE-SC-M-001", "name": "ศูนย์บริการใบอนุญาตทำงานของคนต่างด้าว จังหวัดปทุมธานี"},
]


def _booking_fetch_branches(page: Page, token: str, log=print) -> list[dict[str, str]]:
    """ดึงรายชื่อสาขาทั้งประเทศ

    วิธี:
      1. เปิด iframe URL ใน tab ใหม่ (fresh session)
      2. รอ FE call /branch/by-codes อัตโนมัติ + capture response ผ่าน network listener
      3. Parse JSON → return list [{code, name}, ...]
    ถ้าล้มเหลว → fallback ไปใช้ _KNOWN_BRANCHES
    """
    # หา iframe src ล่าสุด (ต้องมี token ในนั้น — สร้างจาก token args)
    # แต่เราไม่ได้เก็บ src ไว้ — ให้เปิดจาก page.evaluate() หา iframe ปัจจุบัน (ถ้ามี)
    iframe_src = ""
    try:
        iframe_src = page.evaluate(
            "() => (document.querySelector('#link_appointment') || {}).src || ''"
        ) or ""
    except Exception:
        iframe_src = ""
    if not iframe_src:
        log("      [branches] ⚠ ไม่พบ iframe src — ใช้ list fallback")
        return list(_KNOWN_BRANCHES)

    # เปิด iframe ใน new page + capture branch response
    ctx = page.context
    branch_response: list[dict[str, Any]] = []

    def _on_response(res):
        u = res.url or ""
        if "/branch/by-codes" in u:
            try:
                data = res.json()
                if isinstance(data, list) and data:
                    branch_response.extend(data)
            except Exception:
                pass

    ctx.on("response", _on_response)
    qp = None
    try:
        qp = ctx.new_page()
        qp.goto(iframe_src, wait_until="networkidle", timeout=30_000)
        qp.wait_for_timeout(4000)  # รอ FE call APIs ให้เสร็จ
    except Exception as e:
        log(f"      [branches] ⚠ เปิด iframe fresh ไม่สำเร็จ: {str(e)[:100]}")
    finally:
        try: ctx.remove_listener("response", _on_response)
        except Exception: pass
        if qp is not None:
            try: qp.close()
            except Exception: pass

    if not branch_response:
        log("      [branches] ⚠ ไม่ได้ response จาก /branch/by-codes — ใช้ list fallback")
        return list(_KNOWN_BRANCHES)

    # unique + normalize
    seen: set[str] = set()
    result: list[dict[str, str]] = []
    for b in branch_response:
        code = (b.get("branch_code_id") or "").strip()
        if not code or code in seen:
            continue
        seen.add(code)
        result.append({
            "code": code,
            "name": (b.get("branch_name_th") or b.get("branch_name_en") or code).strip(),
        })
    log(f"      [branches] ✓ ได้ {len(result)} สาขา (unique)")
    return result if result else list(_KNOWN_BRANCHES)


def _booking_fetch_calendar(
    page: Page, token: str, branch_code: str, year: int, month: int,
    log=print,
) -> list[dict]:
    """เรียก /calendar/get-data/{branch}/{YYYY-MM-01}?month=X&year=Y
    return list ของ {date, open, maxNormal, count_left_Normal, ...}
    ถ้า error → return []
    """
    date_str = f"{year}-{month:02d}-01"
    api_url = (
        f"{QUEUE_BE_ORIGIN}/doe-booking/api/v2/calendar/get-data/"
        f"{branch_code}/{date_str}?month={month}&year={year}"
    )
    try:
        res = page.context.request.get(
            api_url,
            headers={
                "Authorization": f"Bearer {token}",
                "Origin": QUEUE_FE_ORIGIN,
                "Referer": QUEUE_FE_ORIGIN + "/",
            },
            timeout=30_000,
        )
        if not res.ok:
            return []
        parsed = res.json()
        if isinstance(parsed, dict) and isinstance(parsed.get("data"), list):
            return parsed["data"]
        return []
    except Exception:
        return []


def _booking_fetch_rounds(
    page: Page, token: str, branch_code: str, date_str: str,
    log=print,
) -> list[dict]:
    """เรียก /v1/round/get-rounds/?b_id={branch}&date_string={YYYY-MM-DD}

    return list ของ {round: "09:00 - 09:30", roundId: int, capacity: int, left: int}
    - 200 [] = เปิดวันแต่ไม่มี round data (สาขา MULTIPLE จะเป็นแบบนี้)
    - 400 = วัน/สาขาไม่มีสิทธิ์จอง → return []
    - อื่น ๆ error → return []
    """
    api_url = (
        f"{QUEUE_BE_ORIGIN}/doe-booking/api/v1/round/get-rounds/"
        f"?b_id={branch_code}&date_string={date_str}"
    )
    try:
        res = page.context.request.get(
            api_url,
            headers={
                "Authorization": f"Bearer {token}",
                "Origin": QUEUE_FE_ORIGIN,
                "Referer": QUEUE_FE_ORIGIN + "/",
                "Accept": "application/json, text/plain, */*",
            },
            timeout=15_000,
        )
        if not res.ok:
            return []
        parsed = res.json()
        if isinstance(parsed, list):
            return parsed
        return []
    except Exception:
        return []


def booking_scan_once(
    page: Page, token: str,
    branches: list[dict],
    year_month_pairs: list[tuple[int, int]],
    log=print, is_cancelled=None,
    include_full: bool = False,
    on_progress=None,
) -> list[dict]:
    """สแกน 1 รอบ: calendar + rounds ของทุกสาขา×เดือน

    Args:
        branches: [{code, name}, ...]
        year_month_pairs: [(2026, 8), (2026, 9), ...]
        include_full: True = รวมรอบที่เต็มแล้วด้วย (left=0), False = เฉพาะรอบที่มีที่ว่าง
        on_progress: callback(done, total, branch_code, branch_name, month_str) — ถ้ามี

    Returns: list of {branch_code, branch_name, date, round, capacity, left}
        เรียงตาม date → branch → round
    """
    slots: list[dict] = []
    total_iters = len(branches) * len(year_month_pairs)
    done_iters = 0
    for br in branches:
        if is_cancelled and is_cancelled():
            break
        code = br.get("code", "")
        name = br.get("name", "")
        if not code:
            continue
        for (yy, mm) in year_month_pairs:
            if is_cancelled and is_cancelled():
                break
            if on_progress:
                try:
                    on_progress(done_iters, total_iters, code, name, f"{yy}-{mm:02d}")
                except Exception:
                    pass
            days = _booking_fetch_calendar(page, token, code, yy, mm, log=log)
            for d in days:
                left_day = int(d.get("count_left_Normal") or 0)
                date_str = d.get("date", "")
                if not (d.get("open") and date_str):
                    continue
                if not include_full and left_day <= 0:
                    continue
                rounds = _booking_fetch_rounds(page, token, code, date_str, log=log)
                for rd in rounds:
                    try:
                        cap = int(rd.get("capacity") or 0)
                        r_left = int(rd.get("left") or 0)
                    except Exception:
                        cap, r_left = 0, 0
                    if not include_full and r_left <= 0:
                        continue
                    slots.append({
                        "branch_code": code,
                        "branch_name": name,
                        "date": date_str,
                        "round": str(rd.get("round") or ""),
                        "roundId": rd.get("roundId"),
                        "capacity": cap,
                        "left": r_left,
                    })
            done_iters += 1
    if on_progress:
        try:
            on_progress(done_iters, total_iters, "", "", "")
        except Exception:
            pass
    slots.sort(key=lambda s: (s["date"], s["branch_code"], s["round"]))
    return slots



def run_booking_availability(
    cfg: dict,
    login_excel: Path | None,
    out_path: Path,
    months_ahead: int = 3,
    branch_filter: list[str] | None = None,
    log=print,
    progress=None,
    is_cancelled=None,
) -> tuple[int, Path]:
    """โหมด 'ตรวจวันว่างจอง (คิวถ่ายบัตร)'

    Args:
        cfg: {username, password, user_type, method, headless, hide_window}
        login_excel: UsernameLogin.xlsx (ถ้ามี — ใช้บัญชีแรก) หรือ None (fallback cfg)
        out_path: Excel report
        months_ahead: จำนวนเดือนที่จะดูจาก 'เดือนปัจจุบัน' (เช่น 3 = เดือนนี้ + 2 เดือนถัดไป)
        branch_filter: list ของ branch_code_id ที่สนใจ — None = ทุกสาขา
    Returns: (rows_written, out_path)

    Flow:
        1. login → หา 1 record ที่มี booking iframe → ดึง Bearer token
        2. /branch/by-codes → รายชื่อทุกสาขา
        3. Loop branch × months → /calendar/get-data → เก็บ
        4. Save Excel: sheet='ตารางว่างจอง' (สาขา × วัน × slot)
    """
    from datetime import date as _date
    out_path = _timestamped_path(out_path)

    # ---- อ่านบัญชี ----
    acct: dict[str, str] | None = None
    if login_excel and Path(login_excel).exists():
        accounts = _read_login_accounts(Path(login_excel))
        if accounts:
            acct = next(iter(accounts.values()))
    if not acct:
        _u = (cfg.get("username") or "").strip()
        _p = cfg.get("password") or ""
        if _u and _p:
            acct = {
                "username": _u, "password": _p,
                "type": cfg.get("user_type") or "ผู้กระทำการแทน",
                "method": cfg.get("method") or "E-Workpermit",
            }
    if not acct:
        raise ValueError(
            "ไม่มีบัญชี login — ต้องมี UsernameLogin.xlsx หรือ กรอก Username/Password ในช่อง 'ข้อมูลเข้าสู่ระบบ'"
        )

    log(f"[1/4] Login: {acct['username']} ({acct['type']})")
    log(f"      เดือนที่ตรวจ: {months_ahead} เดือนถัดไป (จากเดือนปัจจุบัน)")

    results: list[dict[str, Any]] = []

    with sync_playwright() as pw:
        browser = _launch_chromium(pw, cfg, ["--ignore-certificate-errors", "--start-maximized"])
        ctx = browser.new_context(
            locale="th-TH", ignore_https_errors=True,
            viewport={"width": 1600, "height": 1000},
        )
        page = ctx.new_page()
        try:
            login(page, {
                "username": acct["username"], "password": acct["password"],
                "user_type": acct["type"],
                "method": acct.get("method") or cfg.get("method", "E-Workpermit"),
            })
            page.wait_for_timeout(1200)

            # 2/4: ดึง token
            log(f"[2/4] หา Bearer token จาก 1 record ที่มี booking iframe...")
            token = _booking_get_token(page, log=log)
            if not token:
                raise RuntimeError(
                    "ไม่พบ record ที่มี booking iframe — บัญชีนี้อาจไม่มีคำขอสถานะ AP/APSS"
                )

            # 3/4: fetch branches
            log(f"[3/4] ดึงรายชื่อสาขา...")
            branches = _booking_fetch_branches(page, token, log=log)
            if branch_filter:
                bf = set(b.strip().upper() for b in branch_filter if b)
                before = len(branches)
                branches = [b for b in branches if (b["code"] or "").upper() in bf]
                log(f"      filter สาขา: {len(branches)}/{before}")
            if not branches:
                raise RuntimeError("ไม่มีสาขาให้ตรวจ (branch_filter อาจไม่ตรง)")

            # 4/4: loop calendar
            today = _date.today()
            year_month_pairs: list[tuple[int, int]] = []
            y, m = today.year, today.month
            for _ in range(max(1, int(months_ahead))):
                year_month_pairs.append((y, m))
                m += 1
                if m > 12:
                    m = 1; y += 1

            total_iterations = len(branches) * len(year_month_pairs)
            done = 0
            log(f"[4/4] ตรวจ calendar × {len(branches)} สาขา × {len(year_month_pairs)} เดือน = {total_iterations} calls")
            if progress:
                try: progress(0, total_iterations)
                except Exception: pass

            for br in branches:
                if is_cancelled and is_cancelled():
                    log("[!] ผู้ใช้ยกเลิก — หยุด"); break
                code = br["code"]
                name = br["name"]
                for (yy, mm) in year_month_pairs:
                    if is_cancelled and is_cancelled():
                        break
                    days = _booking_fetch_calendar(page, token, code, yy, mm, log=log)
                    for d in days:
                        max_normal = int(d.get("maxNormal") or 0)
                        left = int(d.get("count_left_Normal") or 0)
                        date_str = d.get("date", "")
                        is_open = bool(d.get("open", False))
                        # fetch rounds เฉพาะวันที่เปิดและมีที่ว่าง (ประหยัด calls)
                        rounds: list[dict] = []
                        if is_open and left > 0 and date_str:
                            rounds = _booking_fetch_rounds(page, token, code, date_str, log=log)
                        results.append({
                            "branch_code": code,
                            "branch_name": name,
                            "year_month": f"{yy}-{mm:02d}",
                            "date": date_str,
                            "open": is_open,
                            "max_normal": max_normal,
                            "count_left_normal": left,
                            "used": max(max_normal - left, 0),
                            "left_pct": round((left / max_normal * 100), 1) if max_normal else 0,
                            "rounds": rounds,
                        })
                    done += 1
                    if done % 5 == 0:
                        log(f"      progress: {done}/{total_iterations}")
                    if progress:
                        try: progress(done, total_iterations)
                        except Exception: pass
        finally:
            ctx.close(); browser.close()

    # เขียน Excel report
    _save_booking_availability_report(results, out_path, log=log)
    return len(results), out_path


def _save_booking_availability_report(
    rows: list[dict[str, Any]], out_path: Path, log=print,
) -> None:
    """เซฟรายงาน 2 sheets:
    - 'ตารางว่างจอง' (detail): สาขา × วัน × slot
    - 'สรุปตามสาขา' (summary): สาขา × (จำนวน slot รวม, เหลือรวม, %)
    """
    wb = Workbook()

    # Sheet 1: Detail
    ws = wb.active
    ws.title = "ตารางว่างจอง"
    headers = [
        "รหัสสาขา", "ชื่อสาขา", "ปี-เดือน", "วันที่", "เปิด",
        "เต็มวัน (max)", "จองไปแล้ว", "ที่เหลือ", "% เหลือ",
    ]
    ws.append(headers)
    head_fill = PatternFill("solid", fgColor="305496")
    ok_fill = PatternFill("solid", fgColor="C6EFCE")     # เขียวอ่อน = เหลือเยอะ
    low_fill = PatternFill("solid", fgColor="FFEB9C")    # เหลืองอ่อน = เหลือน้อย
    full_fill = PatternFill("solid", fgColor="FFC7CE")   # แดงอ่อน = เต็ม
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # sort: ตามวันที่ (ใกล้สุดก่อน)
    rows_sorted = sorted(rows, key=lambda r: (r.get("date") or "", r.get("branch_code") or ""))
    for r in rows_sorted:
        ws.append([
            r["branch_code"], r["branch_name"], r["year_month"], r["date"],
            "เปิด" if r["open"] else "ปิด",
            r["max_normal"], r["used"], r["count_left_normal"],
            f"{r['left_pct']}%",
        ])
        # ไฮไลต์แถวตาม slot ที่เหลือ
        row_idx = ws.max_row
        left = r["count_left_normal"]
        max_n = r["max_normal"]
        if not r["open"]:
            pass  # ปิดไม่ไฮไลต์
        elif left == 0:
            for c in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=c).fill = full_fill
        elif max_n > 0 and left <= max_n * 0.2:  # เหลือ ≤ 20%
            for c in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=c).fill = low_fill
        elif max_n > 0 and left > 0:
            for c in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=c).fill = ok_fill

    widths = [16, 55, 12, 14, 10, 14, 14, 12, 12]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions

    # Sheet 2: สรุปตามสาขา
    ws2 = wb.create_sheet("สรุปตามสาขา")
    ws2.append([
        "รหัสสาขา", "ชื่อสาขา",
        "จำนวนวันที่เก็บได้", "วันที่เปิด", "วันที่เต็มแล้ว",
        "slot รวม (max)", "จองไปรวม", "ที่เหลือรวม", "% เหลือ",
    ])
    for cell in ws2[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # aggregate per branch
    by_branch: dict[str, dict[str, Any]] = {}
    for r in rows:
        k = r["branch_code"]
        if k not in by_branch:
            by_branch[k] = {
                "code": r["branch_code"], "name": r["branch_name"],
                "days": 0, "days_open": 0, "days_full": 0,
                "max": 0, "used": 0, "left": 0,
            }
        b = by_branch[k]
        b["days"] += 1
        if r["open"]:
            b["days_open"] += 1
        if r["open"] and r["count_left_normal"] == 0:
            b["days_full"] += 1
        b["max"] += r["max_normal"]
        b["used"] += r["used"]
        b["left"] += r["count_left_normal"]
    for b in sorted(by_branch.values(), key=lambda x: x["code"]):
        pct = round(b["left"] / b["max"] * 100, 1) if b["max"] else 0
        ws2.append([
            b["code"], b["name"],
            b["days"], b["days_open"], b["days_full"],
            b["max"], b["used"], b["left"], f"{pct}%",
        ])
    widths2 = [16, 55, 18, 14, 16, 16, 14, 14, 12]
    for i, w in enumerate(widths2, start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.freeze_panes = "A2"
    if ws2.max_row > 1:
        ws2.auto_filter.ref = ws2.dimensions

    # Sheet 3: ช่วงเวลาที่ว่าง (rounds)
    ws3 = wb.create_sheet("ช่วงเวลาที่ว่าง")
    ws3.append([
        "รหัสสาขา", "ชื่อสาขา", "วันที่", "ช่วงเวลา",
        "รอบ ID", "ที่นั่งรวม (max)", "ที่เหลือ", "% เหลือ",
    ])
    for cell in ws3[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    round_rows: list[tuple] = []
    for r in rows:
        rounds = r.get("rounds") or []
        for rd in rounds:
            try:
                cap = int(rd.get("capacity") or 0)
                left = int(rd.get("left") or 0)
            except Exception:
                cap, left = 0, 0
            pct = round(left / cap * 100, 1) if cap else 0
            round_rows.append((
                r["branch_code"], r["branch_name"], r["date"],
                str(rd.get("round") or ""), rd.get("roundId"),
                cap, left, pct,
            ))
    # sort: วันที่ → สาขา → เวลา
    round_rows.sort(key=lambda t: (t[2], t[0], t[3]))
    for tup in round_rows:
        code, name, date, round_str, rid, cap, left, pct = tup
        ws3.append([code, name, date, round_str, rid, cap, left, f"{pct}%"])
        row_idx = ws3.max_row
        if cap == 0:
            pass
        elif left == 0:
            for c in range(1, 9):
                ws3.cell(row=row_idx, column=c).fill = full_fill
        elif left <= cap * 0.2:
            for c in range(1, 9):
                ws3.cell(row=row_idx, column=c).fill = low_fill
        else:
            for c in range(1, 9):
                ws3.cell(row=row_idx, column=c).fill = ok_fill
    widths3 = [16, 55, 14, 18, 12, 16, 12, 10]
    for i, w in enumerate(widths3, start=1):
        ws3.column_dimensions[get_column_letter(i)].width = w
    ws3.freeze_panes = "A2"
    if ws3.max_row > 1:
        ws3.auto_filter.ref = ws3.dimensions

    out_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(out_path)
        log(f"      · เซฟรายงาน → {out_path}")
    except Exception as e:
        log(f"      · ⚠ เซฟรายงานไม่สำเร็จ: {str(e)[:160]}")


def run_appointment(
    cfg: dict,
    login_excel: Path,
    out_path: Path,
    request_type: str = "",
    row_range: str | None = None,
    log=print,
    progress=None,
    is_cancelled=None,
) -> tuple[int, Path]:
    """โหมด 'นัดหมายถ่ายบัตร' — เก็บที่อยู่ (จังหวัด + อำเภอ/เขต) จากใบเสร็จค่าธรรมเนียมใบอนุญาตทำงาน

    Flow:
        UsernameLogin.xlsx → login ทีละบัญชี →
        Tracking → กรองสถานะ 'รอนัดหมาย' (AP) → วนทุก pagination →
        ทำทุก req_no ที่เจอ:
            tab ข้อมูลคนต่างด้าว (เก็บชื่อสถานประกอบการ(ไทย)) →
            tab การชำระเงิน →
                Case 1: กดปุ่ม 'หลักฐานการชำระเงิน' ของ 'ค่าธรรมเนียมใบอนุญาตทำงาน' →
                Case 2: กด 'ดูใบเสร็จรับเงิน' → ค้นหาเลขคำขอ → กด record ที่ 2 →
            ดาวน์โหลด PDF → parse จังหวัด/อำเภอ → เซฟรายงาน (hyperlink PDF)
    """
    out_path = _timestamped_path(out_path)
    accounts = _read_login_accounts(login_excel)
    log(f"[1/3] บัญชี login จาก {Path(login_excel).name}: {len(accounts)} บัญชี")
    log("      รายการคำขอ: ทั้งหมด | สถานะ: AP (รอนัดหมาย)")

    save_dir = out_path.parent / "appointment_receipts"
    save_dir.mkdir(parents=True, exist_ok=True)

    results: list[dict[str, Any]] = []
    success = 0
    total_known = 0

    if progress:
        try: progress(0, 1)
        except Exception: pass

    with sync_playwright() as pw:
        browser = _launch_chromium(pw, cfg, ["--ignore-certificate-errors", "--start-maximized"])
        ctx = browser.new_context(
            locale="th-TH", ignore_https_errors=True,
            viewport={"width": 1920, "height": 1080},
        )
        page = ctx.new_page()
        try:
            for u_idx, (ukey, acct) in enumerate(accounts.items(), start=1):
                if is_cancelled and is_cancelled():
                    log("[!] ผู้ใช้ยกเลิก — หยุด")
                    break
                username = acct["username"]
                log(f"  [บัญชี {u_idx}/{len(accounts)}] === Login: {username} ({acct['type']}) ===")
                try:
                    if u_idx > 1:
                        _logout_safely(page)
                        page.wait_for_timeout(1000)
                    login(page, {
                        "username": acct["username"], "password": acct["password"],
                        "user_type": acct["type"],
                        "method": acct.get("method") or cfg.get("method", "E-Workpermit"),
                    })
                    goto_tracking(page)
                    # โหมดนี้ force 'รายการคำขอ = ทั้งหมด' เสมอ — ไม่สนใจ cfg.request_type
                    apply_wa_filter(page, "", status_ids=["AP"])
                except Exception as e:
                    log(f"      ⛔ login/เปิด tracking ไม่สำเร็จ: {str(e).splitlines()[0][:160]}")
                    results.append({
                        "req_no": "", "username": username,
                        "status": "LOGIN_FAIL", "error": str(e).splitlines()[0][:200],
                        "company_th": "", "province": "", "district": "",
                        "pdf_file": "", "case": "",
                    })
                    continue

                try:
                    rows = collect_all_wa_rows(page, log=log)
                except Exception as e:
                    log(f"      ✗ เก็บรายการคำขอไม่สำเร็จ: {str(e).splitlines()[0][:160]}")
                    continue

                # apply_wa_filter ได้กรอง AP server-side แล้ว — ใช้ทั้งหมดที่ collect มา
                # แต่กันพลาด: เก็บเฉพาะที่ statusText มี 'รอนัดหมาย' ถ้ามีอย่างน้อย 1 รายการ match
                ap_rows = [r for r in rows if "รอนัดหมาย" in str(r.get("statusText", ""))]
                if not ap_rows:
                    ap_rows = rows
                log(f"      พบ {len(ap_rows)} คำขอสถานะ 'รอนัดหมาย'")

                # row_range: ใช้กับลำดับ AP-rows ของบัญชีนี้
                indices = _parse_row_range(row_range, len(ap_rows)) if row_range else list(range(1, len(ap_rows) + 1))
                selected_rows = [ap_rows[i - 1] for i in indices]
                total_known += len(selected_rows)

                for k, row in enumerate(selected_rows, start=1):
                    if is_cancelled and is_cancelled():
                        break
                    req_no = row.get("reqNo", "") or row.get("req_no", "")
                    if not req_no:
                        continue
                    log(f"    ({k}/{len(selected_rows)}) คำขอ {req_no} — {row.get('statusText','')}")
                    rec_in = {"req_no": req_no, "username": username, "_row": row}
                    try:
                        r = _appt_process_record(page, rec_in, save_dir, log=log)
                        r["username"] = username
                        results.append(r)
                        if r.get("status") == "SUCCESS":
                            success += 1
                    except Exception as e:
                        results.append({
                            "req_no": req_no, "username": username,
                            "status": "ERROR", "error": str(e).splitlines()[0][:200],
                            "company_th": "", "province": "", "district": "",
                            "pdf_file": "", "case": "",
                        })
                        log(f"      ✗ ผิดพลาด: {str(e).splitlines()[0][:160]}")
                    if progress:
                        try: progress(len(results), max(total_known, len(results)))
                        except Exception: pass
                    _save_appointment_report(results, out_path, log=lambda *a: None)
        finally:
            ctx.close(); browser.close()

    _save_appointment_report(results, out_path, log=log)
    fail = sum(1 for r in results if r.get("status") not in ("SUCCESS",))
    log(f"[3/3] สรุป: เก็บที่อยู่สำเร็จ {success} คำขอ (ล้มเหลว/empty {fail})")
    return success, out_path


# ─────────────────────────────────────────────────────────────────
# run_bt30_ctn — main entry สำหรับโหมด 'ดาวน์โหลด บต.30 (ใบตอบรับจากระบบ)'
# ─────────────────────────────────────────────────────────────────
def _save_bt30_ctn_report(rows: list[dict[str, Any]], out_path: Path, log=print) -> None:
    """เซฟรายงาน — คอลัมน์ 'ไฟล์ CTN' และ 'ไฟล์ APPOINTMENT' เป็น hyperlink (file://) ไปยัง reports/bt30_ctn/"""
    wb = Workbook()
    ws = wb.active
    ws.title = "บต.30 CTN + Appointment"
    headers = [
        "ลำดับ", "Username", "เลขที่คำขอ", "Passport", "รูปแบบ (BT..)",
        "ป้ายในระบบ", "สถานะ", "ไฟล์ CTN", "ไฟล์ บต.25",
        "ไฟล์ APPOINTMENT", "ไฟล์ Receipt", "หมายเหตุ",
    ]
    ws.append(headers)
    head_fill = PatternFill("solid", fgColor="305496")
    link_font = Font(color="0563C1", underline="single")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    text_cols = {"เลขที่คำขอ", "Passport"}
    link_cols = {"ไฟล์ CTN", "ไฟล์ บต.25", "ไฟล์ APPOINTMENT", "ไฟล์ Receipt"}
    save_subdir = "bt30_ctn"
    for i, r in enumerate(rows, start=1):
        ws.append([
            i,
            r.get("username", ""),
            r.get("req_no", ""),
            r.get("passport", ""),
            r.get("form_code", ""),
            r.get("label", ""),
            r.get("status", ""),
            r.get("pdf_file", ""),
            r.get("bt25_file", ""),
            r.get("appt_file", ""),
            r.get("receipt_file", ""),
            r.get("error", ""),
        ])
        row_idx = ws.max_row
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=row_idx, column=c_idx)
            if h in text_cols and cell.value not in (None, ""):
                cell.value = str(cell.value)
                cell.number_format = "@"
            if h in link_cols and cell.value:
                fname = str(cell.value)
                abs_path = (out_path.parent / save_subdir / fname).resolve()
                try:
                    cell.hyperlink = abs_path.as_uri()
                except Exception:
                    cell.hyperlink = f"{save_subdir}/{fname}"
                cell.value = fname
                cell.font = link_font

    widths = [6, 28, 20, 18, 12, 40, 12, 45, 45, 45, 45, 40]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions

    out_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(out_path)
        log(f"      · เซฟรายงาน → {out_path}")
    except Exception as e:
        log(f"      · ⚠ เซฟรายงานไม่สำเร็จ: {str(e)[:160]}")


def _bt30_ctn_process_record(
    page: Page, rec: dict, save_dir: Path, make_subfolder: bool = False,
    do_ctn: bool = True, do_appointment: bool = False,
    do_bt25: bool = False, do_receipt: bool = False, log=print,
) -> dict[str, Any]:
    """1 คำขอ: เปิด detail → ดาวน์โหลดเอกสารที่เลือก (บต.30 / บต.25 / ใบนัดหมาย / ใบเสร็จค่าธรรมเนียม)
    → บันทึกเป็น {PASSPORT}_BT{XX}_CTN.pdf / {PASSPORT}_BT25.pdf /
      {PASSPORT}_APPOINTMENT.pdf / {PASSPORT}_RECEIPT.pdf
    คืน dict สำหรับใส่ report
    """
    req_no = rec.get("req_no", "")
    row = rec.get("_row", {}) or {}
    out = {
        "req_no": req_no,
        "username": rec.get("username", ""),
        "passport": "",
        "form_code": "",
        "label": "",
        "pdf_file": "",
        "bt25_file": "",
        "appt_file": "",
        "receipt_file": "",
        "status": "FAIL",
        "error": "",
    }
    # 1) เปิด detail (reuse ของ appointment mode)
    ok, err = _appt_open_detail(page, req_no, log=log)
    if not ok:
        out["error"] = err
        return out

    passport = ""
    form_num = ""
    ctn_ok = False
    bt25_ok = False
    appt_ok = False
    receipt_ok = False
    ctn_err = ""
    bt25_err = ""
    appt_err = ""
    receipt_err = ""

    # --- 2A) ดาวน์โหลด ใบตอบรับ (บต.XX) ---
    if do_ctn:
        body, matched_label, derr = _bt30_ctn_download(page, log=log)
        out["label"] = matched_label
        if derr:
            ctn_err = derr
        else:
            pdf_text = _bt30_ctn_read_pdf_text(body)
            passport = _extract_passport_from_pdf(pdf_text)
            form_num = _extract_bt_form_number_from_pdf(pdf_text) or "30"
            out["passport"] = passport
            out["form_code"] = f"BT{form_num}" if form_num else ""

            prefix = passport if passport else (req_no or "unknown")
            fname = f"{prefix}_BT{form_num}_CTN.pdf" if form_num else f"{prefix}_BT_CTN.pdf"
            target_dir = save_dir / prefix if make_subfolder else save_dir
            target_dir.mkdir(parents=True, exist_ok=True)
            target_path = target_dir / fname
            if target_path.exists():
                n = 2
                stem = target_path.stem
                while True:
                    alt = target_dir / f"{stem}_{n}.pdf"
                    if not alt.exists():
                        target_path = alt
                        break
                    n += 1
            try:
                target_path.write_bytes(body)
                try:
                    rel = target_path.relative_to(save_dir).as_posix()
                except Exception:
                    rel = target_path.name
                out["pdf_file"] = rel
                ctn_ok = True
                log(f"      · CTN: {rel}  (passport={passport or '-'}, form=BT{form_num or '?'})")
            except Exception as e:
                ctn_err = f"บันทึก PDF ล้มเหลว: {str(e)[:160]}"

    # --- 2A2) ดาวน์โหลด แบบ บต.25 (เอกสารตอบรับจากระบบ) ---
    if do_bt25:
        b25, lbl25, e25 = _bt30_ctn_download(
            page, log=log, pat=r"บต\.?\s*25", doc_name="บต.25",
        )
        if not out["label"]:
            out["label"] = lbl25
        if e25:
            bt25_err = e25
        else:
            if not passport:
                passport = _extract_passport_from_pdf(_bt30_ctn_read_pdf_text(b25))
                if passport:
                    out["passport"] = passport
            prefix = passport if passport else (req_no or "unknown")
            target_dir = save_dir / prefix if make_subfolder else save_dir
            target_dir.mkdir(parents=True, exist_ok=True)
            target_path = _uniq_pdf_path(target_dir / f"{prefix}_BT25.pdf")
            try:
                target_path.write_bytes(b25)
                try:
                    rel = target_path.relative_to(save_dir).as_posix()
                except Exception:
                    rel = target_path.name
                out["bt25_file"] = rel
                bt25_ok = True
                log(f"      · BT25: {rel}  (passport={passport or '-'})")
            except Exception as e:
                bt25_err = f"บันทึก บต.25 PDF ล้มเหลว: {str(e)[:160]}"

    # --- 2B) ดาวน์โหลด ใบนัดหมาย (APPOINTMENT) ---
    if do_appointment:
        appt_bytes, appt_pp, appt_name, aerr = _bt30_ctn_download_appointment(page, log=log)
        if aerr:
            appt_err = aerr
        elif appt_bytes:
            # ใช้ passport จาก CTN ก่อน — ถ้าไม่มี ใช้จาก iframe
            appt_prefix = passport or appt_pp or (req_no or "unknown")
            if not out["passport"] and appt_pp:
                out["passport"] = appt_pp
            appt_fname = f"{appt_prefix}_APPOINTMENT.pdf"
            target_dir = save_dir / appt_prefix if make_subfolder else save_dir
            target_dir.mkdir(parents=True, exist_ok=True)
            appt_path = target_dir / appt_fname
            if appt_path.exists():
                n = 2
                stem = appt_path.stem
                while True:
                    alt = target_dir / f"{stem}_{n}.pdf"
                    if not alt.exists():
                        appt_path = alt
                        break
                    n += 1
            try:
                appt_path.write_bytes(appt_bytes)
                try:
                    rel = appt_path.relative_to(save_dir).as_posix()
                except Exception:
                    rel = appt_path.name
                out["appt_file"] = rel
                appt_ok = True
                log(f"      · APPOINTMENT: {rel}  (passport={appt_prefix}, name={appt_name or '-'})")
            except Exception as e:
                appt_err = f"บันทึก APPOINTMENT PDF ล้มเหลว: {str(e)[:160]}"

    # --- 2C) ดาวน์โหลด Receipt ค่าธรรมเนียมใบอนุญาตทำงาน (แท็บการชำระเงิน) ---
    #     ทำเป็นขั้นสุดท้าย เพราะ case2 จะ navigate ออกไปหน้า MultiplePayments
    if do_receipt:
        pay_case = ""
        try:
            _appt_open_payment_tab(page)
            pay_case = _appt_detect_payment_case(page)
        except Exception as e:
            receipt_err = f"เปิดแท็บการชำระเงินไม่สำเร็จ: {str(e)[:120]}"
        if not receipt_err:
            if pay_case == "empty":
                receipt_err = "ไม่มี Record การชำระเงิน (ยังไม่ชำระ / ไม่มีใบเสร็จ)"
            else:
                rbody = b""; ramount = ""; rkind = ""; rerr = ""
                if pay_case == "case1":
                    rbody, ramount, rkind, rerr = _bt30_ctn_grab_receipt_case1(page, log=log)
                else:
                    rbody, _rfname_tmp, rerr = _appt_case2_download(page, req_no, save_dir, log=log)
                    if rbody:
                        ramount = _extract_pdf_amount(rbody)
                        rkind = _bt30_ctn_receipt_fee_kind(rbody)
                        # ลบไฟล์ชั่วคราวของ case2 (จะบันทึกใหม่ด้วยชื่อที่มียอดเงิน)
                        try:
                            if _rfname_tmp:
                                (save_dir / _rfname_tmp).unlink()
                        except Exception:
                            pass
                if rerr or not rbody:
                    receipt_err = rerr or "ดาวน์โหลดใบเสร็จไม่สำเร็จ"
                else:
                    if not passport:
                        passport = _extract_passport_from_pdf(_bt30_ctn_read_pdf_text(rbody))
                        if passport and not out["passport"]:
                            out["passport"] = passport
                    prefix = passport if passport else (req_no or "unknown")
                    target_dir = save_dir / prefix if make_subfolder else save_dir
                    target_dir.mkdir(parents=True, exist_ok=True)
                    amt_tag = ramount or ""
                    final_path = _uniq_pdf_path(target_dir / f"{prefix}_RECEIPT{amt_tag}.pdf")
                    try:
                        final_path.write_bytes(rbody)
                        try:
                            rel = final_path.relative_to(save_dir).as_posix()
                        except Exception:
                            rel = final_path.name
                        out["receipt_file"] = rel
                        receipt_ok = True
                        _warn = "" if rkind == "workpermit" else f"  ⚠(ประเภท={rkind or 'unknown'})"
                        log(f"      · RECEIPT: {rel}  (case={pay_case}, ยอด={amt_tag or '-'}, passport={prefix}){_warn}")
                    except Exception as e:
                        receipt_err = f"บันทึกใบเสร็จล้มเหลว: {str(e)[:160]}"

    # --- 3) สรุปสถานะ ---
    wanted = []
    if do_ctn: wanted.append(("ctn", ctn_ok, ctn_err))
    if do_bt25: wanted.append(("bt25", bt25_ok, bt25_err))
    if do_appointment: wanted.append(("appt", appt_ok, appt_err))
    if do_receipt: wanted.append(("receipt", receipt_ok, receipt_err))
    n_ok = sum(1 for _, ok, _ in wanted if ok)
    if n_ok == len(wanted) and wanted:
        out["status"] = "SUCCESS"
    elif n_ok > 0:
        out["status"] = "PARTIAL"
    else:
        # ไม่มีอะไรสำเร็จ → เช็คว่าเป็น "ไม่มีเอกสาร" ทุกอย่างไหม
        errs = [e for _, _, e in wanted if e]
        no_doc_markers = ("ไม่พบ", "ไม่มี", "ยังไม่ได้จอง", "ยังไม่ได้")
        if all(any(m in e for m in no_doc_markers) for e in errs if e):
            out["status"] = "NO_DOC"
        else:
            out["status"] = "FAIL"
    err_parts = []
    if ctn_err:  err_parts.append(f"CTN: {ctn_err}")
    if bt25_err: err_parts.append(f"BT25: {bt25_err}")
    if appt_err: err_parts.append(f"APPT: {appt_err}")
    if receipt_err: err_parts.append(f"RECEIPT: {receipt_err}")
    if err_parts:
        out["error"] = " | ".join(err_parts)
    if do_ctn and ctn_ok and not passport:
        # download success แต่หา passport ไม่เจอ → mark PARTIAL
        if out["status"] == "SUCCESS":
            out["status"] = "PARTIAL"
        if not out["error"]:
            out["error"] = "ดาวน์โหลด PDF สำเร็จ แต่หา passport ในไฟล์ไม่พบ"
    return out


def run_bt30_ctn(
    cfg: dict,
    login_excel: Path | None,
    out_path: Path,
    row_range: str | None = None,
    make_subfolder: bool = False,
    do_ctn: bool = True,
    do_appointment: bool = False,
    do_bt25: bool = False,
    do_receipt: bool = False,
    log=print,
    progress=None,
    is_cancelled=None,
) -> tuple[int, Path]:
    """โหมด 'ดาวน์โหลด (ใบตอบรับ / ใบนัดหมาย)' — วนทุกบัญชี → filter สถานะ → ทุกคำขอ

    Args:
        do_ctn: True = ดาวน์โหลด 'ใบตอบรับ บต.30' จาก tab เอกสารตอบรับ (สถานะ WP2)
        do_bt25: True = ดาวน์โหลด 'แบบ บต.25' จาก tab เอกสารตอบรับ
        do_appointment: True = ดาวน์โหลด 'ใบนัดหมาย' จาก tab การนัดหมาย (สถานะ AP/APSS)
        do_receipt: True = ดาวน์โหลด 'ใบเสร็จค่าธรรมเนียมใบอนุญาตทำงาน' จาก tab การชำระเงิน
        (เลือกได้หลายอย่างพร้อมกัน → รันในลูปเดียว แต่ควรเลือก status filter ให้ครอบคลุม)

    Flow:
        UsernameLogin.xlsx (หรือบัญชีเดียวจาก cfg) → login ทีละบัญชี →
        Tracking → กรอง status=WP (รอชำระเงิน) [+ request_types + date filter ถ้ามีใน cfg] →
        ทุก req_no ที่เจอ:
            เปิด detail → tab 'เอกสารตอบรับจากระบบ' →
            หา row 'แบบ บต.30' → ดาวน์โหลด PDF →
            อ่าน passport + form number จาก PDF text →
            เซฟเป็น reports/bt30_ctn/{PASSPORT}_BT{XX}_CTN.pdf

    Args:
        cfg: ตัวเลือกรวม (headless, hide_window, request_types, date_from, date_to,
             username/password/user_type — ใช้เป็น single-user fallback)
        login_excel: UsernameLogin.xlsx (Username, Password, Type) — None หรือไฟล์ไม่มี/อ่านไม่ได้
                     → fallback ไปใช้ cfg['username']+cfg['password'] (บัญชีเดียว)
        out_path: ไฟล์ Excel report
        row_range: '1-10,15' → ทำเฉพาะ row ที่เลือกในแต่ละบัญชี (ว่าง = ทั้งหมด)
        make_subfolder: True → เซฟใน reports/bt30_ctn/{PASSPORT}/{PASSPORT}_BT30_CTN.pdf

    Returns: (count_success, out_path)
    """
    if not (do_ctn or do_appointment or do_bt25 or do_receipt):
        raise ValueError(
            "ต้องเลือกอย่างน้อย 1 เอกสาร (บต.30 / บต.25 / ใบนัดหมาย / ใบเสร็จ) ก่อนรันโหมดนี้"
        )
    out_path = _timestamped_path(out_path)

    # ---- อ่าน accounts จาก UsernameLogin.xlsx หรือ fallback ไป single-user จาก cfg ----
    accounts: dict[str, dict[str, str]] = {}
    _src_desc = ""
    if login_excel and Path(login_excel).exists():
        try:
            accounts = _read_login_accounts(Path(login_excel))
            _src_desc = f"ไฟล์ {Path(login_excel).name}"
        except Exception as e:
            log(f"      ⚠ อ่านไฟล์ {Path(login_excel).name} ไม่ได้: {str(e).splitlines()[0][:160]}")
            log("        · จะลอง fallback ไปใช้ Username/Password จากช่องด้านบนแทน")
            accounts = {}
    # fallback: บัญชีเดียวจาก cfg
    if not accounts:
        _u = (cfg.get("username") or "").strip()
        _p = cfg.get("password") or ""
        if _u and _p:
            accounts = {_u.lower(): {
                "username": _u,
                "password": _p,
                "type": (cfg.get("user_type") or "ผู้กระทำการแทน"),
                "method": (cfg.get("method") or "E-Workpermit"),
            }}
            _src_desc = f"Username/Password ด้านบน (บัญชีเดียว: {_u})"
    if not accounts:
        raise ValueError(
            "ไม่มีบัญชี login — ต้องมีไฟล์ UsernameLogin.xlsx (มีข้อมูลอย่างน้อย 1 แถว) "
            "หรือกรอก Username/Password ในช่อง 'ข้อมูลเข้าสู่ระบบ' ด้านบน"
        )
    log(f"[1/3] บัญชี login จาก {_src_desc}: {len(accounts)} บัญชี")

    # ---- Status filter: ถ้า UI ระบุ filter_status_ids มา → ใช้ตามนั้น (ไม่ล็อก WP)
    # ถ้าไม่ระบุ → default = WP + local filter เฉพาะ WP2 (พฤติกรรมเดิม สำหรับใบตอบรับ บต.30 MoU)
    _user_status_ids = list(cfg.get("filter_status_ids") or [])
    _user_status_ids = [s for s in _user_status_ids if s]  # ตัด empty
    if _user_status_ids:
        status_ids_for_filter = _user_status_ids
        wp2_only = False
        log(f"      สถานะ: {','.join(status_ids_for_filter)} (จาก UI) | "
            f"Tab: เอกสารตอบรับจากระบบ | เอกสาร: แบบ บต.30")
    else:
        status_ids_for_filter = ["WP"]
        wp2_only = True
        log("      สถานะ: WP (รอชำระเงิน) — filter WP2 เฉพาะ | "
            "Tab: เอกสารตอบรับจากระบบ | เอกสาร: แบบ บต.30")

    # request_types (multi-select) + date filter — เอามาจาก cfg ที่ GUI ส่งเข้ามา
    req_types_list: list[str] = []
    for c in (cfg.get("request_types") or []):
        s = (c or "").strip()
        if s and s not in {"0", "ALL", "all"} and s not in req_types_list:
            req_types_list.append(s)
    if not req_types_list:
        single = (cfg.get("request_type") or "").strip()
        if single and single not in {"0", "ALL", "all"}:
            req_types_list = [single]
    if req_types_list:
        log(f"      ฟิลเตอร์รายการคำขอ: {', '.join(req_types_list) if len(req_types_list) <= 3 else str(len(req_types_list)) + ' รายการ'}")
    date_from = (cfg.get("date_from") or "").strip()
    date_to = (cfg.get("date_to") or "").strip()
    if date_from or date_to:
        log(f"      วันที่ยื่นคำขอ: {date_from or '(ต้นสุด)'} → {date_to or '(ล่าสุด)'}")

    save_dir = out_path.parent / "bt30_ctn"
    save_dir.mkdir(parents=True, exist_ok=True)

    results: list[dict[str, Any]] = []
    success = 0
    total_known = 0

    if progress:
        try: progress(0, 1)
        except Exception: pass

    with sync_playwright() as pw:
        browser = _launch_chromium(pw, cfg, ["--ignore-certificate-errors", "--start-maximized"])
        ctx = browser.new_context(
            locale="th-TH", ignore_https_errors=True,
            viewport={"width": 1920, "height": 1080},
        )
        page = ctx.new_page()
        try:
            for u_idx, (ukey, acct) in enumerate(accounts.items(), start=1):
                if is_cancelled and is_cancelled():
                    log("[!] ผู้ใช้ยกเลิก — หยุด")
                    break
                username = acct["username"]
                log(f"  [บัญชี {u_idx}/{len(accounts)}] === Login: {username} ({acct['type']}) ===")
                try:
                    if u_idx > 1:
                        _logout_safely(page)
                        page.wait_for_timeout(1000)
                    login(page, {
                        "username": acct["username"], "password": acct["password"],
                        "user_type": acct["type"],
                        "method": acct.get("method") or cfg.get("method", "E-Workpermit"),
                        "login_timeout_ms": cfg.get("login_timeout_ms", 30_000),
                    })
                    goto_tracking(page)
                except Exception as e:
                    log(f"      ⛔ login/เปิด tracking ไม่สำเร็จ: {str(e).splitlines()[0][:160]}")
                    results.append({
                        "req_no": "", "username": username,
                        "status": "LOGIN_FAIL", "error": str(e).splitlines()[0][:200],
                        "passport": "", "form_code": "", "label": "", "pdf_file": "",
                    })
                    continue

                # collect rows (ต่อ 1 บัญชี) — วน request_types ถ้ามีหลาย
                all_rows: list[dict] = []
                seen_req: set = set()
                try:
                    if not req_types_list:
                        apply_wa_filter(
                            page, "", status_ids=status_ids_for_filter,
                            date_from=date_from, date_to=date_to,
                        )
                        all_rows = collect_all_wa_rows(page, log=log)
                    else:
                        for rt_i, rt in enumerate(req_types_list, 1):
                            if is_cancelled and is_cancelled():
                                break
                            if len(req_types_list) > 1:
                                log(f"      [{rt_i}/{len(req_types_list)}] filter: {rt}")
                            apply_wa_filter(
                                page, rt, status_ids=status_ids_for_filter,
                                date_from=date_from, date_to=date_to,
                            )
                            chunk = collect_all_wa_rows(page, log=log)
                            for r in chunk:
                                k = str(r.get("reqNo") or "")
                                if k and k in seen_req:
                                    continue
                                if k:
                                    seen_req.add(k)
                                all_rows.append(r)
                except Exception as e:
                    log(f"      ✗ เก็บรายการคำขอไม่สำเร็จ: {str(e).splitlines()[0][:160]}")
                    continue

                # ---- Local filter: ถ้าใช้ default WP → คัดเฉพาะ WP2 (BT30 อยู่แค่ WP2)
                # ถ้า UI ระบุสถานะเอง → ไม่กรองต่อ (เชื่อ user)
                _before = len(all_rows)
                if wp2_only:
                    wp_rows = [
                        r for r in all_rows
                        if str(r.get("status") or "").upper() == "WP2"
                    ]
                    _target_label = "'รอชำระเงิน' (WP2)"
                else:
                    wp_rows = list(all_rows)
                    _target_label = f"({','.join(status_ids_for_filter)})"
                if _before == 0:
                    _hint = "" if not req_types_list else \
                            f" (filter รายการคำขอ: {', '.join(req_types_list)})"
                    log(f"      ⚠ บัญชีนี้ไม่มีคำขอในสถานะ {','.join(status_ids_for_filter)}{_hint}")
                    log("        · ตรวจ badge สถานะบนหน้าเว็บว่ามีจริงไหม")
                    log("        · หรือลองล้าง filter รายการคำขอ (เลือก 'ล้าง' ในกล่อง Listbox) แล้วรันใหม่")
                else:
                    _sub_counts: dict[str, int] = {}
                    for r in all_rows:
                        s = str(r.get("status") or "?").upper()
                        _sub_counts[s] = _sub_counts.get(s, 0) + 1
                    _sub_txt = ", ".join(f"{k}={v}" for k, v in sorted(_sub_counts.items()))
                    if wp2_only:
                        log(
                            f"      สถานะย่อยที่เจอ: {_sub_txt}  → คัดเฉพาะ WP2: "
                            f"{len(wp_rows)}/{_before}"
                        )
                    else:
                        log(f"      สถานะย่อยที่เจอ: {_sub_txt}  (ใช้ทั้งหมด: {len(wp_rows)})")
                log(f"      พบ {len(wp_rows)} คำขอสถานะ {_target_label}")

                # row_range: ต่อบัญชี
                indices = _parse_row_range(row_range, len(wp_rows)) if row_range else list(range(1, len(wp_rows) + 1))
                selected_rows = [wp_rows[i - 1] for i in indices if 1 <= i <= len(wp_rows)]
                total_known += len(selected_rows)

                for k, row in enumerate(selected_rows, start=1):
                    if is_cancelled and is_cancelled():
                        break
                    req_no = row.get("reqNo", "") or row.get("req_no", "")
                    if not req_no:
                        continue
                    log(f"    ({k}/{len(selected_rows)}) คำขอ {req_no} — {row.get('statusText','')}")
                    rec_in = {"req_no": req_no, "username": username, "_row": row}
                    try:
                        r = _bt30_ctn_process_record(
                            page, rec_in, save_dir, make_subfolder=make_subfolder,
                            do_ctn=do_ctn, do_appointment=do_appointment,
                            do_bt25=do_bt25, do_receipt=do_receipt,
                            log=log,
                        )
                        r["username"] = username
                        results.append(r)
                        if r.get("status") == "SUCCESS":
                            success += 1
                    except Exception as e:
                        results.append({
                            "req_no": req_no, "username": username,
                            "status": "ERROR", "error": str(e).splitlines()[0][:200],
                            "passport": "", "form_code": "", "label": "", "pdf_file": "",
                            "bt25_file": "", "appt_file": "", "receipt_file": "",
                        })
                        log(f"      ✗ ผิดพลาด: {str(e).splitlines()[0][:160]}")
                    if progress:
                        try: progress(len(results), max(total_known, len(results)))
                        except Exception: pass
                    _save_bt30_ctn_report(results, out_path, log=lambda *a: None)
        finally:
            ctx.close(); browser.close()

    _save_bt30_ctn_report(results, out_path, log=log)
    fail = sum(1 for r in results if r.get("status") not in ("SUCCESS",))
    log(f"[3/3] สรุป: ดาวน์โหลด บต.30 สำเร็จ {success} คำขอ (ล้มเหลว/partial {fail})")
    return success, out_path


# ─────────────────────────────────────────────────────────────────
# run_namelist_alien — ดึงรายชื่อคนต่างด้าวจากหน้า NameListAlien
# URL: /Requtst63_2/NameListAlien?form_type=<FORM_TYPE>
# DataTable server-side (id=AlienRenewFormCrList_table) — วน pagination
# ทุกหน้า (100 rows/หน้า) → เขียนออก Excel
# ─────────────────────────────────────────────────────────────────
NAMELIST_ALIEN_URL_TMPL = (
    "https://eworkpermit.doe.go.th/Requtst63_2/NameListAlien?form_type={form_type}"
)
NAMELIST_ALIEN_TABLE_ID = "AlienRenewFormCrList_table"


def _parse_namelist_alien_cells(cells: list[str]) -> dict[str, str]:
    """แยกข้อมูลจาก text ของแต่ละ cell ในตาราง NameListAlien
    cells: list ของ innerText 7 cell (index 0..6)
        [0] ลำดับ
        [1] "RA... \n เลขประจำตัวคนต่างด้าว : XXX \n เลขที่ใบอนุญาตทำงาน : XXX"
        [2] "Mr. NAME (Eng) \n Thai name \n เพศ : ..."
        [3] ผลตรวจ+ประกัน (badges + note)
        [4] สัญชาติ
        [5] "รอยื่นคำขอ..." หรือ "ยื่นคำขอแล้ว \n เลขที่คำขอ : 69125200710969"
        [6] ผู้ยื่นคำขอ (name หรือ '-')
    """
    out = {
        "ref_no": "",
        "alien_id": "",
        "work_permit_no": "",
        "name_eng": "",
        "name_th": "",
        "gender": "",
        "health_status": "",
        "nationality": "",
        "request_status": "",
        "request_no": "",
        "submitter": "",
    }
    if not cells:
        return out
    # cell[1] — reference/id block
    c1 = cells[1] if len(cells) > 1 else ""
    lines = [ln.strip() for ln in re.split(r"[\r\n]+", c1) if ln.strip()]
    for ln in lines:
        if ln.startswith("RA") and not out["ref_no"]:
            out["ref_no"] = ln.strip()
        elif "เลขประจำตัวคนต่างด้าว" in ln:
            m = re.search(r":\s*(\S+)", ln)
            if m:
                out["alien_id"] = m.group(1)
        elif "เลขที่ใบอนุญาตทำงาน" in ln:
            m = re.search(r":\s*(\S+)", ln)
            if m:
                out["work_permit_no"] = m.group(1)
    # cell[2] — name/gender block
    c2 = cells[2] if len(cells) > 2 else ""
    lines2 = [ln.strip() for ln in re.split(r"[\r\n]+", c2) if ln.strip()]
    for ln in lines2:
        if re.match(r"^(Mr|Mrs|Miss|Ms|Master|Mstr)\.?\s+", ln, flags=re.IGNORECASE):
            out["name_eng"] = ln
        elif re.match(r"^(นาย|นาง|นางสาว|เด็กชาย|เด็กหญิง)\s", ln):
            out["name_th"] = ln
        elif "เพศ" in ln:
            m = re.search(r":\s*(\S+)", ln)
            if m:
                out["gender"] = m.group(1)
    # cell[3] — health status (เก็บทั้งบล็อค — badges + note)
    if len(cells) > 3:
        out["health_status"] = re.sub(r"\s+", " ", cells[3]).strip()[:400]
    # cell[4] — nationality
    if len(cells) > 4:
        out["nationality"] = cells[4].strip()
    # cell[5] — request status + req_no
    c5 = cells[5] if len(cells) > 5 else ""
    lines5 = [ln.strip() for ln in re.split(r"[\r\n]+", c5) if ln.strip()]
    status_lines = []
    for ln in lines5:
        if "เลขที่คำขอ" in ln:
            m = re.search(r":\s*(\S+)", ln)
            if m:
                out["request_no"] = m.group(1)
        else:
            status_lines.append(ln)
    out["request_status"] = " ".join(status_lines)[:200]
    # cell[6] — submitter
    if len(cells) > 6:
        s = cells[6].strip()
        out["submitter"] = "" if s == "-" else s
    return out


def _namelist_alien_collect_all(
    page: Page, log=print, limit: int = 0, is_cancelled=None,
) -> list[dict[str, str]]:
    """ตั้ง page length = 100 → วน pagination ทุกหน้า → parse rows
    - limit > 0: หยุดทันทีเมื่อ collected ครบจำนวน
    - is_cancelled(): callable → หยุดถ้าคืน True
    คืน list ของ dict (parsed)
    """
    tid = NAMELIST_ALIEN_TABLE_ID
    # ตั้ง page length = 100 + ไปหน้าแรก
    page.evaluate(
        """(tid) => {
            try {
                const dt = window.jQuery('#' + tid).DataTable();
                dt.page.len(100).page(0).draw('page');
            } catch(e) {}
        }""", tid,
    )
    info = _wait_datatable_idle(page, tid, max_wait_ms=45_000)
    pages = int(info.get("pages") or 1)
    total = int(info.get("recordsDisplay") or 0)
    length = int(info.get("length") or 100)
    log(f"      [paginate] ทั้งหมด {total:,} แถว / {pages:,} หน้า ({length} แถว/หน้า)")
    if limit and limit > 0:
        log(f"      [paginate] จำกัด: {limit:,} แถวแรก")

    all_rows: list[dict[str, str]] = []
    seen: set = set()
    cur_page = 0
    while True:
        if is_cancelled and is_cancelled():
            log("      [paginate] ยกเลิกโดยผู้ใช้")
            break
        rows_now = page.evaluate(
            r"""(tid) => {
                const out = [];
                document.querySelectorAll('#' + tid + ' tbody tr').forEach(tr => {
                    const cells = Array.from(tr.querySelectorAll('td'))
                        .map(td => (td.innerText || '').trim());
                    if (cells.length) out.push(cells);
                });
                return out;
            }""", tid,
        )
        added = 0
        for cells in rows_now:
            parsed = _parse_namelist_alien_cells(cells)
            key = parsed.get("ref_no") or (parsed.get("alien_id") + "|" + parsed.get("name_eng"))
            if key in seen:
                continue
            seen.add(key)
            all_rows.append(parsed)
            added += 1
            if limit and limit > 0 and len(all_rows) >= limit:
                break
        log(f"      [paginate] หน้า {cur_page + 1}/{pages} → +{added} (สะสม {len(all_rows):,})")

        # ครบ limit → stop
        if limit and limit > 0 and len(all_rows) >= limit:
            log(f"      [paginate] ถึง limit ({limit:,}) — หยุด")
            break
        if cur_page >= pages - 1:
            break
        # go next
        page.evaluate(
            """(tid) => {
                try { window.jQuery('#' + tid).DataTable().page('next').draw('page'); } catch(e) {}
            }""", tid,
        )
        info = _wait_datatable_idle(page, tid, max_wait_ms=45_000)
        new_page = int(info.get("page") or -1)
        if new_page == cur_page:
            log("      [paginate] ไม่ขยับหน้า — หยุด")
            break
        cur_page = new_page
        if cur_page > pages + 5:
            log("      [paginate] เกินหน้าสุดท้าย — หยุด")
            break
    return all_rows


def _save_namelist_alien_report(rows: list[dict[str, str]], out_path: Path, log=print) -> None:
    """เซฟรายงาน — 12 คอลัมน์"""
    wb = Workbook()
    ws = wb.active
    ws.title = "รายชื่อคนต่างด้าว"
    headers = [
        "ลำดับ", "หมายเลขอ้างอิง", "เลขประจำตัวคนต่างด้าว", "เลขที่ใบอนุญาตทำงาน",
        "ชื่อ (Eng)", "ชื่อ (ไทย)", "เพศ",
        "ผลตรวจ+ประกัน", "สัญชาติ", "สถานะคำขอ", "เลขที่คำขอ", "ผู้ยื่นคำขอ",
    ]
    ws.append(headers)
    head_fill = PatternFill("solid", fgColor="305496")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    text_cols = {"เลขประจำตัวคนต่างด้าว", "เลขที่ใบอนุญาตทำงาน", "เลขที่คำขอ", "หมายเลขอ้างอิง"}
    for i, r in enumerate(rows, start=1):
        ws.append([
            i,
            r.get("ref_no", ""),
            r.get("alien_id", ""),
            r.get("work_permit_no", ""),
            r.get("name_eng", ""),
            r.get("name_th", ""),
            r.get("gender", ""),
            r.get("health_status", ""),
            r.get("nationality", ""),
            r.get("request_status", ""),
            r.get("request_no", ""),
            r.get("submitter", ""),
        ])
        row_idx = ws.max_row
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=row_idx, column=c_idx)
            if h in text_cols and cell.value not in (None, ""):
                cell.value = str(cell.value)
                cell.number_format = "@"
            if h == "ผลตรวจ+ประกัน" and cell.value:
                cell.alignment = Alignment(wrap_text=True, vertical="top")

    widths = [6, 22, 18, 18, 30, 30, 8, 45, 12, 22, 18, 22]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions

    out_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(out_path)
        log(f"      · เซฟรายงาน → {out_path}")
    except Exception as e:
        log(f"      · ⚠ เซฟรายงานไม่สำเร็จ: {str(e)[:160]}")


def run_namelist_alien(
    cfg: dict,
    out_path: Path,
    form_type: str = "MT_63_2_3103_RENEWAL",
    limit: int = 0,
    log=print,
    progress=None,
    is_cancelled=None,
) -> tuple[int, Path]:
    """ดึงรายชื่อคนต่างด้าวจากหน้า NameListAlien (ทุก pagination) → บันทึก Excel

    Flow:
        1) Login (บัญชีเดียวจาก cfg — ไม่ต้องใช้ UsernameLogin.xlsx)
        2) goto /Requtst63_2/NameListAlien?form_type=<form_type>
        3) DataTable set len=100 + วนทุก page
        4) ทุก row: parse text ในแต่ละ cell → เขียน Excel

    Args:
        cfg: {username, password, user_type, method, headless, hide_window}
        out_path: ไฟล์ Excel
        form_type: default MT_63_2_3103_RENEWAL (ตาม URL ผู้ใช้ระบุ)
        limit: 0 = ทั้งหมด, >0 = จำกัดแถวแรก N (ประหยัดเวลาตอนทดสอบ)

    Returns: (count, out_path)
    """
    out_path = _timestamped_path(out_path)
    log(f"[1/3] URL: {NAMELIST_ALIEN_URL_TMPL.format(form_type=form_type)}")
    log(f"      ไฟล์รายงาน: {out_path.name}")

    def _cancelled() -> bool:
        return bool(is_cancelled and is_cancelled())

    if progress:
        try: progress(0, 1)
        except Exception: pass

    with sync_playwright() as pw:
        browser = _launch_chromium(pw, cfg, ["--disable-blink-features=AutomationControlled"])
        ctx = browser.new_context(
            locale="th-TH", timezone_id="Asia/Bangkok",
            viewport={"width": 1600, "height": 1000},
        )
        page = ctx.new_page()
        try:
            log("[2/3] กำลังเข้าสู่ระบบ...")
            login(page, cfg)
            log(f"      ล็อกอินสำเร็จ ({page.url})")

            url = NAMELIST_ALIEN_URL_TMPL.format(form_type=form_type)
            log(f"      เปิดหน้ารายชื่อ...")
            page.goto(url, wait_until="domcontentloaded", timeout=45_000)
            page.wait_for_timeout(3500)
            if _is_logged_out(page):
                log("      ⚠ session หมด — login ใหม่")
                login(page, cfg)
                page.goto(url, wait_until="domcontentloaded", timeout=45_000)
                page.wait_for_timeout(3500)

            # ตรวจว่าตารางโหลด
            try:
                page.wait_for_selector(f"#{NAMELIST_ALIEN_TABLE_ID}", state="visible", timeout=20_000)
            except Exception:
                raise RuntimeError(f"ไม่พบตาราง {NAMELIST_ALIEN_TABLE_ID} — เปิดหน้าไม่สำเร็จ")

            all_rows = _namelist_alien_collect_all(
                page, log=log, limit=limit, is_cancelled=is_cancelled,
            )
            if limit and limit > 0 and len(all_rows) > limit:
                all_rows = all_rows[:limit]
                log(f"      จำกัดเฉพาะ {limit:,} แถวแรก")

            if progress:
                try: progress(len(all_rows), len(all_rows))
                except Exception: pass

            log(f"[3/3] บันทึก Excel {len(all_rows):,} แถว...")
            _save_namelist_alien_report(all_rows, out_path, log=log)
            return len(all_rows), out_path
        finally:
            ctx.close(); browser.close()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0, help="0 = ทั้งหมด, >0 = จำกัด N รายการ")
    ap.add_argument("--out", type=Path, default=REPORTS_DIR / "WA_report.xlsx")
    ap.add_argument(
        "--request-type", default=None,
        help="รหัสรายการคำขอ เช่น MT_59_MOU_RENEWAL (default ใช้จาก .env)",
    )
    ap.add_argument(
        "--mode", choices=["etracking", "aliens", "register", "receipts", "results", "inform", "bt30", "bt44", "bill_payment", "payment_receipts", "appointment"], default="etracking",
        help="etracking=ดึงจาก e-Tracking, aliens=ดึงรายชื่อคนต่างด้าว, register=ลงทะเบียน, receipts=ดาวน์โหลดใบเสร็จ, results=ดาวน์โหลดใบแจ้งผล/ใบรับคำขอ, inform=แจ้งการจ้างคนต่างด้าวเข้าทำงาน (บต.52), bt30=ยื่นต่ออายุใบอนุญาตทำงานตาม MoU (แบบ บต.30), bt44=แจ้งการทำงาน/เปลี่ยนรายการในใบอนุญาต ซึ่งไม่กระทบ (แบบ บต.44), bill_payment=ดาวน์โหลดใบแจ้งชำระเงิน (รอจ่ายค่าธรรมเนียม), payment_receipts=ดาวน์โหลดใบเสร็จรับเงินทั้งชุด (900+100+อื่นๆ) ตามเลขคำขอ, appointment=นัดหมายถ่ายบัตร — เก็บที่อยู่ (จังหวัด+อำเภอ/เขต) จากใบเสร็จค่าธรรมเนียมใบอนุญาตทำงาน",
    )
    ap.add_argument(
        "--sub-tabs", default=None,
        help="(โหมด aliens) ชื่อ sub-tab คั่นด้วย | (default = ทั้งหมด)",
    )
    ap.add_argument(
        "--excel-input", type=Path, default=ROOT / "Exm.xlsx",
        help="(โหมด register) ไฟล์ Excel ข้อมูลคนต่างด้าว",
    )
    ap.add_argument(
        "--request-excel", type=Path, default=ROOT / "RequestData.xlsx",
        help="(โหมด receipts) ไฟล์ Excel เลขคำขอ + Username",
    )
    ap.add_argument(
        "--login-excel", type=Path, default=ROOT / "UsernameLogin.xlsx",
        help="(โหมด receipts) ไฟล์ Excel Username/Password/Type",
    )
    ap.add_argument(
        "--multi-login", action="store_true",
        help="(โหมด etracking) ดึงจากหลายบัญชีใน --login-excel (วนทุก Username/Password/Type)",
    )
    ap.add_argument(
        "--row-range", default=None,
        help="(โหมด register) ช่วงแถวที่จะทำ เช่น '1-10,15,20-25' (default=ทั้งหมด)",
    )
    ap.add_argument(
        "--request-types", default=None,
        help="(โหมด bill_payment) รหัสรายการคำขอที่จะกรอง คั่นด้วย , เช่น 'MT_59_MOU_RENEWAL,MT_63_RENEWAL' (default=ทั้งหมด)",
    )
    ap.add_argument(
        "--doc-types", default=None,
        help="(โหมด receipts) ประเภทเอกสารที่จะดาวน์โหลด คั่นด้วย , เช่น 'receipt,bt44,bt22,bt53,bt56' หรือ 'all' (default=ทั้งหมด)",
    )
    ap.add_argument(
        "--result-doc-types", default=None,
        help="(โหมด results) ประเภทเอกสาร คั่นด้วย , เช่น 'result_notice,request_receipt' หรือ 'all' (default=ทั้งหมด)",
    )
    ap.add_argument(
        "--result-status-ids", default=None,
        help="(โหมด results) status ids ที่จะดึง คั่นด้วย , — default 'AP,SS' (AP=รอนัดหมาย, SS=ดำเนินการเสร็จสิ้น). ตัวเลือก: WP,WCOSNA,WA,AP,SS",
    )
    ap.add_argument(
        "--ref-excel", type=Path, default=ROOT / "Ref_number.xlsx",
        help="(โหมด results) ไฟล์ Excel เลขคำขอ (Ref_number) + user สำหรับค้นหา-ดาวน์โหลดตามเลขคำขอ",
    )
    ap.add_argument(
        "--inform-excel", type=Path, default=ROOT / "FormRequestEmployment.xlsx",
        help="(โหมด inform) ไฟล์ Excel ข้อมูลแจ้งจ้างคนต่างด้าว",
    )
    ap.add_argument(
        "--bt30-excel", type=Path, default=ROOT / "from_bt30.xlsx",
        help="(โหมด bt30) ไฟล์ Excel ข้อมูลคนต่างด้าว (No., คำนำหน้า, ชื่อ, สัญชาติ, เพศ, วันเกิด)",
    )
    ap.add_argument(
        "--bt30-step1-only", action="store_true",
        help="(โหมด bt30) ทำเฉพาะขั้นตอนที่ 1 (เพิ่ม/บันทึกคนต่างด้าว) — ไม่ทำขั้นตอนที่ 2 (ที่อยู่/เอกสาร/ตม./ถัดไป)",
    )
    ap.add_argument(
        "--bt44-excel", type=Path, default=ROOT / "from_bt44.xlxs.xlsx",
        help="(โหมด bt44) ไฟล์ Excel ข้อมูลคนต่างด้าว (No., คำนำหน้า, หมายเลขอ้างอิง, ชื่อ, สัญชาติ, เพศ, เกิดวันที่)",
    )
    ap.add_argument(
        "--bt44-dry-run", action="store_true",
        help="(โหมด bt44) 'ทดลองยื่น' — หยุดก่อนส่งคำขอจริง (ตามค่า --bt44-dry-stop)",
    )
    ap.add_argument(
        "--bt44-dry-stop", choices=["step2", "step3"], default="step2",
        help="(โหมด bt44, ใช้คู่กับ --bt44-dry-run) ระดับการหยุด: "
             "step2=หลังตรวจเลขใบอนุญาต (Step 2.3) | "
             "step3=หลังเลือกนายจ้าง+ประเภทกิจการ+งาน (Step 3.3-3.4) ก่อนแนบเอกสาร",
    )
    ap.add_argument(
        "--bt44-skip-doc-check", action="store_true",
        help="(โหมด bt44) ข้ามการตรวจไฟล์แนบ Step 4 ก่อนรัน — สำหรับทดสอบ flow เท่านั้น",
    )
    ap.add_argument(
        "--request-receipt-excel", type=Path, default=ROOT / "Request_Receipt.xlsx",
        help="(โหมด payment_receipts) ไฟล์ Excel เลขคำขอ + Username",
    )
    ap.add_argument(
        "--commit", action="store_true",
        help="(โหมด inform) ส่งคำขอจริง (กดข้อ 7+8) — ค่า default จะหยุดก่อนยืนยัน",
    )
    ap.add_argument(
        "--list-request-types", action="store_true",
        help="พิมพ์ list option 'รายการคำขอ' จาก dropdown ในหน้า e-Tracking แล้วจบโปรแกรม",
    )
    args = ap.parse_args()

    # โหมดเหล่านี้ใช้บัญชี login จากไฟล์ Excel (--login-excel) ไม่ต้องบังคับ .env
    _excel_login_modes = {"receipts", "results", "bt30", "bt44", "bill_payment", "payment_receipts", "appointment"}
    cfg = load_config(require_login=args.mode not in _excel_login_modes)
    if args.request_type is not None:
        cfg["request_type"] = args.request_type
    if args.list_request_types:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=False)
            ctx = browser.new_context(locale="th-TH", ignore_https_errors=True)
            page = ctx.new_page()
            try:
                login(page, cfg)
                page.wait_for_timeout(1500)
                goto_tracking(page)
                page.wait_for_timeout(2500)
                selects = page.evaluate(
                    r"""() => {
                      const out = [];
                      document.querySelectorAll('select').forEach(s => {
                        const opts = Array.from(s.options).map(o => ({value: o.value, text: (o.textContent||'').trim()}));
                        out.push({id: s.id, name: s.name, count: opts.length, opts});
                      });
                      return out;
                    }"""
                )
                import json as _json
                print(_json.dumps(selects, ensure_ascii=False, indent=2))
            finally:
                ctx.close(); browser.close()
        return 0
    if args.mode == "aliens":
        sub_tabs = [s.strip() for s in args.sub_tabs.split("|")] if args.sub_tabs else None
        out = args.out
        if str(out).endswith("WA_report.xlsx"):
            out = REPORTS_DIR / "WA_aliens.xlsx"
        run_aliens_scrape(cfg, out, sub_tabs=sub_tabs, log=print)
    elif args.mode == "register":
        out = args.out
        if str(out).endswith("WA_report.xlsx"):
            out = REPORTS_DIR / "WA_register_report.xlsx"
        run_register(cfg, args.excel_input, out, row_range=args.row_range, log=print)
    elif args.mode == "receipts":
        out = args.out
        if str(out).endswith("WA_report.xlsx"):
            out = REPORTS_DIR / "WA_receipts_report.xlsx"
        if args.doc_types:
            raw = args.doc_types.strip().lower()
            if raw in ("all", "*", ""):
                doc_types = list(DOC_TYPES)
            else:
                doc_types = [s.strip() for s in raw.split(",") if s.strip()]
                unknown = [d for d in doc_types if d not in DOC_TYPES]
                if unknown:
                    print(f"[!] ไม่รู้จัก doc-types: {unknown} (เลือกได้: {list(DOC_TYPES)})")
                    return 2
        else:
            doc_types = list(DOC_TYPES_DEFAULT)
        run_receipts(
            cfg, args.request_excel, args.login_excel, out,
            row_range=args.row_range, doc_types=doc_types, log=print,
        )
    elif args.mode == "results":
        out = args.out
        if str(out).endswith("WA_report.xlsx"):
            out = REPORTS_DIR / "WA_result_docs_report.xlsx"
        if args.result_doc_types:
            raw = args.result_doc_types.strip().lower()
            if raw in ("all", "*", ""):
                result_keys = list(RESULT_DOC_TYPES)
            else:
                result_keys = [s.strip() for s in raw.split(",") if s.strip()]
                unknown = [d for d in result_keys if d not in RESULT_DOC_TYPES]
                if unknown:
                    print(f"[!] ไม่รู้จัก result-doc-types: {unknown} (เลือกได้: {list(RESULT_DOC_TYPES)})")
                    return 2
        else:
            result_keys = list(RESULT_DOC_TYPES_DEFAULT)
        # ปรับ status filter (default AP+SS) — CLI ทับค่า default ได้
        if args.result_status_ids:
            raw_sids = args.result_status_ids.strip().upper()
            if raw_sids in ("ALL", "*"):
                cfg["result_status_ids"] = ["WP", "WCOSNA", "WA", "AP", "SS"]
            else:
                sids = [s.strip().upper() for s in raw_sids.split(",") if s.strip()]
                if sids:
                    cfg["result_status_ids"] = sids
        # ถ้ามีไฟล์ Ref_number.xlsx → ค้นหา-ดาวน์โหลดตามเลขคำขอ
        # ถ้าไม่มี → fallback ดึงจาก e-Tracking (สถานะรอนัดหมาย) แบบเดิม
        if args.ref_excel and Path(args.ref_excel).exists():
            print(f"[i] โหมด results: ค้นหาตามเลขคำขอจาก {Path(args.ref_excel).name}")
            run_result_docs_by_ref(
                cfg, args.ref_excel, args.login_excel, out,
                row_range=args.row_range, doc_keys=result_keys, log=print,
            )
        else:
            print(f"[i] โหมด results: ดึงจาก e-Tracking (ไม่พบไฟล์ {Path(args.ref_excel).name})")
            run_result_docs(
                cfg, args.login_excel, out,
                request_type=cfg.get("request_type", ""),
                doc_keys=result_keys, log=print,
            )
    elif args.mode == "inform":
        out = args.out
        if str(out).endswith("WA_report.xlsx"):
            out = REPORTS_DIR / "WA_inform_report.xlsx"
        run_inform_employer(
            cfg, args.inform_excel, args.login_excel, out,
            row_range=args.row_range, commit=args.commit, log=print,
        )
    elif args.mode == "bt30":
        out = args.out
        if str(out).endswith("WA_report.xlsx"):
            out = REPORTS_DIR / "WA_bt30_report.xlsx"
        run_bt30(
            cfg, args.bt30_excel, args.login_excel, out,
            row_range=args.row_range, do_step2=not args.bt30_step1_only, log=print,
        )
    elif args.mode == "bt44":
        out = args.out
        if str(out).endswith("WA_report.xlsx"):
            out = REPORTS_DIR / "WA_bt44_report.xlsx"
        run_bt44(
            cfg, args.bt44_excel, args.login_excel, out,
            row_range=args.row_range, dry_run=args.bt44_dry_run,
            dry_stop_at=args.bt44_dry_stop,
            check_docs=not args.bt44_skip_doc_check,
            log=print,
        )
    elif args.mode == "bill_payment":
        out = args.out
        if str(out).endswith("WA_report.xlsx"):
            out = REPORTS_DIR / "WA_bill_payment_report.xlsx"
        _bp_types = None
        if args.request_types:
            _bp_types = [s.strip() for s in args.request_types.split(",") if s.strip()]
        run_bill_payment(
            cfg, args.login_excel, out,
            row_range=args.row_range, request_types=_bp_types, log=print,
        )
    elif args.mode == "payment_receipts":
        out = args.out
        if str(out).endswith("WA_report.xlsx"):
            out = REPORTS_DIR / "WA_payment_receipts_report.xlsx"
        run_payment_receipts(
            cfg, args.request_receipt_excel, args.login_excel, out,
            row_range=args.row_range, log=print,
        )
    elif args.mode == "appointment":
        out = args.out
        if str(out).endswith("WA_report.xlsx"):
            out = REPORTS_DIR / "WA_appointment_report.xlsx"
        run_appointment(
            cfg, args.login_excel, out,
            request_type=cfg.get("request_type", ""),
            row_range=args.row_range, log=print,
        )
    else:
        if args.multi_login:
            run_scrape_multi(
                cfg, args.login_excel, args.out, limit=args.limit, log=print,
            )
        else:
            run_scrape(cfg, args.out, limit=args.limit, log=print)
    return 0


if __name__ == "__main__":
    sys.exit(main())
