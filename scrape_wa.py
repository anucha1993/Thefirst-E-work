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
    """สร้าง URL หน้า detail ตาม form_type — ตามตรรกะของ openDetail()"""
    ft = row.get("form_type", "")
    endpoint = FORM_TYPE_ENDPOINTS.get(ft)
    if not endpoint:
        # fallback — ลองใช้ endpoint ของ MT_41_4_59 ไปก่อน
        endpoint = "/RequestForm41/DetailRequest41"
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


def load_config() -> dict:
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
    if not cfg["username"] or not cfg["password"]:
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
    page.wait_for_url(lambda u: "/Login" not in u, timeout=30_000)
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
        apply_wa_filter(page, request_type or "", status_ids=status_ids)
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
    if ("/Login" in url) or url.rstrip("/").endswith("doe.go.th"):
        return True
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


def apply_wa_filter(page: Page, request_type: str = "", status_ids: list[str] | None = None) -> None:
    """ติ๊ก checkbox สถานะ (อาจหลายตัว) + ตั้งค่ารายการคำขอ แล้ว trigger รีเฟรช

    Args:
        request_type: รหัสรายการคำขอ (เช่น MT_13_EXIT, MT_59_MOU_RENEWAL)
        status_ids: list ของ checkbox id ที่ต้องติ๊ก (เช่น ["WA"], ["WCOSNA","WA"])
                    ถ้า None → ใช้ profile ของ request_type
    """
    if status_ids is None:
        status_ids = list(get_profile(request_type).get("filter_status_ids") or ["WA"])

    # ขั้นที่ 1: ติ๊ก status + เซ็ต dropdown แล้วยิง GetData ครั้งแรก
    page.evaluate(
        """({rt, ids}) => {
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
            if (rt) {
                const sel = document.getElementById('Filter_request_list');
                if (sel) {
                    sel.value = rt;
                    if (window.jQuery) window.jQuery(sel).val(rt).trigger('change');
                }
            }
            if (typeof GetDataRequestFormEtrackingForAlien === 'function') {
                GetDataRequestFormEtrackingForAlien();
            }
        }""",
        {"rt": request_type or "", "ids": status_ids},
    )
    page.wait_for_timeout(3500)
    # ขั้นที่ 2: ยิงอีกรอบเพื่อกัน race กับ select2
    page.evaluate(
        """({rt, ids}) => {
            ids.forEach(id => {
                const cb = document.getElementById(id);
                if (cb) cb.checked = true;
            });
            if (rt) {
                const sel = document.getElementById('Filter_request_list');
                if (sel && window.jQuery) window.jQuery(sel).val(rt).trigger('change');
            }
            if (typeof GetDataRequestFormEtrackingForAlien === 'function') {
                GetDataRequestFormEtrackingForAlien();
            }
        }""",
        {"rt": request_type or "", "ids": status_ids},
    )
    page.wait_for_timeout(3500)


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


def scrape_detail(page: Page, row: dict, cfg: dict | None = None, log=print, capture_extra_notes: bool = True) -> dict:
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


def save_excel(rows: list[dict], out_path: Path, fast: bool = False) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "WA Summary"

    base_cols = [
        "ลำดับ",
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
) -> tuple[int, Path]:
    """รัน scraping ทั้ง pipeline. ใช้ได้ทั้ง CLI/GUI

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
            req_type = (cfg.get("request_type") or "").strip()
            if req_type and req_type not in {"0", "ALL", "all"}:
                log(f"      ฟิลเตอร์รายการคำขอ: {req_type}")
            else:
                req_type = ""
            profile = get_profile(req_type)
            # Override จาก cfg (GUI) ถ้ามี
            status_ids = cfg.get("filter_status_ids") or profile.get("filter_status_ids") or ["WA"]
            capture_extra = cfg.get("capture_extra_notes")
            if capture_extra is None:
                capture_extra = profile.get("capture_extra_notes", True)
            log(f"      ติ๊ก checkbox สถานะ: {', '.join(status_ids)}")
            apply_wa_filter(page, req_type, status_ids=status_ids)
            # เก็บ rows ทุกหน้าผ่าน DataTables pagination (server cap ~1000/req)
            rows = collect_all_wa_rows(page, log=log)
            # ถ้าเก็บไม่ได้ + session หลุด → ฟื้นแล้วลองอีกครั้ง
            if not rows and _is_logged_out(page):
                if ensure_tracking_ready(page, cfg, req_type, log=log):
                    rows = collect_wa_rows(page)
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
                if not chunk_rows:
                    return None
                p = _chunk_path(out_path, chunk_index) if CHUNK_SIZE else out_path
                save_excel(chunk_rows, p, fast=not final)
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
                        )
                        # ถ้า session relogin ไม่ผ่าน → ลองฟื้น tracking แล้วเรียกใหม่ 1 ครั้ง
                        if "session:relogin_failed" in (detail.get("scrape_errors") or ""):
                            if ensure_tracking_ready(page, cfg, req_type, log=log):
                                log("      ↻ ลองดึงรายการนี้อีกครั้งหลังฟื้น session...")
                                detail = scrape_detail(
                                    page, row, cfg=cfg, log=log,
                                    capture_extra_notes=capture_extra,
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
) -> tuple[int, Path]:
    """โหมด e-Tracking หลายบัญชี — วน login ทุกบัญชีใน UsernameLogin.xlsx
    แล้วเรียก run_scrape ต่อบัญชี (แต่ละบัญชีได้ checkpoint/resume + แยกไฟล์ของตัวเอง)

    Args:
        cfg: ตัวเลือกรวม (headless, request_type, filter_status_ids, ฯลฯ) — ไม่ต้องมี username/password
        login_excel: ไฟล์ UsernameLogin.xlsx (คอลัมน์ Username, Password, Type)
        out_path: ไฟล์ฐาน — จะถูกแตกเป็น {stem}_{username}{suffix} ต่อบัญชี

    Returns: (จำนวนแถวรวมทุกบัญชี, โฟลเดอร์ผลลัพธ์)
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
    total_rows = 0
    produced: list[Path] = []
    log(f"[Multi] เริ่มดึง e-Tracking จาก {n} บัญชี (ไฟล์ login: {login_excel.name})")
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
    try:
        return page.evaluate(
            r"""() => {
              for (const el of document.querySelectorAll('.swal2-popup, .swal2-container')) {
                if (el.offsetParent === null) continue;
                const t = el.querySelector('.swal2-title');
                const c = el.querySelector('.swal2-html-container, .swal2-content');
                return ((t?t.textContent.trim():'') + ' | ' + (c?c.textContent.trim():'')).slice(0, 600);
              }
              for (const el of document.querySelectorAll('.modal.show, .modal[style*="display: block"]')) {
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


def _strip_title(name: str) -> str:
    """ตัดคำนำหน้า (Miss/Mr/นางสาว ฯลฯ) ออกจากชื่อ"""
    return _TITLE_PREFIX_RE.sub("", (name or "").strip()).strip()


def _read_request_data(path: Path) -> list[dict[str, Any]]:
    """อ่าน RequestData.xlsx → [{seq, name_eng, req_no, username, passport, row_index}]
    คอลัมน์: ลำดับ, ชื่อคนต่างด้าว(Eng), เลขที่คำขอ, Username, PASSPORT NUMBER
    """
    wb = load_workbook(path, data_only=True)
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
    retry หลายรอบเผื่อแท็บโหลดช้า — คืน '' ถ้าดึงไม่ได้ (อย่า fallback ไปชื่อ login)
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
    for attempt in range(4):  # retry เผื่อแท็บ/ข้อมูลยังโหลดไม่เสร็จ
        page.evaluate(
            r"""() => {
                const f = [...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a')]
                    .find(a => /ข้อมูลคนต่างด้าว/.test(a.innerText || ''));
                if (f) f.click();
            }"""
        )
        page.wait_for_timeout(1500 if attempt == 0 else 1200)
        try:
            name = page.evaluate(read_js)
        except Exception:
            name = ""
        name = (name or "").strip()
        if name:
            return name
    return ""


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
    """อ่านยอดเงินรวมจากไฟล์ PDF ใบเสร็จ → คืนเป็นสตริง เช่น '100', '1800', '1350'
    (ว่าง = หาไม่เจอ). ใช้ตัวเลขที่มีทศนิยม 2 ตำแหน่ง (รูปแบบเงินบาท) แล้วเลือกค่ามากสุด
    เพราะ 'รวมเป็นเงินทั้งสิ้น' จะเป็นยอดที่ใหญ่ที่สุดในใบเสร็จ
    """
    try:
        import io
        from pypdf import PdfReader

        reader = PdfReader(io.BytesIO(pdf_bytes))
        text = ""
        for pg in reader.pages:
            text += (pg.extract_text() or "") + "\n"
    except Exception:
        return ""
    # จับเฉพาะตัวเลขที่มี .NN (จำนวนเงิน) เช่น 100.00 / 1,800.00 — กันเลขคำขอ/พาสปอร์ตปน
    amounts = re.findall(r"\d{1,3}(?:,\d{3})*\.\d{2}", text)
    vals: list[float] = []
    for a in amounts:
        try:
            v = float(a.replace(",", ""))
        except ValueError:
            continue
        if v > 0:
            vals.append(v)
    if not vals:
        return ""
    top = max(vals)
    if top == int(top):
        return str(int(top))
    return f"{top:.2f}".rstrip("0").rstrip(".").replace(".", "_")


def _download_all_receipts(
    page: Page,
    receipts_dir: Path,
    passport_safe: str,
    log=print,
) -> list[dict[str, str]]:
    """ดาวน์โหลด 'หลักฐานการชำระเงิน' ทุกใบในแท็บการชำระเงิน
    คืน list ของ {amount, file, status, error} (1 รายการต่อ 1 ใบเสร็จ)
    ตั้งชื่อไฟล์ตามราคาที่อ่านได้จากใน PDF: {PASSPORT}_RECEIPT{ราคา}.pdf
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
            base = f"{passport_safe}_RECEIPT{amount}"
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
                 if not DOC_TYPES[dt].get("multi")}

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
    log=print,
    progress=None,
    is_cancelled=None,
) -> tuple[int, Path]:
    """ดาวน์โหลดเอกสาร (ใบเสร็จ 400 / บต.44 / บต.22) ตามเลขคำขอใน RequestData.xlsx
    - จัดกลุ่มตาม Username → login ครั้งเดียว/บัญชี
    - รองรับ session timeout: login ใหม่ด้วยบัญชีเดิมแล้วทำต่อ
    - resume ได้: ข้ามไฟล์ PDF ที่มีอยู่แล้ว (แยกตามประเภท)
    - doc_types: รายการคีย์จาก DOC_TYPES (default = ทั้งหมด)
    """
    doc_types = [d for d in (doc_types or DOC_TYPES_DEFAULT) if d in DOC_TYPES]
    if not doc_types:
        raise ValueError("ต้องเลือกประเภทเอกสารอย่างน้อย 1 อย่าง")
    out_path = _timestamped_path(out_path)
    records = _read_request_data(request_excel)
    accounts = _read_login_accounts(login_excel)
    total = len(records)
    indices = _parse_row_range(row_range, total)
    selected = [records[i - 1] for i in indices]
    log(f"[1/3] อ่าน RequestData: {request_excel} ({total} แถว) → จะทำ {len(selected)} แถว: {row_range or 'ทั้งหมด'}")
    log(f"      บัญชี login: {login_excel} ({len(accounts)} บัญชี)")
    log(f"      เอกสารที่จะดาวน์โหลด: {', '.join(DOC_TYPES[d]['label'] for d in doc_types)}")

    receipts_dir = out_path.parent / "receipts"
    receipts_dir.mkdir(parents=True, exist_ok=True)

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
) -> dict[str, Any]:
    """ค้นหา 1 เลขคำขอ → เปิด detail → ดาวน์โหลดเอกสารตามที่เลือก
    มี session recovery: ถ้าหลุด login ระหว่างทาง → login ใหม่ด้วยบัญชีเดิม แล้วลองอีกครั้ง
    """
    seq = rec.get("seq", "")
    req_no = rec.get("req_no", "")
    name_excel = rec.get("name_eng", "")
    passport = rec.get("passport", "") or seq  # fallback = ลำดับ
    passport_safe = _receipt_safe_name(passport)

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
        if cfg_dt.get("multi"):
            existing = sorted(receipts_dir.glob(f"{passport_safe}_RECEIPT*.pdf"))
            if existing:
                docs_state[dt]["status"] = "SKIP_EXISTS"
                docs_state[dt]["pdf_file"] = ", ".join(p.name for p in existing)
                log(f"     ↷ {cfg_dt['label']}: มีไฟล์แล้ว ({len(existing)} ใบ) — ข้าม")
            else:
                todo.append(dt)
            continue
        out_pdf = receipts_dir / f"{passport_safe}_{cfg_dt['suffix']}.pdf"
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
            if cfg_dt.get("multi"):
                rec_results = _download_all_receipts(page, receipts_dir, passport_safe, log=log)
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
            out_pdf = receipts_dir / f"{passport_safe}_{cfg_dt['suffix']}.pdf"
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
    if "detailrequest" not in url:
        return False
    try:
        return bool(page.evaluate(
            r"""() => !![...document.querySelectorAll('a[href^="#"], .nav-link, .nav-tabs a, .nav a')]
                .find(a => /เอกสารตอบรับ|ข้อมูลคนต่างด้าว/.test(a.innerText || ''))"""
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
    base_headers = ["ลำดับ", "ชื่อคนต่างด้าว(Eng)", "เลขที่คำขอ", "Username", "สถานะคำขอ", "Status รวม"]
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
            r.get("seq", ""), r.get("name_eng", ""), r.get("req_no", ""),
            r.get("username", ""), r.get("status_text", ""), r.get("status", ""),
        ]
        docs = r.get("docs", {})
        for dk in doc_keys:
            d = docs.get(dk, {})
            row_vals.append(d.get("status", ""))
            row_vals.append(d.get("file", ""))
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
                cell.value = f'=HYPERLINK("result_docs/{fname}","{fname}")'
                cell.font = link_font

    base_widths = [8, 30, 20, 26, 22, 12]
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
        "name_eng": "", "status_text": "",
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
        # ดึงชื่อคนต่างด้าวไม่ได้ → ไม่ดาวน์โหลด (กันได้ไฟล์ชื่อผิด/ชื่อหาย)
        # บันทึกเป็น NO_NAME ให้ตัวรันกลับมาทำ 'รอบสอง' อีกครั้ง
        res["status"] = "NO_NAME"
        res["error"] = "ดึงชื่อคนต่างด้าวไม่สำเร็จ — ข้ามไว้ทำรอบสอง"
        log(f"     ⏭ ข้าม {req_no}: ดึงชื่อคนต่างด้าวไม่ได้ (จะรวมไว้ทำรอบสอง)")
        return res
    res["name_eng"] = name_eng
    name_safe = (_receipt_safe_name(name_eng) or "").strip(" -_.").strip() or _receipt_safe_name(req_no)

    def _download_all() -> None:
        for dk in doc_keys:
            r = _download_response_doc_named(page, docs_dir, name_safe, RESULT_DOC_TYPES[dk], log=log, req_no=req_no)
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
                        results.append({**rec, "name_eng": "", "status_text": "",
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
                    login(page, login_cfg)
                except Exception as e:
                    log(f"      ✗ login ไม่สำเร็จ: {e} — ข้ามกลุ่มนี้")
                    for rec in recs:
                        results.append({**rec, "name_eng": "", "status_text": "",
                                        "docs": {dk: {"status": "", "file": "", "error": ""} for dk in doc_keys},
                                        "status": "LOGIN_FAIL", "error": str(e)[:200]})
                        done_count += 1
                    _save_result_docs_report(results, out_path, doc_keys, log=lambda *_: None)
                    continue

                RELOGIN_EVERY = 80  # re-login เชิงรุกทุก N รายการ กัน session timeout
                MAX_RETRY_PASSES = 3  # จำนวนรอบ retry คำขอที่ยังไม่สำเร็จ (หลังรอบแรก)

                # เก็บรายการคำขอของบัญชีนี้ครั้งเดียว (filter สถานะ AP/APSS) เพื่อ map
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
    # หมายเหตุ: โหมดนี้ล็อกที่ 'รอนัดหมาย' เสมอ — ไม่ใช้ filter_status_ids ของโหมด e-Tracking
    status_ids = list(cfg.get("result_status_ids") or ["AP"])

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
                        # ดึงชื่อไม่ได้ → ไม่ดาวน์โหลด ปล่อยให้รอบสองมาทำใหม่
                        rec["status"] = "NO_NAME"
                        rec["error"] = "ดึงชื่อคนต่างด้าวไม่สำเร็จ — ข้ามไว้ทำรอบสอง"
                        log(f"     ⏭ ข้าม {req_no}: ดึงชื่อคนต่างด้าวไม่ได้ (จะรวมไว้ทำรอบสอง)")
                        return rec
                    rec["name_eng"] = name_eng
                    name_safe = (_receipt_safe_name(name_eng) or _receipt_safe_name(req_no)).strip(" -_.").strip()

                    for dk in doc_keys:
                        r = _download_response_doc_named(page, docs_dir, name_safe, RESULT_DOC_TYPES[dk], log=log, req_no=req_no)
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
                                    for dk in doc_keys:
                                        r = _download_response_doc_named(page, docs_dir, name_safe, RESULT_DOC_TYPES[dk], log=log, req_no=req_no)
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


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0, help="0 = ทั้งหมด, >0 = จำกัด N รายการ")
    ap.add_argument("--out", type=Path, default=REPORTS_DIR / "WA_report.xlsx")
    ap.add_argument(
        "--request-type", default=None,
        help="รหัสรายการคำขอ เช่น MT_59_MOU_RENEWAL (default ใช้จาก .env)",
    )
    ap.add_argument(
        "--mode", choices=["etracking", "aliens", "register", "receipts", "results", "inform"], default="etracking",
        help="etracking=ดึงจาก e-Tracking, aliens=ดึงรายชื่อคนต่างด้าว, register=ลงทะเบียน, receipts=ดาวน์โหลดใบเสร็จ, results=ดาวน์โหลดใบแจ้งผล/ใบรับคำขอ, inform=แจ้งการจ้างคนต่างด้าวเข้าทำงาน (บต.52)",
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
        "--doc-types", default=None,
        help="(โหมด receipts) ประเภทเอกสารที่จะดาวน์โหลด คั่นด้วย , เช่น 'receipt,bt44,bt22,bt53,bt56' หรือ 'all' (default=ทั้งหมด)",
    )
    ap.add_argument(
        "--result-doc-types", default=None,
        help="(โหมด results) ประเภทเอกสาร คั่นด้วย , เช่น 'result_notice,request_receipt' หรือ 'all' (default=ทั้งหมด)",
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
        "--commit", action="store_true",
        help="(โหมด inform) ส่งคำขอจริง (กดข้อ 7+8) — ค่า default จะหยุดก่อนยืนยัน",
    )
    ap.add_argument(
        "--list-request-types", action="store_true",
        help="พิมพ์ list option 'รายการคำขอ' จาก dropdown ในหน้า e-Tracking แล้วจบโปรแกรม",
    )
    args = ap.parse_args()

    cfg = load_config()
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
