"""GUI สำหรับดึงรายงาน / ดาวน์โหลดเอกสาร / ยื่นคำขอ จากระบบ e-WorkPermit
รัน: python gui_app.py

โครงสร้างหน้าจอ (V31):
  ┌ Sidebar          │ หัวข้อโหมด + ป้ายความปลอดภัย + คำอธิบาย                     ┐
  │ (เลือกโหมด       │ การ์ดตัวเลือก (เลื่อนได้) — แสดงเฉพาะที่โหมดนั้นใช้จริง        │
  │  จัดกลุ่ม 4 หมวด) │ แถบปุ่ม เริ่ม/พัก/ต่อ/ยกเลิก + progress (อยู่ล่างสุดเสมอ)      │
  │                  │ บันทึกการทำงาน (พับ/กางได้)                                 │
  └──────────────────┴────────────────────────────────────────────────────────┘
"""
from __future__ import annotations

import json
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import usage_tracker

from scrape_wa import (
    run_scrape,
    run_scrape_multi,
    run_scrape_by_ref,
    run_aliens_scrape,
    run_register,
    run_receipts,
    run_result_docs,
    run_result_docs_by_ref,
    run_inform_employer,
    run_bt30,
    run_bt44,
    run_bill_payment,
    run_payment_receipts,
    run_appointment,
    run_bt30_ctn,
    run_namelist_alien,
    run_booking_availability,
    run_permit_report,
    run_permit_report_multi,
    run_ticket_report,
    booking_scan_once,
    _read_bt30_excel,
    _bt30_preflight_doc_sizes,
    _parse_row_range,
    _BT30_MAX_DOC_MB,
    REQUEST_PROFILES,
    DEFAULT_STATUS_WHITELIST,
    get_profile,
    ALIEN_SUB_TABS,
)

ROOT = Path(__file__).parent
# โฟลเดอร์เก็บไฟล์รายงาน (Report) ทุกโหมด ให้เป็นระเบียบ
REPORTS_DIR = ROOT / "reports"
APP_VERSION = "31"
# ไฟล์จำค่าที่กรอกล่าสุด (ไม่เก็บรหัสผ่าน)
SETTINGS_FILE = ROOT / "_gui_settings.json"

# (code, label) ของรายการคำขอ — ดึงจาก dropdown #Filter_request_list ของเว็บจริง
# เรียงตามลำดับเดียวกับ select บนเว็บ
REQUEST_TYPES: list[tuple[str, str]] = [
    ("", "ทั้งหมด (ไม่กรอง)"),
    ("MT_63_2_19_RENEWAL", "การต่ออายุใบอนุญาตทำงานให้กับคนต่างด้าวสัญชาติลาว เมียนมา และเวียดนาม ตามมติคณะรัฐมนตรีเมื่อวันที่ 14 กรกฎาคม 2569"),
    ("MT_63_RENEWAL", "การยื่นคำขอต่ออายุใบอนุญาตทำงานของคนต่างด้าว"),
    ("MT_63_1_RENEWAL", "การยื่นคำขอต่ออายุใบอนุญาตทำงานของคนต่างด้าว (มาตรา 63 วรรค 1)"),
    ("MT_59_RENEWAL", "การยื่นคำขอต่ออายุใบอนุญาตทำงานของคนต่างด้าวซึ่งได้รับอนุญาตให้อยู่ในราชอาณาจักรเป็นการชั่วคราว (Non-Immigrant) เพื่อทำงานที่ใช้ความรู้ความสามารถทักษะฝีมือ"),
    ("MT_59_MOU_RENEWAL", "การยื่นคำขอต่ออายุใบอนุญาตทำงานของคนต่างด้าวที่ได้รับอนุญาตทำงานตาม MoU"),
    ("MT_63_2_AGN_RENEWAL", "การยื่นคำขอต่ออายุใบอนุญาตทำงานของคนต่างด้าวสัญชาติเมียนมาตามมติ ครม. เมื่อวันที่ 8 กรกฎาคม พ.ศ. 2568"),
    ("MT_63_2_3103_RENEWAL", "การยื่นคำขอต่ออายุใบอนุญาตทำงานให้กับคนต่างด้าวสัญชาติลาว เมียนมา และเวียดนาม ตามมติคณะรัฐมนตรีเมื่อวันที่ 2 ธ.ค. 2568"),
    ("MT_63_2_1302_RENEWAL", "การยื่นคำขอต่ออายุใบอนุญาตทำงานให้กับคนต่างด้าวสัญชาติลาว และเวียดนาม ตามมติคณะรัฐมนตรีเมื่อวันที่ 2 ธ.ค.2568"),
    ("MT_63_V_2", "การยื่นคำขอรับใบอนุญาตทำงานขอคนต่างด้าวที่อยู่ในพื้นที่พักพิงชั่วคราวสำหรับผู้หนีภัยการสู้รบจากเมียนมา"),
    ("MT_63_1", "การยื่นคำขอรับใบอนุญาตทำงานของคนต่างด้าว (มาตรา 63 วรรค 1)"),
    ("MT_63", "การยื่นคำขอรับใบอนุญาตทำงานของคนต่างด้าว"),
    ("MT_64", "การยื่นคำขอรับใบอนุญาตทำงานของคนต่างด้าว ที่ถือบัตรผ่านแดน (Border Pass)"),
    ("MT_59", "การยื่นคำขอรับใบอนุญาตทำงานของคนต่างด้าวซึ่งได้รับอนุญาตให้อยู่ในราชอาณาจักรเป็นการชั่วคราว"),
    ("MT_60_1", "การยื่นคำขอรับใบอนุญาตทำงานของคนต่างด้าวที่อยู่ต่างประเทศที่จะทำงานกับนายจ้างในประเทศไทย"),
    ("MT_62", "การยื่นคำขอรับใบอนุญาตทำงานของคนต่างด้าวที่เข้ามาทำงานตามกฎหมายว่าด้วยการนิคมอุตสาหกรรมแห่งประเทศไทย กรมเชื้อเพลิง ธรรมชาติ และกฎหมายว่าด้วยอนุญาโตตุลาการ"),
    ("MT_62_BOI", "การยื่นคำขอรับใบอนุญาตทำงานของคนต่างด้าวที่เข้ามาทำงานตามกฎหมายว่าด้วยการส่งเสริมการลงทุน (BOI)"),
    ("REPLACE_CARD_25", "การยื่นคำขอรับใบแทนใบอนุญาตทำงานกรณีใบอนุญาตทำงานสูญหายถูกทำลายหรือชำรุดในสาระสำคัญ"),
    ("MT_63_2_19", "การยื่นคำขอใบอนุญาตทำงานของคนต่างด้าวที่มีสถานะไม่ถูกต้องตามกฎหมายสัญชาติลาว เมียนมา และเวียดนาม ตามมติ ครม. เมื่อวันที่ 11 พฤศจิกายน พ.ศ. 2568"),
    ("MT_41_4_59", "การยื่นคำร้องขอนำคนต่างด้าวสัญชาติกัมพูชา ลาว เวียดนาม เข้ามาทำงานกับนายจ้างในประเทศ (MoU) ของผู้รับใบอนุญาตนำคนต่างด้าวมาทำงานในประเทศ (บริษัทนำเข้า)"),
    ("MT_13_1_INFORM", "การแจ้งการจ้างคนต่างด้าวเข้าทํางานกับนายจ้าง"),
    ("MT_61", "การแจ้งการทำงานที่มีลักษณะจำเป็นหรือเร่งด่วนหรือเป็นงานเฉพาะกิจที่มีระยะเวลาทำงานให้เสร็จสิ้นภายใน 15 วัน"),
    ("MT_43", "การแจ้งการส่งมอบคนต่างด้าวที่นำเข้ามาทำงานตาม MoU ให้กับนายจ้าง (บริษัทนำเข้าฯ ส่งมอบคนงานให้นายจ้าง)"),
    ("MT_62_RENEWAL", "การแจ้งขยายระยะเวลาทำงานของคนต่างด้าวที่เข้ามาทำงานตามกฎหมายว่าด้วยการนิคมอุตสาหกรรมแห่งประเทศไทย กรมเชื้อเพลิง ธรรมชาติ และกฎหมายว่าด้วยอนุญาโตตุลาการ"),
    ("MT_62_BOI_RENEWAL", "การแจ้งขยายระยะเวลาทำงานของคนต่างด้าวที่เข้ามาทำงานตามกฎหมายว่าด้วยการส่งเสริมการลงทุน (BOI)"),
    ("MT_61_EXTEND", "การแจ้งขยายระยะเวลาเพื่อทำงานอันมีลักษณะจำเป็นหรือเร่งด่วนหรืองานเฉพาะกิจ"),
    ("MT_13_EXIT", "การแจ้งคนต่างด้าวออกจากงานของนายจ้าง"),
    ("REPLACE_CARD_MOU", "ขอออกใบอนุญาตทำงานใหม่เนื่องจากใบอนุญาตทำงานสูญหาย ถูกทำลาย หรือชำรุดในสาระสำคัญ"),
    ("CHANGE_44", "คำขอเปลี่ยนรายการในใบอนุญาตทำงาน"),
    ("CHANGE_44_45", "คำขอเปลี่ยนรายการในใบอนุญาตทำงาน/คำขอเปลี่ยนรายการใบอนุญาตทำงาน กรณีเปลี่ยนหรือเพิ่มประเภทงาน (บต.44 และ บต.45)"),
    ("CHANGE_44_22", "คำขอเปลี่ยนรายการในใบอนุญาตทำงาน/คำขอเปลี่ยนรายการใบอนุญาตทำงาน กรณีเปลี่ยนหรือเพิ่มประเภทงาน (บต.44 และ บต.22)"),
    ("CHANGE_45", "คำขอเปลี่ยนรายการในใบอนุญาตทำงานกรณีเปลี่ยนหรือเพิ่มประเภทงานตามบัญชีสองและบัญชีสามท้ายประกาศกระทรวงแรงงาน เรื่อง กำหนดงานที่ห้ามคำต่างด้าวทำลงวันที่ 1 เมษายน 2563"),
    ("CHANGE_22", "คำขอเปลี่ยนรายการใบอนุญาตทำงาน กรณีเปลี่ยนหรือเพิ่มประเภทงาน"),
    ("MT_60_2", "นายจ้างยื่นคำขอรับใบอนุญาตทำงานแทนคนต่างด้าวที่อยู่ต่างประเทศเพื่อทำงานที่ใช้ความรู้ความสามารถทักษะฝีมือ"),
    ("MT_46_59", "นายจ้างยื่นคำร้องขอนำคนต่างด้าวสัญชาติกัมพูชา ลาว เมียนมา และเวียดนาม เข้ามาทำงานกับนายจ้างในประเทศ (MOU)"),
    ("MT_63_2_AGN_CHANGE", "เปลี่ยนรายการ Name List มติครม. 8 ก.ค. 68"),
    ("MT_REVOKE", "เพิกถอนใบอนุญาตทำงาน"),
    ("MT_50_1", "แจ้งไม่รับคนต่างด้าวเข้าทํางาน/คนต่างด้าวไม่ยินยอมทํางาน (คนต่างด้าวที่บริษัทนําเข้าฯ ส่งมอบ)"),
]

# Status checkbox (id → label) ตาม UI ของเว็บ
STATUS_FILTERS = [
    ("WP", "รอชำระเงิน"),
    ("WCOSNA", "รอตรวจสอบเอกสารของบริษัท OS"),
    ("WA", "รอยื่นเอกสารเพิ่มเติม"),
    ("AP", "รอนัดหมาย"),
    ("SS", "ดำเนินการเสร็จสิ้น"),
]

# Whitelist สถานะที่จะ scrape detail (substring match)
WHITELIST_OPTIONS = [
    "รออนุมัติคำขอของนายทะเบียน (สำนักงานจัดหางาน)",
    "รอตรวจสอบเอกสารของบริษัท OS",
    "รอยื่นเอกสารเพิ่มเติม",
    "รอพิจารณาคำขอของผู้ช่วยนายทะเบียน (สำนักงานจัดหางาน)",
    "รอชำระเงิน",
    "รอนัดหมาย",
    "นัดหมายแล้ว",
    "ดำเนินการเสร็จสิ้น",
]

# ---------------------------------------------------------------------------
# ทะเบียนโหมด — แหล่งข้อมูลเดียวสำหรับ sidebar / หัวข้อ / login ที่ต้องใช้ / ไฟล์ผลลัพธ์
#   login : "creds"  = ใช้ Username/Password ในหน้าจอ
#           "file"   = ใช้บัญชีจาก Excel (UsernameLogin ฯลฯ) ไม่ต้องกรอกด้านบน
#           "either" = ใช้ได้ทั้งสองแบบ
#           "none"   = ไม่ต้อง login
#   safety: "read" (อ่านอย่างเดียว) / "write" (บันทึกข้อมูลลงระบบ) / "danger" (ส่งคำขอจริง/มีค่าธรรมเนียม)
# ---------------------------------------------------------------------------
MODE_GROUPS: list[tuple[str, str]] = [
    ("report", "📊  รายงาน / ดึงข้อมูล"),
    ("download", "📥  ดาวน์โหลดเอกสาร"),
    ("submit", "✍️  ยื่นคำขอ / บันทึกข้อมูล"),
    ("booking", "📅  คิวถ่ายบัตร"),
]

MODES: list[dict] = [
    dict(key="etracking", group="report", nav="e-Tracking (รายการคำขอ)",
         title="e-Tracking — ดึงรายการคำขอ",
         desc="ดึงรายการคำขอจากหน้า e-Tracking ตามรายการคำขอ/สถานะ/วันที่ แล้วบันทึกเป็น Excel "
              "(รายงาน WA + ข้อมูลคนต่างด้าว)",
         login="creds", safety="read", out="WA_report.xlsx"),
    dict(key="permit_report", group="report", nav="รายงานข้อมูลการขออนุญาต",
         title="รายงานข้อมูลการขออนุญาต",
         desc="สถานะคำขอ (อนุมัติคำขอ) + ข้อมูลสถานประกอบการ (บริษัท/จังหวัด) ตามรายการคำขอ/สถานะ/วันที่ที่เลือก",
         login="creds", safety="read", out="WA_permit_report.xlsx"),
    dict(key="aliens", group="report", nav="ข้อมูลคนต่างด้าว (จัดการบัญชี)",
         title="จัดการบัญชี → ข้อมูลคนต่างด้าว",
         desc="ดึงรายชื่อคนต่างด้าวทั้งหมดจากหน้าจัดการบัญชี (อ่านเฉพาะตาราง ไม่เปิดรายละเอียดรายคน — เร็วมาก)",
         login="creds", safety="read", out="WA_aliens.xlsx"),
    dict(key="namelist_alien", group="report", nav="รายชื่อยื่นต่ออายุแล้ว (NameList)",
         title="รายชื่อคนต่างด้าวที่ยื่นคำขอต่ออายุแล้ว (NameListAlien)",
         desc="เปิดหน้า NameListAlien ตาม Form Type แล้ววนทุกหน้า (100 แถว/หน้า) บันทึกลง Excel",
         login="creds", safety="read", out="WA_namelist_alien_report.xlsx"),
    dict(key="ticket", group="report", nav="ตรวจสอบ Ticket ID",
         title="ตรวจสอบ Ticket ID (แจ้งปัญหาการใช้งาน)",
         desc="หน้า /TicketRequest เปิดสาธารณะ — ไม่ต้อง login ค้นหาทีละเลข Ticket ID จาก Excel แล้วเก็บผลทุกคอลัมน์",
         login="none", safety="read", out="WA_ticket_report.xlsx"),

    dict(key="receipts", group="download", nav="ใบเสร็จ / บต.44 / บต.22 ฯลฯ",
         title="ดาวน์โหลดใบเสร็จและแบบฟอร์ม (บต.44 / บต.22 / บต.53 / บต.55 / บต.52 / บต.56)",
         desc="อ่านเลขคำขอจาก RequestData.xlsx → login ตามบัญชีเจ้าของคำขอ → ดาวน์โหลด PDF ลง reports/receipts",
         login="file", safety="read", out="WA_receipts_report.xlsx"),
    dict(key="results", group="download", nav="ใบแจ้งผล / ใบรับคำขอ",
         title="ดาวน์โหลดใบแจ้งผล / ใบรับคำขอ / บต.50",
         desc="โหมด A: ตามเลขคำขอใน Ref_number.xlsx  |  โหมด B: ดึงจาก e-Tracking ตามรายการคำขอ (สถานะ 'รอนัดหมาย')",
         login="file", safety="read", out="WA_result_docs_report.xlsx"),
    dict(key="bill_payment", group="download", nav="ใบแจ้งชำระเงิน",
         title="ดาวน์โหลดใบแจ้งชำระเงิน (รอจ่ายค่าธรรมเนียม)",
         desc="วนทุกบัญชีใน UsernameLogin.xlsx → กรองสถานะ 'รอชำระเงิน' → พิมพ์แบบฟอร์มการชำระเงินเก็บไว้ (ไม่มีการตัดเงิน)",
         login="file", safety="read", out="WA_bill_payment_report.xlsx"),
    dict(key="payment_receipts", group="download", nav="ใบเสร็จรับเงินทั้งชุด",
         title="ดาวน์โหลดใบเสร็จรับเงินทั้งชุด (900 + 100 + อื่นๆ)",
         desc="อ่านเลขคำขอจาก Request_Receipt.xlsx → แท็บ 'การชำระเงิน' → ดูใบเสร็จรับเงิน → แยก PDF เป็น 1 ใบ/1 คน/1 ไฟล์",
         login="file", safety="read", out="WA_payment_receipts_report.xlsx"),
    dict(key="bt30_ctn", group="download", nav="ใบตอบรับ / ใบนัดหมาย",
         title="ดาวน์โหลดใบตอบรับ (บต.30) / บต.25 / ใบนัดหมาย / ใบเสร็จ",
         desc="วนทุกบัญชี (หรือบัญชีเดียวด้านล่าง) → e-Tracking ตามรายการคำขอ/สถานะ/วันที่ → เปิดแต่ละคำขอแล้วดาวน์โหลด PDF ลง reports/bt30_ctn",
         login="either", safety="read", out="WA_bt30_ctn_report.xlsx"),
    dict(key="appointment", group="download", nav="นัดหมายถ่ายบัตร (ที่อยู่จากใบเสร็จ)",
         title="นัดหมายถ่ายบัตร — เก็บที่อยู่จากใบเสร็จค่าธรรมเนียม",
         desc="วนทุกบัญชีใน UsernameLogin.xlsx → สถานะ 'รอนัดหมาย' (AP) → ดาวน์โหลดใบเสร็จ → อ่านจังหวัด/อำเภอจาก PDF",
         login="file", safety="read", out="WA_appointment_report.xlsx"),

    dict(key="register", group="submit", nav="ลงทะเบียนคนต่างด้าว",
         title="ลงทะเบียนคนต่างด้าวจาก Excel (Register + ดึงข้อมูลใบอนุญาต)",
         desc="สร้างบัญชีคนต่างด้าวทีละแถวจาก Excel ด้วยบัญชีของคนต่างด้าวแต่ละคน",
         login="file", safety="write", out="WA_register_report.xlsx"),
    dict(key="inform", group="submit", nav="แจ้งเข้านายจ้าง (บต.52)",
         title="แจ้งเข้านายจ้าง (แบบ บต.52) — INFORM_ENTER_EXIT",
         desc="กรอกฟอร์มจาก FormRequestEmployment.xlsx ด้วยบัญชีแรกใน UsernameLogin.xlsx — ค่าเริ่มต้นเป็นโหมดทดสอบ (ไม่ส่งจริง)",
         login="file", safety="write", out="WA_inform_report.xlsx"),
    dict(key="bt30", group="submit", nav="ต่อใบอนุญาต MoU (บต.30)",
         title="ยื่นต่อใบอนุญาตทำงาน (แบบ บต.30 / MoU)",
         desc="กรอกค้นหาข้อมูลคนต่างด้าวจาก from_bt30.xlsx → ทำขั้นตอน 2–6 → หยุดก่อนชำระเงิน (เว้นแต่เปิด 'ส่งคำขอจริง')",
         login="file", safety="danger", out="WA_bt30_report.xlsx"),
    dict(key="bt44", group="submit", nav="เปลี่ยนย้ายนายจ้าง (บต.44)",
         title="เปลี่ยนย้ายนายจ้างในระบบ (บต.44)",
         desc="กรอกข้อมูลจาก from_bt44 → บันทึก (ไม่กดถัดไป/ไม่ส่งคำขอ) — มีโหมดทดลองยื่น (Dry-Run)",
         login="file", safety="write", out="WA_bt44_report.xlsx"),

    dict(key="booking_availability", group="booking", nav="ตรวจวันว่างจอง (Real-Time)",
         title="ตรวจวันว่างจอง (คิวถ่ายบัตร) — Real-Time",
         desc="หา token จากคำขอสถานะ AP/SS → ดึงรายชื่อสาขา → วนปฏิทิน × ช่วงเวลา → Excel 3 sheet  |  หรือเปิดหน้าจอ Monitor สด ๆ",
         login="either", safety="read", out="WA_booking_availability.xlsx"),
]
MODE_BY_KEY: dict[str, dict] = {m["key"]: m for m in MODES}
DEFAULT_OUT_NAMES: set[str] = {m["out"] for m in MODES}

SAFETY_BADGE = {
    "read": ("อ่านอย่างเดียว — ไม่ส่งคำขอ/ไม่ตัดเงิน", "Safe.Badge.TLabel"),
    "write": ("บันทึกข้อมูลลงระบบจริง", "Warn.Badge.TLabel"),
    "danger": ("อาจส่งคำขอจริง + มีค่าธรรมเนียม", "Danger.Badge.TLabel"),
}

# ---------------------------------------------------------------------------
# Theme
# ---------------------------------------------------------------------------
PALETTE = {
    "bg": "#F3F4F6",        # พื้นหลังหน้าต่าง
    "surface": "#FFFFFF",   # การ์ด/ช่องกรอก
    "sidebar": "#1F2937",   # sidebar เข้ม
    "sidebar_fg": "#E5E7EB",
    "sidebar_sel": "#2563EB",
    "border": "#E5E7EB",
    "text": "#111827",
    "muted": "#6B7280",
    "accent": "#2563EB",
    "accent_hover": "#1D4ED8",
    "accent_fg": "#FFFFFF",
    "ok": "#15803D",
    "ok_bg": "#DCFCE7",
    "warn": "#B45309",
    "warn_bg": "#FEF3C7",
    "danger": "#B91C1C",
    "danger_bg": "#FEE2E2",
    "info": "#1D4ED8",
}
FONT_UI = ("Segoe UI", 10)
FONT_UI_BOLD = ("Segoe UI", 10, "bold")
FONT_SMALL = ("Segoe UI", 9)
FONT_TITLE = ("Segoe UI", 14, "bold")
FONT_MONO = ("Consolas", 10)


def apply_theme(root: tk.Tk) -> None:
    """ตั้งค่า ttk style ให้ดูทันสมัย/สม่ำเสมอทั้งแอป (ใช้ธีม clam เป็นฐานเพื่อให้ปรับสีได้)"""
    st = ttk.Style(root)
    try:
        st.theme_use("clam")
    except Exception:
        pass
    P = PALETTE
    root.configure(bg=P["bg"])
    root.option_add("*Font", FONT_UI)
    root.option_add("*TCombobox*Listbox.font", FONT_UI)

    st.configure(".", background=P["bg"], foreground=P["text"], font=FONT_UI,
                 bordercolor=P["border"], lightcolor=P["border"], darkcolor=P["border"],
                 troughcolor=P["border"], focuscolor=P["accent"])
    st.configure("TFrame", background=P["bg"])
    st.configure("Surface.TFrame", background=P["surface"])
    st.configure("Header.TFrame", background=P["surface"])
    st.configure("Bottom.TFrame", background=P["surface"])
    st.configure("TLabel", background=P["bg"], foreground=P["text"])
    st.configure("Surface.TLabel", background=P["surface"])
    st.configure("Title.TLabel", background=P["surface"], font=FONT_TITLE)
    st.configure("Desc.TLabel", background=P["surface"], foreground=P["muted"], font=FONT_UI)
    st.configure("Muted.TLabel", foreground=P["muted"], font=FONT_SMALL)
    st.configure("Hint.TLabel", foreground=P["muted"], font=FONT_SMALL)
    st.configure("Info.TLabel", foreground=P["info"], font=FONT_SMALL)
    st.configure("Ok.TLabel", foreground=P["ok"], font=FONT_SMALL)
    st.configure("Warn.TLabel", foreground=P["warn"], font=FONT_SMALL)
    st.configure("Danger.TLabel", foreground=P["danger"], font=("Segoe UI", 9, "bold"))
    st.configure("Section.TLabel", foreground=P["text"], font=FONT_UI_BOLD)

    # Badge
    for name, fg, bg in (
        ("Safe", P["ok"], P["ok_bg"]),
        ("Warn", P["warn"], P["warn_bg"]),
        ("Danger", P["danger"], P["danger_bg"]),
    ):
        st.configure(f"{name}.Badge.TLabel", foreground=fg, background=bg,
                     font=("Segoe UI", 9, "bold"), padding=(8, 2))

    # Card (LabelFrame)
    st.configure("TLabelframe", background=P["bg"], bordercolor=P["border"],
                 relief="solid", borderwidth=1, padding=10)
    st.configure("TLabelframe.Label", background=P["bg"], foreground=P["text"], font=FONT_UI_BOLD)

    # Inputs
    st.configure("TEntry", fieldbackground=P["surface"], padding=4)
    st.configure("TCombobox", fieldbackground=P["surface"], padding=3, arrowsize=14)
    st.map("TCombobox", fieldbackground=[("readonly", P["surface"])])
    st.configure("TSpinbox", fieldbackground=P["surface"], padding=3, arrowsize=12)
    st.configure("TCheckbutton", background=P["bg"])
    st.configure("TRadiobutton", background=P["bg"])
    st.configure("Surface.TCheckbutton", background=P["surface"])

    # Buttons
    st.configure("TButton", padding=(10, 5), background=P["surface"], relief="flat",
                 bordercolor=P["border"])
    st.map("TButton",
           background=[("active", "#E5E7EB"), ("disabled", "#F3F4F6")],
           foreground=[("disabled", "#9CA3AF")])
    st.configure("Accent.TButton", background=P["accent"], foreground=P["accent_fg"],
                 font=FONT_UI_BOLD, padding=(16, 7), bordercolor=P["accent"])
    st.map("Accent.TButton",
           background=[("active", P["accent_hover"]), ("disabled", "#93C5FD")],
           foreground=[("disabled", "#EFF6FF")])
    st.configure("Danger.TButton", foreground=P["danger"])
    st.configure("Link.TButton", padding=(4, 2), background=P["bg"], relief="flat",
                 foreground=P["info"])
    st.map("Link.TButton", background=[("active", P["bg"])])

    # Progress
    st.configure("TProgressbar", troughcolor=P["border"], background=P["accent"],
                 bordercolor=P["border"], lightcolor=P["accent"], darkcolor=P["accent"])

    # Sidebar nav (Treeview)
    st.configure("Nav.Treeview", background=P["sidebar"], fieldbackground=P["sidebar"],
                 foreground=P["sidebar_fg"], rowheight=30, borderwidth=0, relief="flat",
                 font=FONT_UI)
    st.map("Nav.Treeview",
           background=[("selected", P["sidebar_sel"])],
           foreground=[("selected", "#FFFFFF")])
    st.layout("Nav.Treeview", [("Nav.Treeview.treearea", {"sticky": "nswe"})])

    # Notebook / Scrollbar
    st.configure("TScrollbar", troughcolor=P["bg"], background="#D1D5DB", arrowsize=12,
                 bordercolor=P["bg"])
    st.configure("Treeview", rowheight=24, fieldbackground=P["surface"], background=P["surface"])
    st.configure("Treeview.Heading", font=FONT_UI_BOLD, padding=4)


class Collapsible(ttk.Frame):
    """กล่องพับ/กางได้ — header เป็นปุ่ม ▸/▾ + ชื่อ, เนื้อหาอยู่ใน self.body"""

    def __init__(self, master, title: str, open: bool = False, hint: str = ""):
        super().__init__(master)
        self._open = tk.BooleanVar(value=open)
        self._title = title
        hdr = ttk.Frame(self)
        hdr.pack(fill="x")
        self._btn = ttk.Button(hdr, text="", style="Link.TButton", command=self.toggle)
        self._btn.pack(side="left")
        if hint:
            ttk.Label(hdr, text=hint, style="Hint.TLabel").pack(side="left", padx=(6, 0))
        self.body = ttk.Frame(self, padding=(18, 4, 0, 4))
        self._refresh()

    def toggle(self) -> None:
        self._open.set(not self._open.get())
        self._refresh()

    def set_open(self, v: bool) -> None:
        self._open.set(bool(v))
        self._refresh()

    def is_open(self) -> bool:
        return bool(self._open.get())

    def _refresh(self) -> None:
        if self._open.get():
            self._btn.config(text=f"▾  {self._title}")
            self.body.pack(fill="x")
        else:
            self._btn.config(text=f"▸  {self._title}")
            self.body.pack_forget()


class Tooltip:
    """tooltip เล็ก ๆ สำหรับ widget (hover 500ms)"""

    def __init__(self, widget, text: str):
        self.widget = widget
        self.text = text
        self._tip: tk.Toplevel | None = None
        self._job = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _e=None):
        self._cancel()
        self._job = self.widget.after(500, self._show)

    def _cancel(self):
        if self._job:
            try:
                self.widget.after_cancel(self._job)
            except Exception:
                pass
            self._job = None

    def _show(self):
        if self._tip or not self.text:
            return
        try:
            x = self.widget.winfo_rootx() + 12
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
            self._tip = tk.Toplevel(self.widget)
            self._tip.wm_overrideredirect(True)
            self._tip.wm_geometry(f"+{x}+{y}")
            tk.Label(
                self._tip, text=self.text, justify="left", wraplength=420,
                background="#111827", foreground="#F9FAFB", font=FONT_SMALL,
                padx=8, pady=5,
            ).pack()
        except Exception:
            self._tip = None

    def _hide(self, _e=None):
        self._cancel()
        if self._tip:
            try:
                self._tip.destroy()
            except Exception:
                pass
            self._tip = None


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("e-WorkPermit Tools")
        self.geometry("1180x800")
        self.minsize(960, 640)
        try:
            self.tk.call("tk", "scaling", 1.2)
        except Exception:
            pass
        apply_theme(self)

        self._log_q: queue.Queue = queue.Queue()      # str (log) หรือ tuple ("progress", done, total)
        self._worker: threading.Thread | None = None
        self._cancel = threading.Event()
        # _pause_evt: set = ทำงานปกติ, clear = พักค้างไว้ (worker จะรอที่จุดปลอดภัย)
        self._pause_evt = threading.Event()
        self._pause_evt.set()
        # ตัวแปรที่จะจำค่าไว้ข้ามการเปิดโปรแกรม (ไม่รวมรหัสผ่าน)
        self._persist: dict[str, tk.Variable] = {}

        # โหมด receipts: เลือกประเภทเอกสารที่จะดาวน์โหลด
        self.doc_receipt = tk.BooleanVar(value=True)
        self.doc_bt44 = tk.BooleanVar(value=True)
        self.doc_bt22 = tk.BooleanVar(value=True)
        self.doc_bt53 = tk.BooleanVar(value=False)
        self.doc_bt55 = tk.BooleanVar(value=False)
        self.doc_bt52 = tk.BooleanVar(value=False)
        self.doc_bt56 = tk.BooleanVar(value=False)
        # โหมด results: เลือกเอกสารผลอนุญาต
        self.doc_result_notice = tk.BooleanVar(value=True)
        self.doc_request_receipt = tk.BooleanVar(value=True)
        self.doc_result_bt50 = tk.BooleanVar(value=False)

        self._last_progress: tuple[int, int] = (0, 0)
        self._job_meta: dict | None = None
        self._build_ui()
        self._load_settings()
        self._on_mode_changed()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<F5>", lambda _e: self._on_start())
        self.after(120, self._drain_log)
        # ส่งบันทึกการใช้งานที่ค้างอยู่ (ถ้ามี) ตอนเปิดโปรแกรม
        self.after(1500, lambda: usage_tracker.flush_async(log=self._log))

    # ---------- usage tracking ----------
    def _usage_status(self) -> None:
        """ปุ่ม 'สถานะการบันทึกใช้งาน' — แสดงสถานะ + ทดสอบเชื่อมต่อ MySQL"""
        pending = usage_tracker.pending_count()
        ok, msg = usage_tracker.test_connection()
        operator = self.operator.get().strip() or usage_tracker.default_operator()
        lines = [
            ("✓ " if ok else "✗ ") + msg,
            "",
            f"ผู้ใช้งาน: {operator}",
            f"รหัสเครื่อง: {usage_tracker.install_id()}",
            f"รายการค้างส่ง: {pending}",
        ]
        if ok and pending:
            usage_tracker.flush_async(log=self._log)
            lines.append("(กำลังส่งรายการค้างอยู่ — ดูผลใน log)")
        txt = "\n".join(lines)
        (messagebox.showinfo if ok else messagebox.showwarning)("การบันทึกการใช้งาน", txt)

    # ---------- settings (จำค่าล่าสุด) ----------
    def _remember(self, name: str, var: tk.Variable) -> tk.Variable:
        self._persist[name] = var
        return var

    def _load_settings(self) -> None:
        try:
            if not SETTINGS_FILE.exists():
                return
            data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                return
            for name, var in self._persist.items():
                if name in data:
                    try:
                        var.set(data[name])
                    except Exception:
                        pass
            geo = data.get("_geometry")
            if isinstance(geo, str) and "x" in geo:
                try:
                    self.geometry(geo)
                except Exception:
                    pass
            if "_log_open" in data:
                self._log_box.set_open(bool(data["_log_open"]))
            mode = data.get("_mode")
            if mode in MODE_BY_KEY:
                self.source_mode.set(mode)
            # ค่าที่เก็บเป็น set → บันทึกเป็น list
            sel = data.get("_req_selected")
            if isinstance(sel, list):
                self._req_selected_codes = {c for c in sel if c in self._req_codes}
                self._render_request_listbox("")
                self._on_request_changed()
            bsel = data.get("_booking_selected")
            if isinstance(bsel, list):
                self._booking_selected_codes = set(bsel)
                self._booking_render_listbox(self.booking_search.get())
            self._on_etk_multi_changed()
            self._on_bt44_dryrun_toggle()
        except Exception:
            pass

    def _save_settings(self) -> None:
        try:
            data: dict = {}
            for name, var in self._persist.items():
                try:
                    data[name] = var.get()
                except Exception:
                    pass
            data["_geometry"] = self.geometry()
            data["_log_open"] = self._log_box.is_open()
            data["_mode"] = self.source_mode.get()
            self._req_persist_selection()
            data["_req_selected"] = sorted(self._req_selected_codes)
            try:
                self._booking_persist_selection()
            except Exception:
                pass
            data["_booking_selected"] = sorted(self._booking_selected_codes)
            SETTINGS_FILE.write_text(
                json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8",
            )
        except Exception:
            pass

    def _on_close(self) -> None:
        if self._worker is not None and self._worker.is_alive():
            if not messagebox.askyesno(
                "งานยังทำอยู่",
                "ยังมีงานทำงานอยู่ — ปิดโปรแกรมจะยกเลิกงานทันที\nต้องการปิดหรือไม่?",
                icon="warning",
            ):
                return
            self._cancel.set()
            self._pause_evt.set()
        self._save_settings()
        self.destroy()

    # ---------- small UI helpers ----------
    def _file_row(self, parent, row: int, label: str, var: tk.StringVar,
                  hint: str = "", clearable: bool = False, save: bool = False):
        """แถว: label | Entry | [เลือก...] [ล้าง]  (+ hint ใต้แถวถ้ามี)"""
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=3)
        ent = ttk.Entry(parent, textvariable=var)
        ent.grid(row=row, column=1, sticky="we", padx=8, pady=3)
        box = ttk.Frame(parent)
        box.grid(row=row, column=2, sticky="w", pady=3)
        if save:
            btn = ttk.Button(box, text="เลือก...", width=9, command=self._choose_out)
        else:
            btn = ttk.Button(box, text="เลือก...", width=9,
                             command=lambda v=var: self._choose_into(v))
        btn.pack(side="left")
        if clearable:
            ttk.Button(box, text="ล้าง", width=5, command=lambda v=var: v.set("")).pack(
                side="left", padx=(4, 0))
        if hint:
            Tooltip(ent, hint)
        return ent, btn

    def _text_row(self, parent, row: int, label: str, var: tk.Variable,
                  hint: str = "", width: int = 18):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=3)
        ent = ttk.Entry(parent, textvariable=var, width=width)
        ent.grid(row=row, column=1, sticky="w", padx=8, pady=3)
        if hint:
            ttk.Label(parent, text=hint, style="Hint.TLabel").grid(
                row=row, column=2, sticky="w")
        return ent

    def _note(self, parent, text: str, style: str = "Hint.TLabel", row: int | None = None,
              wrap: int = 760, pady=(2, 0)):
        lbl = ttk.Label(parent, text=text, style=style, wraplength=wrap, justify="left")
        if row is None:
            lbl.pack(anchor="w", pady=pady)
        else:
            lbl.grid(row=row, column=0, columnspan=3, sticky="w", pady=pady)
        return lbl

    def _card(self, title: str) -> ttk.LabelFrame:
        """สร้างการ์ดใน content area (ยังไม่ pack — _on_mode_changed จัดการ)"""
        return ttk.LabelFrame(self._body, text=title)

    # ---------- UI ----------
    def _build_ui(self) -> None:
        P = PALETTE
        self.source_mode = tk.StringVar(value="etracking")

        # ================= Header =================
        header = ttk.Frame(self, style="Header.TFrame", padding=(14, 8))
        header.pack(fill="x")
        ttk.Label(header, text="e-WorkPermit Tools", style="Title.TLabel").pack(side="left")
        ttk.Label(header, text="  ดึงรายงาน · ดาวน์โหลดเอกสาร · ยื่นคำขอ",
                  style="Desc.TLabel").pack(side="left", pady=(4, 0))
        # ชื่อผู้ใช้งาน — ใช้แยกยอดการใช้งานต่อคน (บันทึกลง MySQL)
        ttk.Label(header, text="ผู้ใช้งาน:", style="Surface.TLabel").pack(side="left", padx=(28, 4), pady=(3, 0))
        self.operator = self._remember("operator", tk.StringVar(value=usage_tracker.default_operator()))
        op_ent = ttk.Entry(header, textvariable=self.operator, width=16)
        op_ent.pack(side="left", pady=(2, 0))
        Tooltip(op_ent, "ชื่อที่ใช้แยกยอดการใช้งานเพื่อคิดค่าใช้จ่าย — ใส่ชื่อตัวเอง/รหัสหน่วยงาน (จำค่าไว้ให้)")
        ttk.Button(header, text="สถานะการบันทึกใช้งาน", style="Link.TButton",
                   command=self._usage_status).pack(side="left", padx=(4, 0), pady=(2, 0))
        # ตัวเลือกส่วนกลาง (ใช้ทุกโหมด)
        self.headless = tk.BooleanVar(value=False)   # headless จริงซ่อนจาก UI (ใช้ซ่อนนอกจอแทน)
        self.hide_window = self._remember("hide_window", tk.BooleanVar(value=False))
        cb = ttk.Checkbutton(header, text="ซ่อนหน้าต่างเบราว์เซอร์ขณะทำงาน",
                             variable=self.hide_window, style="Surface.TCheckbutton")
        cb.pack(side="right")
        Tooltip(cb, "เปิดเบราว์เซอร์จริงแต่ย้ายออกนอกจอ — ไม่รบกวนการใช้งานเครื่อง และเว็บตรวจไม่พบว่าเป็น headless")
        ttk.Separator(self, orient="horizontal").pack(fill="x")

        # ================= Main split: sidebar | content =================
        main = ttk.Frame(self)
        main.pack(fill="both", expand=True)

        # ---- Sidebar ----
        side = tk.Frame(main, bg=P["sidebar"], width=250)
        side.pack(side="left", fill="y")
        side.pack_propagate(False)
        tk.Label(side, text="โหมดการทำงาน", bg=P["sidebar"], fg="#9CA3AF",
                 font=FONT_SMALL, anchor="w", padx=12, pady=8).pack(fill="x")
        self.nav = ttk.Treeview(side, show="tree", style="Nav.Treeview", selectmode="browse")
        self.nav.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        self._nav_group_ids: dict[str, str] = {}
        for gkey, glabel in MODE_GROUPS:
            gid = self.nav.insert("", "end", iid=f"grp:{gkey}", text=glabel, open=True)
            self._nav_group_ids[gkey] = gid
        for m in MODES:
            self.nav.insert(self._nav_group_ids[m["group"]], "end", iid=f"mode:{m['key']}",
                            text="    " + m["nav"])
        self.nav.bind("<<TreeviewSelect>>", self._on_nav_select)
        tk.Label(side, text="F5 = เริ่มทำงาน", bg=P["sidebar"], fg="#6B7280",
                 font=FONT_SMALL, anchor="w", padx=12, pady=6).pack(fill="x", side="bottom")

        # ---- Content (right) ----
        right = ttk.Frame(main)
        right.pack(side="left", fill="both", expand=True)

        # หัวข้อโหมด
        head = ttk.Frame(right, style="Header.TFrame", padding=(16, 10, 16, 8))
        head.pack(fill="x")
        hrow = ttk.Frame(head, style="Header.TFrame")
        hrow.pack(fill="x")
        self.mode_title = ttk.Label(hrow, text="", style="Title.TLabel")
        self.mode_title.pack(side="left")
        self.mode_badge = ttk.Label(hrow, text="", style="Safe.Badge.TLabel")
        self.mode_badge.pack(side="left", padx=(12, 0), pady=(4, 0))
        self.mode_desc = ttk.Label(head, text="", style="Desc.TLabel", wraplength=820,
                                   justify="left")
        self.mode_desc.pack(fill="x", pady=(2, 0))
        ttk.Separator(right, orient="horizontal").pack(fill="x")

        # พื้นที่ตัวเลือก — เลื่อนได้
        outer = ttk.Frame(right)
        outer.pack(fill="both", expand=True)
        self._scroll_canvas = tk.Canvas(outer, borderwidth=0, highlightthickness=0,
                                        bg=P["bg"])
        _vsb = ttk.Scrollbar(outer, orient="vertical", command=self._scroll_canvas.yview)
        self._scroll_canvas.configure(yscrollcommand=_vsb.set)
        _vsb.pack(side="right", fill="y")
        self._scroll_canvas.pack(side="left", fill="both", expand=True)
        self._body = ttk.Frame(self._scroll_canvas, padding=(12, 8, 12, 8))
        self._body_window = self._scroll_canvas.create_window(
            (0, 0), window=self._body, anchor="nw",
        )

        def _on_body_configure(_evt=None):
            self._scroll_canvas.configure(scrollregion=self._scroll_canvas.bbox("all"))

        def _on_canvas_configure(evt):
            self._scroll_canvas.itemconfig(self._body_window, width=evt.width)

        self._body.bind("<Configure>", _on_body_configure)
        self._scroll_canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(evt):
            # เลื่อนเฉพาะเมื่อเมาส์อยู่ในหน้าต่างหลัก (ไม่ใช่ Monitor/Dialog) และไม่อยู่บน Listbox/Text
            try:
                w = evt.widget
                if not isinstance(w, tk.Misc) or w.winfo_toplevel() is not self:
                    return
                p = w
                while p is not None:
                    if isinstance(p, (tk.Listbox, tk.Text, ttk.Treeview)):
                        return
                    if p is self._scroll_canvas:
                        break
                    p = getattr(p, "master", None)
                if p is None:
                    return  # อยู่นอกพื้นที่ตัวเลือก (เช่น sidebar/log)
                self._scroll_canvas.yview_scroll(int(-1 * (evt.delta / 120)), "units")
            except Exception:
                pass

        self.bind_all("<MouseWheel>", _on_mousewheel)

        # ================= Bottom bar (ปุ่ม + progress) =================
        ttk.Separator(right, orient="horizontal").pack(fill="x", side="bottom")
        self._log_box = Collapsible(right, "บันทึกการทำงาน", open=True)
        self._log_box.pack(fill="both", side="bottom", padx=10, pady=(0, 6))
        self._log_box.body.configure(padding=(0, 2, 0, 0))
        logf = ttk.Frame(self._log_box.body)
        logf.pack(fill="both", expand=True)
        self.log_text = tk.Text(logf, height=8, wrap="word", state="disabled", font=FONT_MONO,
                                bg=P["surface"], relief="flat", highlightthickness=1,
                                highlightbackground=P["border"], padx=6, pady=4)
        self.log_text.pack(fill="both", expand=True, side="left")
        sb = ttk.Scrollbar(logf, command=self.log_text.yview)
        sb.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=sb.set)
        self.log_text.tag_configure("err", foreground=P["danger"])
        self.log_text.tag_configure("ok", foreground=P["ok"])
        self.log_text.tag_configure("warn", foreground=P["warn"])
        ttk.Button(self._log_box.body, text="ล้าง log", style="Link.TButton",
                   command=self._clear_log).pack(anchor="e")

        bottom = ttk.Frame(right, style="Bottom.TFrame", padding=(12, 8))
        bottom.pack(fill="x", side="bottom")
        self.btn_start = ttk.Button(bottom, text="▶  เริ่มทำงาน", style="Accent.TButton",
                                    command=self._on_start)
        self.btn_start.pack(side="left")
        self.btn_pause = ttk.Button(bottom, text="⏸  พัก", command=self._on_pause, state="disabled")
        self.btn_pause.pack(side="left", padx=(8, 0))
        self.btn_resume = ttk.Button(bottom, text="▶  ต่อ", command=self._on_resume, state="disabled")
        self.btn_resume.pack(side="left", padx=(4, 0))
        self.btn_cancel = ttk.Button(bottom, text="✕  ยกเลิก", command=self._on_cancel,
                                     state="disabled", style="Danger.TButton")
        self.btn_cancel.pack(side="left", padx=(4, 0))
        self.btn_open = ttk.Button(bottom, text="📁  เปิดโฟลเดอร์ผลลัพธ์", command=self._open_folder)
        self.btn_open.pack(side="right")
        self.progress_lbl = ttk.Label(bottom, text="พร้อมทำงาน", style="Surface.TLabel", width=16,
                                      anchor="e")
        self.progress_lbl.pack(side="right", padx=(8, 12))
        self.progress = ttk.Progressbar(bottom, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=(16, 0))

        # ================= Cards (สร้างทั้งหมด แล้วให้ _on_mode_changed เลือกแสดง) =================
        self._build_login_card()
        self._build_etracking_cards()
        self._build_output_card()
        self._build_mode_cards()

        # เลือกโหมดเริ่มต้นใน sidebar
        self._select_nav(self.source_mode.get())

    # ---------- การ์ด: ข้อมูลเข้าสู่ระบบ ----------
    def _build_login_card(self) -> None:
        frm = self._card("บัญชีเข้าสู่ระบบ")
        self.login_frame = frm
        self.login_hint = ttk.Label(frm, text="", style="Info.TLabel", wraplength=760, justify="left")
        self.login_hint.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 4))

        ttk.Label(frm, text="Username (อีเมล):").grid(row=1, column=0, sticky="w", pady=3)
        self.username = self._remember("username", tk.StringVar())
        ttk.Entry(frm, textvariable=self.username, width=44).grid(row=1, column=1, sticky="we", padx=8)

        ttk.Label(frm, text="Password:").grid(row=2, column=0, sticky="w", pady=3)
        self.password = tk.StringVar()
        self.pw_entry = ttk.Entry(frm, textvariable=self.password, width=44, show="•")
        self.pw_entry.grid(row=2, column=1, sticky="we", padx=8)
        self.pw_entry.bind("<Return>", lambda _e: self._on_start())
        self.show_pw = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm, text="แสดงรหัสผ่าน", variable=self.show_pw,
                        command=self._toggle_pw).grid(row=2, column=2, sticky="w")

        ttk.Label(frm, text="ประเภทผู้ใช้:").grid(row=3, column=0, sticky="w", pady=3)
        self.user_type = self._remember("user_type", tk.StringVar(value="ผู้กระทำการแทน"))
        ttk.Combobox(
            frm, textvariable=self.user_type, width=42, state="readonly",
            values=["ผู้กระทำการแทน", "นายจ้าง-คนไทย", "นายจ้าง-คนต่างชาติ",
                    "บริษัทนำเข้า (บนจ.)", "คนต่างด้าว"],
        ).grid(row=3, column=1, sticky="we", padx=8)

        ttk.Label(frm, text="ระบบ:").grid(row=4, column=0, sticky="w", pady=3)
        self.method = self._remember("method", tk.StringVar(value="E-Workpermit"))
        ttk.Combobox(frm, textvariable=self.method, width=42, state="readonly",
                     values=["E-Workpermit", "E-Service"]).grid(row=4, column=1, sticky="we", padx=8)
        frm.columnconfigure(1, weight=1)

    # ---------- การ์ด: ไฟล์ผลลัพธ์ + การทำงาน ----------
    def _build_output_card(self) -> None:
        frm = self._card("ไฟล์ผลลัพธ์")
        self.output_frame = frm
        self.out_path = tk.StringVar(value=str(REPORTS_DIR / "WA_report.xlsx"))
        self._file_row(frm, 0, "บันทึกรายงานเป็น:", self.out_path, save=True,
                       hint="ระบบจะต่อท้ายวันที่-เวลาให้อัตโนมัติ กันชื่อซ้ำ")
        # แถวตัวเลือกการทำงาน (เฉพาะ e-Tracking / permit_report)
        self.run_opts_row = ttk.Frame(frm)
        self.run_opts_row.grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 0))
        ttk.Label(self.run_opts_row, text="จำกัดจำนวน:").pack(side="left")
        self.limit = tk.IntVar(value=0)
        sp = ttk.Spinbox(self.run_opts_row, from_=0, to=10000, textvariable=self.limit, width=7)
        sp.pack(side="left", padx=(4, 4))
        ttk.Label(self.run_opts_row, text="(0 = ทั้งหมด — ใส่ 5 เพื่อทดสอบ)", style="Hint.TLabel").pack(side="left")
        ttk.Label(self.run_opts_row, text="Login ใหม่ทุก:").pack(side="left", padx=(18, 0))
        self.relogin_every = self._remember("relogin_every", tk.StringVar(value="500"))
        cbx = ttk.Combobox(self.run_opts_row, textvariable=self.relogin_every, width=9, state="normal",
                           values=["ปิด (ไม่ login ใหม่)", "100", "200", "300", "500", "1000", "2000"])
        cbx.pack(side="left", padx=(4, 4))
        ttk.Label(self.run_opts_row, text="รายการ (กัน session หมดอายุ — 0/ว่าง = ปิด)",
                  style="Hint.TLabel").pack(side="left")
        frm.columnconfigure(1, weight=1)

    # ---------- การ์ดชุด e-Tracking (ใช้ร่วมกับ bt30_ctn / permit_report) ----------
    def _build_etracking_cards(self) -> None:
        # ===== 1) รายการคำขอ + วันที่ยื่น =====
        opt = self._card("รายการคำขอ และ วันที่ยื่นคำขอ")
        self.req_frame = opt
        self._etracking_frames: list[ttk.LabelFrame] = [opt]

        _default_code = "MT_59_MOU_RENEWAL"
        # เก็บ StringVar สำหรับ backward compat (ยังใช้กับ mode อื่นที่รับ request_type เดียว)
        self.request_type = tk.StringVar(value=_default_code)
        self.request_label = tk.StringVar(value="")

        rt_head = ttk.Frame(opt)
        rt_head.grid(row=0, column=0, columnspan=3, sticky="we", pady=(0, 2))
        ttk.Label(rt_head, text="เลือกรายการคำขอ (เลือกได้หลายรายการ — Ctrl/Shift + คลิก)").pack(side="left")
        ttk.Button(rt_head, text="เลือกทั้งหมด", command=self._req_select_all, width=12).pack(
            side="right", padx=(4, 0))
        ttk.Button(rt_head, text="ล้าง", command=self._req_clear, width=6).pack(side="right")

        self._req_codes: list[str] = [c for c, _ in REQUEST_TYPES if c]
        self._req_labels: list[str] = [l for c, l in REQUEST_TYPES if c]
        self._req_visible_idx: list[int] = []
        self._req_selected_codes: set[str] = {_default_code}
        rt_box = ttk.Frame(opt)
        rt_box.grid(row=1, column=0, columnspan=3, sticky="we", pady=(0, 4))
        rt_box.columnconfigure(0, weight=1)
        rt_search_row = ttk.Frame(rt_box)
        rt_search_row.grid(row=0, column=0, columnspan=2, sticky="we", pady=(0, 3))
        ttk.Label(rt_search_row, text="🔍").pack(side="left")
        self.request_search = tk.StringVar(value="")
        ttk.Entry(rt_search_row, textvariable=self.request_search).pack(
            side="left", fill="x", expand=True, padx=(4, 4))
        self.request_search.trace_add("write", lambda *_: self._req_apply_search())
        ttk.Button(rt_search_row, text="✕", width=3,
                   command=lambda: self.request_search.set("")).pack(side="left")
        vsb = ttk.Scrollbar(rt_box, orient="vertical")
        hsb = ttk.Scrollbar(rt_box, orient="horizontal")
        self.request_listbox = tk.Listbox(
            rt_box, selectmode=tk.EXTENDED, height=7, exportselection=False,
            yscrollcommand=vsb.set, xscrollcommand=hsb.set, activestyle="none",
            bg=PALETTE["surface"], relief="flat", highlightthickness=1,
            highlightbackground=PALETTE["border"], selectbackground=PALETTE["accent"],
            selectforeground="#FFFFFF", font=FONT_UI,
        )
        vsb.config(command=self.request_listbox.yview)
        hsb.config(command=self.request_listbox.xview)
        self.request_listbox.grid(row=1, column=0, sticky="we")
        vsb.grid(row=1, column=1, sticky="ns")
        hsb.grid(row=2, column=0, sticky="we")
        self.request_listbox.bind("<<ListboxSelect>>", self._on_request_changed)
        self._render_request_listbox("")
        try:
            _idx = next((r for r, m in enumerate(self._req_visible_idx)
                         if self._req_codes[m] == _default_code), None)
            if _idx is not None:
                self.request_listbox.see(_idx)
        except Exception:
            pass

        self.request_summary = tk.StringVar(value="")
        ttk.Label(opt, textvariable=self.request_summary, style="Ok.TLabel").grid(
            row=2, column=0, columnspan=3, sticky="w")

        date_row = ttk.Frame(opt)
        date_row.grid(row=3, column=0, columnspan=3, sticky="we", pady=(6, 0))
        ttk.Label(date_row, text="วันที่ยื่นคำขอ  จาก:").pack(side="left")
        self.date_from = tk.StringVar(value="")
        ttk.Entry(date_row, textvariable=self.date_from, width=13).pack(side="left", padx=(4, 10))
        ttk.Label(date_row, text="ถึง:").pack(side="left")
        self.date_to = tk.StringVar(value="")
        ttk.Entry(date_row, textvariable=self.date_to, width=13).pack(side="left", padx=(4, 10))
        ttk.Label(date_row, text="วว/ดด/ปปปป เช่น 01/07/2026 — เว้นว่าง = ทุกวันที่",
                  style="Hint.TLabel").pack(side="left")
        opt.columnconfigure(1, weight=1)

        # ===== 2) ตัวกรองสถานะ =====
        adv = self._card("ตัวกรองสถานะ (ตามที่เห็นในเว็บ)")
        self.status_frame = adv
        self._etracking_frames.append(adv)
        ttk.Label(adv, text="สถานะที่จะดึง:").grid(row=0, column=0, sticky="w")
        st_row = ttk.Frame(adv)
        st_row.grid(row=0, column=1, columnspan=4, sticky="w")
        self.filter_vars: dict[str, tk.BooleanVar] = {}
        for i, (code, label) in enumerate(STATUS_FILTERS):
            v = tk.BooleanVar(value=False)
            self.filter_vars[code] = v
            ttk.Checkbutton(st_row, text=label, variable=v).pack(side="left", padx=(0, 12), pady=2)

        self.status_adv = Collapsible(
            adv, "ตัวเลือกขั้นสูง: เปิดดูรายละเอียดเฉพาะสถานะ / บันทึกเพิ่มเติม", open=False,
            hint="(ปกติไม่ต้องแก้ — ค่าเริ่มต้นตั้งตามรายการคำขอให้แล้ว)")
        self.status_adv.grid(row=1, column=0, columnspan=5, sticky="we", pady=(6, 0))
        wl = self.status_adv.body
        ttk.Label(wl, text="กดเข้าไปดูรายละเอียดเฉพาะคำขอที่มีสถานะเหล่านี้ (match แบบ substring):",
                  style="Muted.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        self.whitelist_vars: dict[str, tk.BooleanVar] = {}
        for i, status in enumerate(WHITELIST_OPTIONS):
            v = tk.BooleanVar(value=True)
            self.whitelist_vars[status] = v
            ttk.Checkbutton(wl, text=status, variable=v).grid(
                row=1 + i // 2, column=i % 2, sticky="w", padx=(0, 16), pady=1)
        self.capture_extra_notes = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            wl, text="เก็บ 'บันทึกเพิ่มเติม' จากหน้าแก้ไข (ปิดถ้าหน้าไม่มีปุ่มแก้ไข)",
            variable=self.capture_extra_notes,
        ).grid(row=1 + (len(WHITELIST_OPTIONS) + 1) // 2, column=0, columnspan=2, sticky="w", pady=(6, 0))

        # ===== 3) หลายบัญชี (etracking / permit_report) =====
        mf = self._card("บัญชีที่ใช้ดึง")
        self.multi_frame = mf
        self.etk_multi = self._remember("etk_multi", tk.BooleanVar(value=False))
        ttk.Checkbutton(
            mf, text="ดึงหลายบัญชีจากไฟล์ UsernameLogin.xlsx (ไม่ใช้ Username/Password ด้านบน)",
            variable=self.etk_multi, command=self._on_etk_multi_changed,
        ).grid(row=0, column=0, columnspan=3, sticky="w")
        self.etk_login_input = self._remember("etk_login_input",
                                              tk.StringVar(value=str(ROOT / "UsernameLogin.xlsx")))
        self.etk_login_entry, self.etk_login_btn = self._file_row(
            mf, 1, "ไฟล์ UsernameLogin.xlsx:", self.etk_login_input,
            hint="คอลัมน์ที่ต้องมี: Username, Password, Type")
        self.etk_combine = self._remember("etk_combine", tk.BooleanVar(value=False))
        self.etk_combine_cb = ttk.Checkbutton(
            mf, text="รวมทุกบัญชีเป็นไฟล์เดียว (มีคอลัมน์ 'บัญชี (Username)') — ไม่ติ๊ก = แยกไฟล์ต่อบัญชี",
            variable=self.etk_combine)
        self.etk_combine_cb.grid(row=2, column=0, columnspan=3, sticky="w")
        mf.columnconfigure(1, weight=1)

        # ===== 4) รูปแบบรายงาน (etracking เท่านั้น) =====
        tf = self._card("รูปแบบรายงาน (Template)")
        self.etk_template_frame = tf
        self.etk_template = self._remember("etk_template", tk.StringVar(value="main"))
        ttk.Radiobutton(tf, text="Template หลัก — รายละเอียดครบทุกคอลัมน์ (ค่าเริ่มต้น)",
                        variable=self.etk_template, value="main").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(
            tf, text="Template นัดหมาย — ลำดับ | Username | เลขคำขอ | ชื่อบริษัท | ชื่อแรงงาน | วันที่นัด | เวลา | สถานที่",
            variable=self.etk_template, value="appointment").grid(row=1, column=0, sticky="w", pady=(2, 0))
        self._note(tf, "Template นัดหมาย จะเข้าแท็บ 'การนัดหมาย' ของแต่ละรายการ (ช้าลงเล็กน้อย) — "
                       "ควรติ๊กสถานะ 'รอนัดหมาย' ในตัวกรองสถานะ", row=2)

        # ===== 5) ดึงเฉพาะเลขคำขอ (etracking เท่านั้น — พับไว้) =====
        rf = self._card("ดึงเฉพาะเลขคำขอ (ทางเลือก)")
        self.etk_req_no_frame = rf
        self.etk_req_no_box = Collapsible(
            rf, "ระบุเลขคำขอเอง — ข้ามตัวกรองสถานะ/วันที่ด้านบนทั้งหมด", open=False,
            hint="(เว้นว่าง = ดึงตามตัวกรองปกติ)")
        self.etk_req_no_box.pack(fill="x")
        rb = self.etk_req_no_box.body
        ttk.Label(rb, text="โหมด A: ไฟล์ Excel เลขคำขอ — รองรับหลายบัญชี (จัดกลุ่มตาม username แล้ว login ทีละบัญชี)",
                  style="Section.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 2))
        self.etk_req_excel_input = self._remember("etk_req_excel_input", tk.StringVar(value=""))
        self._file_row(rb, 1, "ไฟล์ Ref_number.xlsx:", self.etk_req_excel_input, clearable=True,
                       hint="คอลัมน์ที่ต้องมี: Ref_number (เลขคำขอ), user (Username เจ้าของคำขอ)")
        self.etk_req_login_input = self._remember("etk_req_login_input",
                                                  tk.StringVar(value=str(ROOT / "UsernameLogin.xlsx")))
        self._file_row(rb, 2, "ไฟล์ UsernameLogin.xlsx:", self.etk_req_login_input,
                       hint="คอลัมน์ที่ต้องมี: Username, Password, Type")
        ttk.Separator(rb, orient="horizontal").grid(row=3, column=0, columnspan=3, sticky="we", pady=6)
        ttk.Label(rb, text="โหมด B: พิมพ์/วางเลขคำขอเอง (ใช้บัญชีเดียวจากด้านบน — ใช้เมื่อไม่ระบุไฟล์โหมด A)",
                  style="Section.TLabel").grid(row=4, column=0, columnspan=3, sticky="w", pady=(0, 2))
        self.etk_req_no_text = tk.Text(rb, height=3, width=60, font=FONT_UI, relief="flat",
                                       highlightthickness=1, highlightbackground=PALETTE["border"])
        self.etk_req_no_text.grid(row=5, column=0, columnspan=3, sticky="we")
        rb.columnconfigure(1, weight=1)

        # โหลด default ตาม request_type แรก
        self._apply_profile_defaults(self.request_type.get())
        self._update_request_summary()

    # ---------- การ์ดเฉพาะโหมด ----------
    def _build_mode_cards(self) -> None:
        self._build_card_aliens()
        self._build_card_register()
        self._build_card_receipts()
        self._build_card_results()
        self._build_card_inform()
        self._build_card_bt30()
        self._build_card_bt44()
        self._build_card_billpay()
        self._build_card_payrcpt()
        self._build_card_appointment()
        self._build_card_bt30ctn()
        self._build_card_namelist()
        self._build_card_ticket()
        self._build_card_booking()

    def _build_card_aliens(self) -> None:
        f = self._card("ตัวเลือก — ข้อมูลคนต่างด้าว")
        self.aliens_frame = f
        ttk.Label(f, text="เลือกหมวดข้อมูลที่จะดึง:").pack(anchor="w")
        self.alien_subtab_vars: dict[str, tk.BooleanVar] = {}
        for name in ALIEN_SUB_TABS:
            v = tk.BooleanVar(value=True)
            self.alien_subtab_vars[name] = v
            ttk.Checkbutton(f, text=name, variable=v).pack(anchor="w", padx=8, pady=2)

    def _build_card_register(self) -> None:
        f = self._card("ตัวเลือก — ลงทะเบียนคนต่างด้าว")
        self.register_frame = f
        self.register_input = self._remember("register_input", tk.StringVar(value=str(ROOT / "Exm.xlsx")))
        self._file_row(f, 0, "ไฟล์ Excel input:", self.register_input,
                       hint="คอลัมน์ที่ต้องมี: No., Name, TaxID, PassportNo., Password, Email")
        self.register_row_range = tk.StringVar(value="")
        self._text_row(f, 1, "แถวที่จะลงทะเบียน:", self.register_row_range,
                       'เช่น "4-13" หรือ "1,3,5-8" — เว้นว่าง = ทุกแถว (นับแถวข้อมูล ไม่นับหัวตาราง)')
        self._note(f, "ใช้บัญชีของคนต่างด้าวแต่ละคนจาก Excel (Password/Email) — ไม่ต้องกรอก Username/Password",
                   row=2, pady=(6, 0))
        f.columnconfigure(1, weight=1)

    def _build_card_receipts(self) -> None:
        f = self._card("ตัวเลือก — ดาวน์โหลดใบเสร็จ / แบบฟอร์ม")
        self.receipt_frame = f
        self.receipt_request_input = self._remember(
            "receipt_request_input", tk.StringVar(value=str(ROOT / "RequestData.xlsx")))
        self._file_row(f, 0, "ไฟล์ RequestData.xlsx:", self.receipt_request_input,
                       hint="คอลัมน์ที่ต้องมี: ลำดับ, ชื่อคนต่างด้าว(Eng), เลขที่คำขอ, Username, PASSPORT NUMBER")
        self.receipt_login_input = self._remember(
            "receipt_login_input", tk.StringVar(value=str(ROOT / "UsernameLogin.xlsx")))
        self._file_row(f, 1, "ไฟล์ UsernameLogin.xlsx:", self.receipt_login_input,
                       hint="คอลัมน์ที่ต้องมี: Username, Password, Type — จัดกลุ่มตาม Username แล้ว login อัตโนมัติ")
        self.receipt_row_range = tk.StringVar(value="")
        self._text_row(f, 2, "แถวที่จะดาวน์โหลด:", self.receipt_row_range,
                       'เช่น "1-50" หรือ "2,5,8" — เว้นว่าง = ทุกแถว')

        ttk.Label(f, text="เอกสารที่จะดาวน์โหลด:").grid(row=3, column=0, sticky="nw", pady=(6, 2))
        doc_box = ttk.Frame(f)
        doc_box.grid(row=3, column=1, columnspan=2, sticky="w", pady=(6, 2))
        for i, (text, var) in enumerate((
            ("ใบเสร็จ (ทุกใบ)", self.doc_receipt), ("แบบ บต.44", self.doc_bt44),
            ("แบบ บต.22", self.doc_bt22), ("แบบ บต.53", self.doc_bt53),
            ("แบบ บต.55", self.doc_bt55), ("แบบ บต.52", self.doc_bt52),
            ("แบบ บต.56", self.doc_bt56),
        )):
            ttk.Checkbutton(doc_box, text=text, variable=var).grid(
                row=i // 4, column=i % 4, sticky="w", padx=(0, 14), pady=1)

        self.receipt_name_suffix = tk.StringVar(value="")
        self._text_row(f, 4, "ต่อท้ายชื่อไฟล์:", self.receipt_name_suffix,
                       'เว้นว่าง = ไม่ต่อท้าย | ใส่ "_IO" → {PASSPORT}_BT22_IO.pdf', width=14)
        self.receipt_make_folder = self._remember("receipt_make_folder", tk.BooleanVar(value=False))
        ttk.Checkbutton(f, text="สร้างโฟลเดอร์แยกตาม PASSPORT (receipts/MD1524123/MD1524123_BT44.pdf)",
                        variable=self.receipt_make_folder).grid(
            row=5, column=0, columnspan=3, sticky="w", pady=(4, 2))
        self._note(f, "PDF เก็บใน reports/receipts ตั้งชื่อตามราคา {PASSPORT}_RECEIPT100.pdf / _BT44.pdf ฯลฯ  •  "
                      "รันซ้ำได้: ไฟล์ที่มีอยู่แล้วจะถูกข้าม", row=6)
        f.columnconfigure(1, weight=1)

    def _build_card_results(self) -> None:
        f = self._card("ตัวเลือก — ใบแจ้งผล / ใบรับคำขอ")
        self.result_frame = f
        ttk.Label(f, text="โหมด A — ตามเลขคำขอจาก Ref_number.xlsx (ถ้าระบุไฟล์)", style="Section.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 2))
        self.result_ref_input = self._remember(
            "result_ref_input", tk.StringVar(value=str(ROOT / "Ref_number.xlsx")))
        self._file_row(f, 1, "ไฟล์ Ref_number.xlsx:", self.result_ref_input, clearable=True,
                       hint="คอลัมน์ที่ต้องมี: Ref_number (เลขคำขอ), user (Username เจ้าของคำขอ) — กด 'ล้าง' เพื่อใช้โหมด B")
        self.result_login_input = self._remember(
            "result_login_input", tk.StringVar(value=str(ROOT / "UsernameLogin.xlsx")))
        self._file_row(f, 2, "ไฟล์ UsernameLogin.xlsx:", self.result_login_input,
                       hint="คอลัมน์ที่ต้องมี: Username, Password, Type")

        ttk.Separator(f, orient="horizontal").grid(row=3, column=0, columnspan=3, sticky="we", pady=6)
        ttk.Label(f, text="โหมด B — ดึงจาก e-Tracking (เมื่อช่อง Ref_number ว่าง) สถานะ 'รอนัดหมาย'",
                  style="Section.TLabel").grid(row=4, column=0, columnspan=3, sticky="w", pady=(0, 2))
        ttk.Label(f, text="รายการคำขอ:").grid(row=5, column=0, sticky="w", pady=3)
        self.result_request_type = tk.StringVar(value="")
        self.result_request_label = tk.StringVar(value=REQUEST_TYPES[0][1])
        rr_combo = ttk.Combobox(f, textvariable=self.result_request_label, state="readonly",
                                values=[lbl for _, lbl in REQUEST_TYPES])
        rr_combo.grid(row=5, column=1, columnspan=2, sticky="we", padx=8)
        rr_combo.bind("<<ComboboxSelected>>", self._on_result_request_changed)
        self._text_row(f, 6, "หรือใส่ code เอง:", self.result_request_type, "(เว้นว่าง = ทั้งหมด)", width=28)

        ttk.Label(f, text="เอกสารที่จะดาวน์โหลด:").grid(row=7, column=0, sticky="w", pady=(6, 2))
        rdoc = ttk.Frame(f)
        rdoc.grid(row=7, column=1, columnspan=2, sticky="w", pady=(6, 2))
        ttk.Checkbutton(rdoc, text="ใบแจ้งผลใบอนุญาตทำงาน", variable=self.doc_result_notice).pack(side="left", padx=(0, 14))
        ttk.Checkbutton(rdoc, text="ใบรับคำขอใบอนุญาตทำงาน", variable=self.doc_request_receipt).pack(side="left", padx=(0, 14))
        ttk.Checkbutton(rdoc, text="แบบ บต.50 อ.6", variable=self.doc_result_bt50).pack(side="left")
        self._note(f, "PDF เก็บใน reports/result_docs ชื่อ {ชื่อคนต่างด้าว(Eng)}_{ชื่อเอกสาร}.pdf  •  รันซ้ำได้: ไฟล์ที่มีอยู่แล้วจะถูกข้าม",
                   row=8)
        f.columnconfigure(1, weight=1)

    def _build_card_inform(self) -> None:
        f = self._card("ตัวเลือก — แจ้งเข้านายจ้าง (บต.52)")
        self.inform_frame = f
        self.inform_excel_input = self._remember(
            "inform_excel_input", tk.StringVar(value=str(ROOT / "FormRequestEmployment.xlsx")))
        self._file_row(f, 0, "ไฟล์ FormRequestEmployment.xlsx:", self.inform_excel_input,
                       hint="คอลัมน์ที่ต้องมี: ประเภทนายจ้าง, รหัสนายจ้าง, เลขใบอนุญาตทำงาน, วันที่เริ่มงาน, Files\n"
                            "ไฟล์แนบ (คอลัมน์ Files) ใช้ relative path จากโฟลเดอร์เดียวกับ Excel")
        self.inform_login_input = self._remember(
            "inform_login_input", tk.StringVar(value=str(ROOT / "UsernameLogin.xlsx")))
        self._file_row(f, 1, "ไฟล์ UsernameLogin.xlsx:", self.inform_login_input,
                       hint="ใช้บัญชีแรกที่เจอ (ควรเป็นบัญชี บนจ. ที่ตรงกับฟอร์มนี้)")
        self.inform_row_range = tk.StringVar(value="")
        self._text_row(f, 2, "แถวที่จะทำ:", self.inform_row_range, 'เช่น "1-5" หรือ "2,4" — เว้นว่าง = ทุกแถว')
        self.inform_commit = tk.BooleanVar(value=False)
        ttk.Checkbutton(f, text="⚠ ยืนยันส่งคำขอจริง (ทำข้อ 7+8) — ไม่ติ๊ก = โหมดทดสอบ หยุดก่อนข้อ 7",
                        variable=self.inform_commit).grid(row=3, column=0, columnspan=3, sticky="w", pady=(8, 0))
        self._note(f, "โหมดทดสอบจะเก็บ screenshot หน้าสรุปคำขอใน inform_screenshots/ โดยไม่ส่งจริง",
                   style="Warn.TLabel", row=4)
        f.columnconfigure(1, weight=1)

    def _build_card_bt30(self) -> None:
        f = self._card("ตัวเลือก — ยื่นต่อใบอนุญาตทำงาน (บต.30 / MoU)")
        self.bt30_frame = f
        self.bt30_excel_input = self._remember(
            "bt30_excel_input", tk.StringVar(value=str(ROOT / "from_bt30.xlsx")))
        self._file_row(f, 0, "ไฟล์ from_bt30.xlsx:", self.bt30_excel_input,
                       hint="คอลัมน์ที่ต้องมี: No., คำนำหน้า, ชื่อ, สัญชาติ, เพศ, วันเกิด (ค.ศ. เช่น 01/01/1995)")
        self.bt30_login_input = self._remember(
            "bt30_login_input", tk.StringVar(value=str(ROOT / "UsernameLogin.xlsx")))
        self._file_row(f, 1, "ไฟล์ UsernameLogin.xlsx:", self.bt30_login_input)
        self.bt30_row_range = tk.StringVar(value="")
        self._text_row(f, 2, "แถวที่จะทำ:", self.bt30_row_range, 'เช่น "1-5" หรือ "2,4" — เว้นว่าง = ทุกแถว')

        self.bt30_do_step2 = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            f, text="ทำขั้นตอนที่ 2 ต่อ (2.1–2.9 → Step 4 เอกสารนายจ้าง → 5.1/5.2 สรุปคำขอ → 6.1 ยืนยันตัวตน)",
            variable=self.bt30_do_step2).grid(row=3, column=0, columnspan=3, sticky="w", pady=(8, 0))
        self._note(f, "1 แถว = 1 คำขอ (เปิดฟอร์มใหม่ทุกคน) หยุดหลัง Step 6.1 — ไม่ส่งคำขอ/ไม่ชำระเงิน  •  "
                      "screenshot เก็บใน reports/bt30_screenshots/", row=4)

        ttk.Separator(f, orient="horizontal").grid(row=5, column=0, columnspan=3, sticky="we", pady=8)
        self.bt30_submit = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            f, text="⚠ ส่งคำขอจริง + ชำระเงิน (Step 7) แล้วเก็บข้อมูล + ดาวน์โหลดใบชำระเงิน (Step 8)",
            variable=self.bt30_submit).grid(row=6, column=0, columnspan=3, sticky="w")
        self._note(f, "อันตราย: กด 'ถัดไป' หน้าชำระเงิน = ยืนยันส่งคำขอจริง ย้อนกลับไม่ได้ และมีค่าธรรมเนียม  •  "
                      "ต้องเปิดขั้นตอนที่ 2 ก่อน  •  ใบชำระเงินดาวน์โหลดได้ครั้งเดียว → reports/bt30_submitted/",
                   style="Danger.TLabel", row=7)
        f.columnconfigure(1, weight=1)

    def _build_card_bt44(self) -> None:
        f = self._card("ตัวเลือก — เปลี่ยนย้ายนายจ้าง (บต.44)")
        self.bt44_frame = f
        self.bt44_excel_input = self._remember(
            "bt44_excel_input", tk.StringVar(value=str(ROOT / "from_bt44.xlxs.xlsx")))
        self._file_row(f, 0, "ไฟล์ from_bt44:", self.bt44_excel_input,
                       hint="คอลัมน์ที่ต้องมี: No., คำนำหน้า, หมายเลขอ้างอิงของคนต่างด้าว (ไม่มีก็ได้), ชื่อ, สัญชาติ, เพศ, เกิดวันที่")
        self.bt44_login_input = self._remember(
            "bt44_login_input", tk.StringVar(value=str(ROOT / "UsernameLogin.xlsx")))
        self._file_row(f, 1, "ไฟล์ UsernameLogin.xlsx:", self.bt44_login_input)
        self.bt44_row_range = tk.StringVar(value="")
        self._text_row(f, 2, "แถวที่จะทำ:", self.bt44_row_range, 'เช่น "1-5" หรือ "2,4" — เว้นว่าง = ทุกแถว')
        self._note(f, "หยุดหลังกด 'บันทึก' (ไม่กดถัดไป/ไม่ส่งคำขอ) — ถ้า Modal แจ้ง Error จะเก็บข้อความไว้ในรายงานแล้วข้ามคนถัดไป",
                   style="Warn.TLabel", row=3, pady=(6, 0))

        self.bt44_dry_run = tk.BooleanVar(value=False)
        ttk.Checkbutton(f, text="🧪 ทดลองยื่น (Dry-Run) — ไม่ส่งคำขอจริง + เก็บข้อมูลคนต่างด้าวเต็มลงรายงาน",
                        variable=self.bt44_dry_run, command=self._on_bt44_dryrun_toggle).grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(8, 0))
        self.bt44_dry_stop = tk.StringVar(value="step2")
        self._bt44_dry_stop_frame = ttk.Frame(f)
        self._bt44_dry_stop_frame.grid(row=5, column=0, columnspan=3, sticky="w", padx=(24, 0), pady=(2, 0))
        ttk.Label(self._bt44_dry_stop_frame, text="หยุดที่:").pack(side="left")
        self._bt44_dry_stop_rb2 = ttk.Radiobutton(
            self._bt44_dry_stop_frame, text="Step 2 (ตรวจเลขใบอนุญาต) — ไม่แก้ที่อยู่",
            variable=self.bt44_dry_stop, value="step2")
        self._bt44_dry_stop_rb2.pack(side="left", padx=(6, 12))
        self._bt44_dry_stop_rb3 = ttk.Radiobutton(
            self._bt44_dry_stop_frame, text="Step 3 (เลือกนายจ้าง+ประเภทกิจการ+งาน) — รวมแก้ที่อยู่ หยุดก่อนแนบเอกสาร",
            variable=self.bt44_dry_stop, value="step3")
        self._bt44_dry_stop_rb3.pack(side="left")
        self._on_bt44_dryrun_toggle()

        self.bt44_check_docs = tk.BooleanVar(value=True)
        cb = ttk.Checkbutton(f, text="📎 ตรวจไฟล์แนบ Step 4 ก่อนรัน (เอกสารไม่ครบ/เปิดไม่ได้ → ข้าม record นั้น)",
                             variable=self.bt44_check_docs)
        cb.grid(row=6, column=0, columnspan=3, sticky="w", pady=(8, 0))
        Tooltip(cb, "ปิดได้เมื่อยังไม่ต้องการตรวจไฟล์ (เช่น ระหว่างทดสอบ flow Dry-Run) — ใช้งานจริงควรเปิดไว้")
        f.columnconfigure(1, weight=1)

    def _build_card_billpay(self) -> None:
        f = self._card("ตัวเลือก — ใบแจ้งชำระเงิน (รอจ่ายค่าธรรมเนียม)")
        self.billpay_frame = f
        self.billpay_login_input = self._remember(
            "billpay_login_input", tk.StringVar(value=str(ROOT / "UsernameLogin.xlsx")))
        self._file_row(f, 0, "ไฟล์ UsernameLogin.xlsx:", self.billpay_login_input,
                       hint="วนทุก Username/Password/Type ในไฟล์")
        self.billpay_row_range = tk.StringVar(value="")
        self._text_row(f, 1, "แถวที่จะดาวน์โหลด:", self.billpay_row_range,
                       'เช่น "1-5" หรือ "2,4" — เว้นว่าง = ทุกแถวในแต่ละบัญชี')
        ttk.Label(f, text="รายการคำขอ:").grid(row=2, column=0, sticky="nw", pady=4)
        box = ttk.Frame(f)
        box.grid(row=2, column=1, columnspan=2, sticky="we", padx=8, pady=4)
        sb = ttk.Scrollbar(box, orient="vertical")
        self.billpay_reqtypes_lb = tk.Listbox(
            box, selectmode="multiple", height=6, exportselection=False, yscrollcommand=sb.set,
            activestyle="none", bg=PALETTE["surface"], relief="flat", highlightthickness=1,
            highlightbackground=PALETTE["border"], selectbackground=PALETTE["accent"],
            selectforeground="#FFFFFF", font=FONT_UI)
        sb.config(command=self.billpay_reqtypes_lb.yview)
        sb.pack(side="right", fill="y")
        self.billpay_reqtypes_lb.pack(side="left", fill="both", expand=True)
        self._billpay_reqtype_codes = [code for code, _ in REQUEST_TYPES if code]
        for code in self._billpay_reqtype_codes:
            lbl = next((l for c, l in REQUEST_TYPES if c == code), code)
            self.billpay_reqtypes_lb.insert("end", lbl)
        ttk.Label(f, text="คลิกเลือกได้หลายรายการ (คลิกซ้ำ = ยกเลิก) — ไม่เลือก = ทุกรายการคำขอ",
                  style="Hint.TLabel").grid(row=3, column=1, columnspan=2, sticky="w", padx=8)
        self._note(f, "สร้าง 'ใบแจ้งชำระเงิน' เพื่อพิมพ์เก็บไว้จ่ายภายหลังผ่านธนาคาร — ไม่มีการตัดเงิน  •  "
                      "ไฟล์: {เลขที่คำขอ}_{รายชื่อคนต่างด้าว} → reports/bill_payment/", row=4, pady=(6, 0))
        f.columnconfigure(1, weight=1)

    def _build_card_payrcpt(self) -> None:
        f = self._card("ตัวเลือก — ใบเสร็จรับเงินทั้งชุด")
        self.payrcpt_frame = f
        self.payrcpt_request_input = self._remember(
            "payrcpt_request_input", tk.StringVar(value=str(ROOT / "Request_Receipt.xlsx")))
        self._file_row(f, 0, "ไฟล์ Request_Receipt.xlsx:", self.payrcpt_request_input,
                       hint="คอลัมน์ที่ต้องมี: ลำดับ, เลขที่คำขอ, Username")
        self.payrcpt_login_input = self._remember(
            "payrcpt_login_input", tk.StringVar(value=str(ROOT / "UsernameLogin.xlsx")))
        self._file_row(f, 1, "ไฟล์ UsernameLogin.xlsx:", self.payrcpt_login_input)
        self.payrcpt_row_range = tk.StringVar(value="")
        self._text_row(f, 2, "แถวที่จะทำ:", self.payrcpt_row_range, 'เช่น "1-5" หรือ "2,4" — เว้นว่าง = ทุกแถว')
        self._note(f, "ไฟล์: {เลขที่คำขอ}_{ชื่อคนต่างด้าว}_{ชื่อนายจ้าง}_{เลขอ้างอิง}_RECEIPT{Amount}.pdf → reports/payment_receipts/",
                   row=3, pady=(6, 0))
        f.columnconfigure(1, weight=1)

    def _build_card_appointment(self) -> None:
        f = self._card("ตัวเลือก — นัดหมายถ่ายบัตร")
        self.appt_frame = f
        self.appt_login_input = self._remember(
            "appt_login_input", tk.StringVar(value=str(ROOT / "UsernameLogin.xlsx")))
        self._file_row(f, 0, "ไฟล์ UsernameLogin.xlsx:", self.appt_login_input)
        self.appt_row_range = tk.StringVar(value="")
        self._text_row(f, 1, "แถวที่จะทำ (ต่อบัญชี):", self.appt_row_range,
                       'เช่น "1-5" หรือ "2,4" — เว้นว่าง = ทุกคำขอ AP ของบัญชีนั้น')
        self._note(f, "รายการคำขอ = ทั้งหมด, สถานะ = รอนัดหมาย (AP) เสมอ  •  PDF → reports/appointment_receipts/  •  "
                      "รายงาน Excel มี hyperlink ไปที่ไฟล์ PDF", row=2, pady=(6, 0))
        f.columnconfigure(1, weight=1)

    def _build_card_bt30ctn(self) -> None:
        f = self._card("ตัวเลือก — ใบตอบรับ / บต.25 / ใบนัดหมาย / ใบเสร็จ")
        self.bt30ctn_frame = f
        self.bt30ctn_login_input = self._remember(
            "bt30ctn_login_input", tk.StringVar(value=str(ROOT / "UsernameLogin.xlsx")))
        self._file_row(f, 0, "ไฟล์ UsernameLogin.xlsx:", self.bt30ctn_login_input, clearable=True,
                       hint="วนทุกบัญชีในไฟล์ — ถ้าเว้นว่าง/ไม่พบไฟล์ จะใช้ Username/Password จากการ์ด 'บัญชีเข้าสู่ระบบ'")
        self.bt30ctn_row_range = tk.StringVar(value="")
        self._text_row(f, 1, "แถวที่จะทำ (ต่อบัญชี):", self.bt30ctn_row_range, 'เช่น "1-10,15" — เว้นว่าง = ทั้งหมด')

        ttk.Label(f, text="เอกสารที่จะดาวน์โหลด:").grid(row=2, column=0, sticky="nw", pady=(6, 2))
        box = ttk.Frame(f)
        box.grid(row=2, column=1, columnspan=2, sticky="w", pady=(6, 2))
        self.bt30ctn_do_ctn = tk.BooleanVar(value=True)
        self.bt30ctn_do_appointment = tk.BooleanVar(value=False)
        self.bt30ctn_do_bt25 = tk.BooleanVar(value=False)
        self.bt30ctn_do_receipt = tk.BooleanVar(value=False)
        for i, (text, var, tip) in enumerate((
            ("ใบตอบรับ บต.30 (CTN) — สถานะ WP2", self.bt30ctn_do_ctn,
             "tab 'เอกสารตอบรับจากระบบ' → {PASSPORT}_BT{XX}_CTN.pdf"),
            ("ใบนัดหมาย (APPOINTMENT) — สถานะ AP/APSS", self.bt30ctn_do_appointment,
             "tab 'การนัดหมาย' → iframe queue → {PASSPORT}_APPOINTMENT.pdf"),
            ("แบบ บต.25 (เอกสารตอบรับจากระบบ)", self.bt30ctn_do_bt25,
             "มีเฉพาะรายการคำขอ MT_59 — ระบบจะเสนอเพิ่ม MT_59 ให้อัตโนมัติ → {PASSPORT}_BT25.pdf"),
            ("Receipt ค่าธรรมเนียมใบอนุญาตทำงาน", self.bt30ctn_do_receipt,
             "tab 'การชำระเงิน' → {PASSPORT}_RECEIPT.pdf"),
        )):
            cb = ttk.Checkbutton(box, text=text, variable=var)
            cb.grid(row=i // 2, column=i % 2, sticky="w", padx=(0, 18), pady=1)
            Tooltip(cb, tip)
        self.bt30ctn_make_subfolder = self._remember("bt30ctn_make_subfolder", tk.BooleanVar(value=False))
        ttk.Checkbutton(f, text="สร้างโฟลเดอร์แยกตาม PASSPORT (bt30_ctn/MH788309/MH788309_BT30_CTN.pdf)",
                        variable=self.bt30ctn_make_subfolder).grid(
            row=3, column=0, columnspan=3, sticky="w", pady=(6, 2))
        self._note(f, "ใช้ 'รายการคำขอ / วันที่ยื่น / สถานะ' จากการ์ดด้านบนร่วมกัน  •  PDF → reports/bt30_ctn/  •  "
                      "รายงาน Excel มี hyperlink ไปที่ไฟล์ PDF", style="Info.TLabel", row=4)
        f.columnconfigure(1, weight=1)

    def _build_card_namelist(self) -> None:
        f = self._card("ตัวเลือก — รายชื่อคนต่างด้าว (NameListAlien)")
        self.namelist_frame = f
        self.namelist_form_type = self._remember(
            "namelist_form_type", tk.StringVar(value="MT_63_2_3103_RENEWAL"))
        ttk.Label(f, text="Form Type:").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Combobox(f, textvariable=self.namelist_form_type, state="normal",
                     values=["MT_63_2_3103_RENEWAL", "MT_63_2_1302_RENEWAL", "MT_63_2_AGN_RENEWAL",
                             "MT_63_2_19_RENEWAL"]).grid(row=0, column=1, sticky="we", padx=8)
        ttk.Label(f, text="พิมพ์ code อื่นเองได้", style="Hint.TLabel").grid(row=1 - 1, column=2, sticky="w")
        ttk.Label(f, text="จำกัดจำนวน:").grid(row=1, column=0, sticky="w", pady=3)
        self.namelist_limit = tk.IntVar(value=0)
        ttk.Spinbox(f, from_=0, to=99999, textvariable=self.namelist_limit, width=10).grid(
            row=1, column=1, sticky="w", padx=8)
        ttk.Label(f, text="0 = ทั้งหมด — ใส่ 100 เพื่อทดสอบเร็ว (100 แถว/หน้า)", style="Hint.TLabel").grid(
            row=1, column=2, sticky="w")
        f.columnconfigure(1, weight=1)

    def _build_card_ticket(self) -> None:
        f = self._card("ตัวเลือก — ตรวจสอบ Ticket ID")
        self.ticket_frame = f
        self.ticket_excel_input = self._remember(
            "ticket_excel_input", tk.StringVar(value=str(ROOT / "TicketID.xlsx")))
        self._file_row(f, 0, "ไฟล์ Excel Ticket ID:", self.ticket_excel_input,
                       hint="คอลัมน์ที่ต้องมี: Ticket ID (หนึ่งแถวต่อหนึ่งเลข)")
        f.columnconfigure(1, weight=1)

    def _build_card_booking(self) -> None:
        f = self._card("ตัวเลือก — ตรวจวันว่างจอง (คิวถ่ายบัตร)")
        self.booking_frame = f
        self.booking_login_input = self._remember("booking_login_input", tk.StringVar(value=""))
        self._file_row(f, 0, "ไฟล์ UsernameLogin.xlsx (ไม่บังคับ):", self.booking_login_input, clearable=True,
                       hint="ถ้าระบุจะใช้บัญชีแรกในไฟล์ — เว้นว่าง = ใช้ Username/Password จากการ์ด 'บัญชีเข้าสู่ระบบ'")
        ttk.Label(f, text="จำนวนเดือนที่ตรวจ:").grid(row=1, column=0, sticky="w", pady=3)
        self.booking_months_ahead = self._remember("booking_months_ahead", tk.IntVar(value=3))
        ttk.Spinbox(f, from_=1, to=12, textvariable=self.booking_months_ahead, width=8).grid(
            row=1, column=1, sticky="w", padx=8)
        ttk.Label(f, text="นับจากเดือนนี้ — แนะนำ 1-3 (ประมาณ 2-3 นาที/เดือน/89 สาขา)", style="Hint.TLabel").grid(
            row=1, column=2, sticky="w")

        ttk.Label(f, text="สาขา (เว้นว่าง = ทุกสาขา):").grid(row=2, column=0, sticky="nw", pady=4)
        self.booking_branch_filter = tk.StringVar(value="")  # compat
        box = ttk.Frame(f)
        box.grid(row=2, column=1, columnspan=2, sticky="we", padx=8, pady=4)
        srow = ttk.Frame(box)
        srow.pack(fill="x")
        ttk.Label(srow, text="🔍").pack(side="left")
        self.booking_search = tk.StringVar(value="")
        ttk.Entry(srow, textvariable=self.booking_search).pack(side="left", fill="x", expand=True, padx=(4, 8))
        self.booking_search.trace_add("write", lambda *_: self._booking_apply_search())
        ttk.Button(srow, text="โหลด/รีเฟรชรายชื่อสาขา", command=self._booking_load_branches).pack(
            side="left", padx=(0, 4))
        ttk.Button(srow, text="ล้างเลือก", command=self._booking_clear_selection).pack(side="left")
        lrow = ttk.Frame(box)
        lrow.pack(fill="both", expand=True, pady=(4, 0))
        sb = ttk.Scrollbar(lrow, orient="vertical")
        self.booking_branches_lb = tk.Listbox(
            lrow, selectmode="multiple", height=8, exportselection=False, yscrollcommand=sb.set,
            activestyle="none", bg=PALETTE["surface"], relief="flat", highlightthickness=1,
            highlightbackground=PALETTE["border"], selectbackground=PALETTE["accent"],
            selectforeground="#FFFFFF", font=FONT_UI)
        sb.config(command=self.booking_branches_lb.yview)
        sb.pack(side="right", fill="y")
        self.booking_branches_lb.pack(side="left", fill="both", expand=True)
        self._booking_all_branches: list[dict] = []
        self._booking_visible_idx: list[int] = []
        self._booking_selected_codes: set[str] = set()
        ttk.Label(f, text="คลิกเลือกได้หลายสาขา (คลิกซ้ำ = ยกเลิก) — ครั้งแรกให้กด 'โหลด/รีเฟรชรายชื่อสาขา' (ต้องกรอกบัญชีก่อน)",
                  style="Hint.TLabel").grid(row=3, column=1, columnspan=2, sticky="w", padx=8)

        mon = ttk.Frame(f)
        mon.grid(row=4, column=0, columnspan=3, sticky="we", pady=(10, 0))
        ttk.Button(mon, text="🖥️  เปิดหน้าจอ Monitor (Real-Time)", style="Accent.TButton",
                   command=self._booking_open_monitor).pack(side="left")
        ttk.Label(mon, text="หน้าต่างใหม่ แสดงที่ว่างสด ๆ + auto-refresh (แนะนำเลือก 5-10 สาขาเพื่อรอบ scan สั้น)",
                  style="Hint.TLabel").pack(side="left", padx=(10, 0))
        f.columnconfigure(1, weight=1)
        self._booking_load_branches_from_cache()

    # ---------- helpers ----------
    def _toggle_pw(self) -> None:
        self.pw_entry.config(show="" if self.show_pw.get() else "•")

    def _on_bt44_dryrun_toggle(self) -> None:
        """เปิด/ปิด radio buttons สำหรับเลือกระดับการหยุดของ dry-run"""
        try:
            state = "normal" if self.bt44_dry_run.get() else "disabled"
            self._bt44_dry_stop_rb2.config(state=state)
            self._bt44_dry_stop_rb3.config(state=state)
        except Exception:
            pass

    # ---------- sidebar ----------
    def _select_nav(self, mode: str) -> None:
        iid = f"mode:{mode}"
        try:
            if self.nav.exists(iid):
                self.nav.selection_set(iid)
                self.nav.focus(iid)
                self.nav.see(iid)
        except Exception:
            pass

    def _on_nav_select(self, _evt=None) -> None:
        sel = self.nav.selection()
        if not sel:
            return
        iid = sel[0]
        if iid.startswith("grp:"):
            # คลิกหัวกลุ่ม → กลับไปเลือกโหมดเดิม (ไม่ให้หัวกลุ่มถูกเลือกค้าง)
            self.after(10, lambda: self._select_nav(self.source_mode.get()))
            return
        mode = iid.split(":", 1)[1]
        if mode in MODE_BY_KEY and mode != self.source_mode.get():
            self.source_mode.set(mode)
            self._on_mode_changed()

    def _on_mode_changed(self) -> None:
        """แสดง/ซ่อนการ์ดตามโหมดที่เลือก + อัปเดตหัวข้อ/ป้ายความปลอดภัย/ไฟล์ผลลัพธ์"""
        mode = self.source_mode.get()
        m = MODE_BY_KEY.get(mode, MODE_BY_KEY["etracking"])
        self._select_nav(mode)

        # หัวข้อ
        self.mode_title.config(text=m["title"])
        badge_text, badge_style = SAFETY_BADGE[m["safety"]]
        self.mode_badge.config(text=badge_text, style=badge_style)
        self.mode_desc.config(text=m["desc"])

        # ซ่อนทุกการ์ดก่อน
        for child in self._body.winfo_children():
            child.pack_forget()

        # ไฟล์ผลลัพธ์เริ่มต้นตามโหมด (เปลี่ยนให้เฉพาะเมื่อผู้ใช้ยังไม่ได้ตั้งชื่อเอง)
        cur = self.out_path.get()
        cur_basename = Path(cur).name if cur else ""
        if not cur_basename or cur_basename in DEFAULT_OUT_NAMES:
            self.out_path.set(str(REPORTS_DIR / m["out"]))

        pack = dict(fill="x", pady=(0, 8))
        mode_frames = {
            "aliens": self.aliens_frame, "register": self.register_frame,
            "receipts": self.receipt_frame, "results": self.result_frame,
            "inform": self.inform_frame, "bt30": self.bt30_frame, "bt44": self.bt44_frame,
            "bill_payment": self.billpay_frame, "payment_receipts": self.payrcpt_frame,
            "appointment": self.appt_frame, "bt30_ctn": self.bt30ctn_frame,
            "namelist_alien": self.namelist_frame, "booking_availability": self.booking_frame,
            "ticket": self.ticket_frame,
        }

        # 1) บัญชีเข้าสู่ระบบ — เฉพาะโหมดที่ใช้
        login_kind = m["login"]
        if login_kind in ("creds", "either"):
            self.login_frame.pack(**pack)
            if mode in ("etracking", "permit_report"):
                hint = "ใช้เมื่อดึงบัญชีเดียว — ถ้าติ๊ก 'ดึงหลายบัญชี' หรือระบุไฟล์ Ref_number ด้านล่าง จะใช้บัญชีจากไฟล์แทน"
            elif login_kind == "either":
                hint = "ใช้เมื่อไม่ระบุไฟล์ UsernameLogin.xlsx ในตัวเลือกด้านล่าง"
            else:
                hint = ""
            self.login_hint.config(text=hint)
            if hint:
                self.login_hint.grid()
            else:
                self.login_hint.grid_remove()

        # 2) การ์ดเฉพาะโหมด / ชุด e-Tracking
        if mode == "etracking":
            self.multi_frame.pack(**pack)
            self.req_frame.pack(**pack)
            self.status_frame.pack(**pack)
            self.etk_template_frame.pack(**pack)
            self.etk_req_no_frame.pack(**pack)
        elif mode == "permit_report":
            self.multi_frame.pack(**pack)
            self.req_frame.pack(**pack)
            self.status_frame.pack(**pack)
        elif mode == "bt30_ctn":
            self.req_frame.pack(**pack)
            self.status_frame.pack(**pack)
            self.bt30ctn_frame.pack(**pack)
        else:
            fr = mode_frames.get(mode)
            if fr is not None:
                fr.pack(**pack)

        # 3) ไฟล์ผลลัพธ์ (ทุกโหมด) — ตัวเลือก limit/relogin เฉพาะที่ backend รองรับ
        self.output_frame.pack(**pack)
        if mode in ("etracking", "permit_report"):
            self.run_opts_row.grid()
        else:
            self.run_opts_row.grid_remove()

        # เลื่อนกลับบนสุด
        try:
            self._scroll_canvas.yview_moveto(0)
        except Exception:
            pass

    # ---------- request listbox ----------
    def _render_request_listbox(self, filt: str = "") -> None:
        """เติมรายการใน request_listbox ตามคำค้น + re-apply selection ที่เก็บไว้ (ข้าม filter)"""
        filt = (filt or "").strip().lower()
        self.request_listbox.delete(0, "end")
        self._req_visible_idx = []
        for i, (code, lbl) in enumerate(zip(self._req_codes, self._req_labels)):
            if filt and filt not in lbl.lower() and filt not in code.lower():
                continue
            self._req_visible_idx.append(i)
            self.request_listbox.insert("end", f"{lbl}   [{code}]")
            if code in self._req_selected_codes:
                self.request_listbox.selection_set("end")

    def _req_persist_selection(self) -> None:
        """เก็บ selection ปัจจุบัน (เฉพาะแถวที่มองเห็น) → self._req_selected_codes ข้าม filter"""
        try:
            sel = self.request_listbox.curselection()
        except Exception:
            return
        codes_now = set()
        for row_idx in sel:
            if row_idx < len(self._req_visible_idx):
                m = self._req_visible_idx[row_idx]
                if 0 <= m < len(self._req_codes):
                    codes_now.add(self._req_codes[m])
        visible_codes = {
            self._req_codes[i] for i in self._req_visible_idx
            if 0 <= i < len(self._req_codes)
        }
        self._req_selected_codes -= visible_codes
        self._req_selected_codes |= codes_now

    def _req_apply_search(self) -> None:
        self._req_persist_selection()
        self._render_request_listbox(self.request_search.get())
        self._update_request_summary()

    def _get_selected_request_codes(self) -> list[str]:
        """คืน list ของ code ที่ผู้ใช้เลือก (รวมที่ถูกซ่อนจาก filter) ตามลำดับรายการต้นฉบับ"""
        self._req_persist_selection()
        return [c for c in self._req_codes if c in self._req_selected_codes]

    def _update_request_summary(self) -> None:
        codes = self._get_selected_request_codes()
        n = len(codes)
        if n == 0:
            self.request_summary.set("ยังไม่เลือก → จะดึงทุกรายการคำขอ")
        elif n == 1:
            self.request_summary.set(f"เลือก 1 รายการ: {codes[0]}")
        else:
            preview = ", ".join(codes[:3])
            more = f" (+{n - 3} อื่น)" if n > 3 else ""
            self.request_summary.set(f"เลือก {n} รายการ: {preview}{more}")

    def _req_select_all(self) -> None:
        try:
            self._req_selected_codes = set(self._req_codes)
            self._render_request_listbox(self.request_search.get())
            self._on_request_changed()
        except Exception:
            pass

    def _req_clear(self) -> None:
        try:
            self._req_selected_codes.clear()
            self._render_request_listbox(self.request_search.get())
            self._on_request_changed()
        except Exception:
            pass

    def _on_request_changed(self, _evt=None) -> None:
        codes = self._get_selected_request_codes()
        first = codes[0] if codes else ""
        self.request_type.set(first)
        self._apply_profile_defaults(first)
        self._update_request_summary()

    def _on_result_request_changed(self, _evt=None) -> None:
        lbl = self.result_request_label.get()
        for code, label in REQUEST_TYPES:
            if label == lbl:
                self.result_request_type.set(code)
                return

    def _on_etk_multi_changed(self) -> None:
        """เปิด/ปิด ช่องเลือกไฟล์ UsernameLogin ตามช็กบ็อกซ์หลายบัญชี"""
        on = self.etk_multi.get()
        state = "normal" if on else "disabled"
        for w in (self.etk_login_entry, self.etk_login_btn, self.etk_combine_cb):
            try:
                w.config(state=state)
            except Exception:
                pass

    def _apply_profile_defaults(self, request_code: str) -> None:
        """อัปเดต checkbox สถานะ + capture_extra_notes ตาม profile ของ request_type"""
        prof = get_profile(request_code)
        active = set(prof.get("filter_status_ids") or ["WA"])
        for code, var in self.filter_vars.items():
            var.set(code in active)
        cap = prof.get("capture_extra_notes", True)
        self.capture_extra_notes.set(bool(cap))

    # ---------- file dialogs ----------
    def _choose_register_input(self) -> None:
        self._choose_into(self.register_input)

    def _choose_into(self, var: tk.StringVar) -> None:
        cur = (var.get() or "").strip()
        initdir = str(Path(cur).parent) if cur and Path(cur).parent.exists() else str(ROOT)
        path = filedialog.askopenfilename(
            initialdir=initdir,
            filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            var.set(path)

    def _choose_out(self) -> None:
        REPORTS_DIR.mkdir(parents=True, exist_ok=True)
        cur = Path(self.out_path.get() or (REPORTS_DIR / "WA_report.xlsx"))
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialdir=str(cur.parent if cur.parent.exists() else REPORTS_DIR),
            initialfile=cur.name,
            filetypes=[("Excel", "*.xlsx")],
        )
        if path:
            self.out_path.set(path)

    # ---- Branches (booking mode) ----
    _BRANCHES_CACHE = ROOT / "_branches_cache.json"

    def _booking_clear_selection(self) -> None:
        self._booking_selected_codes.clear()
        self.booking_branches_lb.selection_clear(0, "end")

    def _booking_persist_selection(self) -> None:
        sel = self.booking_branches_lb.curselection()
        codes_now = set()
        for row_idx in sel:
            if row_idx < len(self._booking_visible_idx):
                master_idx = self._booking_visible_idx[row_idx]
                if 0 <= master_idx < len(self._booking_all_branches):
                    codes_now.add(self._booking_all_branches[master_idx]["code"])
        visible_codes = {
            self._booking_all_branches[i]["code"]
            for i in self._booking_visible_idx
            if 0 <= i < len(self._booking_all_branches)
        }
        self._booking_selected_codes -= visible_codes
        self._booking_selected_codes |= codes_now

    def _booking_render_listbox(self, filt: str = "") -> None:
        filt = (filt or "").strip().lower()
        self.booking_branches_lb.delete(0, "end")
        self._booking_visible_idx = []
        for i, br in enumerate(self._booking_all_branches):
            code = br.get("code", "")
            name = br.get("name", "")
            if filt and filt not in code.lower() and filt not in name.lower():
                continue
            self._booking_visible_idx.append(i)
            self.booking_branches_lb.insert("end", f"{code}  —  {name}")
            if code in self._booking_selected_codes:
                self.booking_branches_lb.selection_set("end")

    def _booking_apply_search(self) -> None:
        self._booking_persist_selection()
        self._booking_render_listbox(self.booking_search.get())

    def _booking_load_branches_from_cache(self) -> None:
        try:
            if self._BRANCHES_CACHE.exists():
                data = json.loads(self._BRANCHES_CACHE.read_text(encoding="utf-8"))
                if isinstance(data, list) and data:
                    self._booking_all_branches = [
                        {"code": str(x.get("code", "")), "name": str(x.get("name", ""))}
                        for x in data if x.get("code")
                    ]
                    self._booking_render_listbox()
        except Exception:
            pass

    def _booking_load_branches(self) -> None:
        """โหลดรายชื่อสาขาสด ๆ ผ่าน login+token (รันใน thread)"""
        if getattr(self, "_worker", None) and self._worker.is_alive():
            messagebox.showinfo("รอสักครู่", "มีงานอื่นทำงานอยู่ — กรุณารอให้เสร็จก่อน")
            return
        self._booking_persist_selection()
        cfg = self._collect_cfg_for_branches()
        if not cfg:
            return
        self._log("[branches] กำลังโหลดรายชื่อสาขา — กรุณารอ 20-30 วิ...")

        def _worker() -> None:
            try:
                from scrape_wa import (
                    login as _login,
                    _booking_get_token,
                    _booking_fetch_branches,
                    _launch_chromium,
                )
                from playwright.sync_api import sync_playwright
                branches: list[dict] = []
                with sync_playwright() as pw:
                    browser = _launch_chromium(
                        pw, cfg, ["--ignore-certificate-errors", "--start-maximized"],
                    )
                    ctx = browser.new_context(
                        locale="th-TH", ignore_https_errors=True,
                        viewport={"width": 1600, "height": 1000},
                    )
                    page = ctx.new_page()
                    try:
                        _login(page, {
                            "username": cfg["username"], "password": cfg["password"],
                            "user_type": cfg.get("user_type", "ผู้กระทำการแทน"),
                            "method": cfg.get("method", "E-Workpermit"),
                        })
                        page.wait_for_timeout(1200)
                        token = _booking_get_token(page, log=self._log)
                        if not token:
                            self._log("[branches] ✗ ไม่พบ token — ต้องมี record สถานะ AP/SS อย่างน้อย 1 คำขอ")
                            return
                        branches = _booking_fetch_branches(page, token, log=self._log)
                    finally:
                        ctx.close(); browser.close()
                if branches:
                    branches_clean = [
                        {"code": b.get("code", ""), "name": b.get("name", "")}
                        for b in branches if b.get("code")
                    ]
                    branches_clean.sort(key=lambda x: (x["code"], x["name"]))
                    try:
                        self._BRANCHES_CACHE.write_text(
                            json.dumps(branches_clean, ensure_ascii=False, indent=2),
                            encoding="utf-8",
                        )
                    except Exception:
                        pass

                    def _update() -> None:
                        self._booking_all_branches = branches_clean
                        self._booking_render_listbox(self.booking_search.get())
                        self._log(f"[branches] ✓ ได้ {len(branches_clean)} สาขา (เซฟ cache แล้ว)")
                    self.after(0, _update)
                else:
                    self._log("[branches] ✗ ได้ 0 สาขา")
            except Exception as e:
                self._log(f"[branches] ✗ error: {e}")

        threading.Thread(target=_worker, daemon=True).start()

    def _collect_cfg_for_branches(self) -> dict | None:
        """รวม cfg (username/password/type/method) จากการ์ดบัญชีเข้าสู่ระบบ"""
        u = (self.username.get() or "").strip()
        p = self.password.get() or ""
        if not u or not p:
            messagebox.showwarning(
                "ข้อมูลไม่ครบ",
                "กรุณากรอก Username/Password ในการ์ด 'บัญชีเข้าสู่ระบบ' ก่อน แล้วค่อยกดโหลดสาขา",
            )
            return None
        return {
            "username": u, "password": p,
            "user_type": self.user_type.get(),
            "method": self.method.get(),
            "headless": bool(self.headless.get()),
            "hide_window": bool(self.hide_window.get()),
        }

    # ---------- log / progress ----------
    def _open_folder(self) -> None:
        p = Path(self.out_path.get()).parent
        p.mkdir(parents=True, exist_ok=True)
        if p.exists():
            import os
            os.startfile(str(p))  # type: ignore[attr-defined]

    def _log(self, msg: str) -> None:
        self._log_q.put(str(msg))

    def _clear_log(self) -> None:
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

    @staticmethod
    def _log_tag(msg: str) -> str | None:
        s = msg.lstrip()
        if s.startswith(("[ผิดพลาด]", "Traceback", "⛔")) or "✗" in s[:12] or "[!]" in s[:6]:
            return "err"
        if "✓" in s[:12] or s.startswith(("[✔]", "เสร็จ")):
            return "ok"
        if s.startswith(("[⏸", "[▶", "[i]", "💡", "⚠")):
            return "warn"
        return None

    def _drain_log(self) -> None:
        try:
            while True:
                item = self._log_q.get_nowait()
                if isinstance(item, tuple) and item and item[0] == "progress":
                    _, done, total = item
                    self._last_progress = (done, total)
                    self.progress["maximum"] = max(total, 1)
                    self.progress["value"] = done
                    self.progress_lbl.config(text=f"{done} / {total}" if total else "กำลังทำงาน…")
                    continue
                msg = str(item)
                self.log_text.config(state="normal")
                tag = self._log_tag(msg)
                if tag:
                    self.log_text.insert("end", msg + "\n", tag)
                else:
                    self.log_text.insert("end", msg + "\n")
                self.log_text.see("end")
                self.log_text.config(state="disabled")
        except queue.Empty:
            pass
        self.after(120, self._drain_log)

    def _set_progress(self, done: int, total: int) -> None:
        # เรียกได้จาก worker thread — ส่งผ่าน queue ให้ main thread อัปเดต
        self._log_q.put(("progress", int(done), int(total)))

    def _set_running(self, running: bool) -> None:
        if running:
            self.btn_start.config(state="disabled")
            self.btn_cancel.config(state="normal")
            self.btn_pause.config(state="normal")
            self.btn_resume.config(state="disabled")
            self.progress_lbl.config(text="กำลังเริ่ม…")
        else:
            self.btn_start.config(state="normal")
            self.btn_cancel.config(state="disabled")
            self.btn_pause.config(state="disabled")
            self.btn_resume.config(state="disabled")
        self._pause_evt.set()

    # ---------- actions ----------
    def _relogin_every_value(self) -> int:
        """แปลงตัวเลือก 'Login ใหม่ทุก' (dropdown หรือพิมพ์เอง) → จำนวนรายการ (0 = ปิด)"""
        raw = (self.relogin_every.get() or "").strip().replace(",", "")
        try:
            n = int(raw)
        except ValueError:
            return 0
        return n if n > 0 else 0

    def _warn(self, title: str, msg: str) -> None:
        messagebox.showwarning(title, msg)
        return None

    def _on_start(self) -> None:
        if self._worker is not None and self._worker.is_alive():
            return
        # ตรวจสอบค่าทั้งหมด "ก่อน" เปลี่ยนสถานะปุ่ม — ถ้าค่าผิด ปุ่มจะไม่ค้าง
        try:
            job = self._collect_job()
        except (tk.TclError, ValueError) as e:
            messagebox.showwarning("ค่าไม่ถูกต้อง", f"กรุณาตรวจตัวเลขในช่องตัวเลือก\n\n{e}")
            return
        if job is None:
            return
        self._clear_log()
        self._cancel.clear()
        self._set_running(True)
        self._set_progress(0, 0)
        self._save_settings()
        self._last_progress = (0, 0)
        try:
            self._job_meta = usage_tracker.begin_job(
                mode=job[0], operator=self.operator.get().strip(),
                ewp_username=job[1].get("username", ""),
                login_file=getattr(self, "_job_login_file", None), out_path=job[2],
                app_version=APP_VERSION,
            )
        except Exception:
            self._job_meta = None
        self._worker = threading.Thread(target=self._run, args=job, daemon=True)
        self._worker.start()

    def _collect_job(self) -> tuple | None:
        """อ่าน+ตรวจค่าจากหน้าจอ → คืน args สำหรับ _run หรือ None (แจ้งเตือนแล้ว)"""
        mode = self.source_mode.get()
        m = MODE_BY_KEY[mode]
        etk_multi = (mode in ("etracking", "permit_report") and self.etk_multi.get())
        etk_combine = (mode == "etracking" and etk_multi and self.etk_combine.get())
        etk_req_excel_str = self.etk_req_excel_input.get().strip() if mode == "etracking" else ""

        # ---- ต้องมี Username/Password ในหน้าจอหรือไม่ ----
        needs_creds = m["login"] == "creds"
        if etk_multi or (mode == "etracking" and etk_req_excel_str):
            needs_creds = False
        if mode == "booking_availability" and not (self.booking_login_input.get() or "").strip():
            needs_creds = True
        if needs_creds and (not self.username.get().strip() or not self.password.get()):
            return self._warn("ข้อมูลไม่ครบ", "กรุณากรอก Username และ Password ในการ์ด 'บัญชีเข้าสู่ระบบ'")

        cfg = {
            "username": self.username.get().strip(),
            "password": self.password.get(),
            "user_type": self.user_type.get(),
            "method": self.method.get(),
            "headless": self.headless.get(),
            "hide_window": self.hide_window.get(),
            "request_type": self.request_type.get().strip(),
            "request_types": self._get_selected_request_codes(),
            "date_from": self.date_from.get().strip(),
            "date_to": self.date_to.get().strip(),
            "filter_status_ids": [c for c, v in self.filter_vars.items() if v.get()],
            "status_whitelist": [s for s, v in self.whitelist_vars.items() if v.get()],
            "capture_extra_notes": self.capture_extra_notes.get(),
            "relogin_every": self._relogin_every_value(),
            "req_no_list": (
                self.etk_req_no_text.get("1.0", "end").strip()
                if mode == "etracking" else ""
            ),
            "report_template": (
                self.etk_template.get() if mode == "etracking" else "main"
            ),
        }
        out_str = (self.out_path.get() or "").strip()
        if not out_str:
            return self._warn("ข้อมูลไม่ครบ", "กรุณาระบุไฟล์ผลลัพธ์ (.xlsx)")
        out = Path(out_str)
        limit = int(self.limit.get() or 0) if mode in ("etracking", "permit_report") else 0
        sub_tabs = [n for n, v in self.alien_subtab_vars.items() if v.get()]
        if mode == "aliens" and not sub_tabs:
            return self._warn("ข้อมูลไม่ครบ", "กรุณาเลือกอย่างน้อย 1 ตารางในโหมดข้อมูลคนต่างด้าว")

        etk_req_excel = Path(etk_req_excel_str) if etk_req_excel_str else None
        etk_req_login = Path(self.etk_req_login_input.get()) if mode == "etracking" else None
        if mode == "etracking" and etk_req_excel is not None:
            if not etk_req_excel.exists():
                return self._warn("ไม่พบไฟล์ Ref_number", f"ไม่พบไฟล์: {etk_req_excel}")
            if not etk_req_login or not etk_req_login.exists():
                return self._warn("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ UsernameLogin: {etk_req_login}")

        register_input = Path(self.register_input.get()) if mode == "register" else None
        register_row_range = self.register_row_range.get().strip() or None
        if mode == "register":
            if not register_input or not register_input.exists():
                return self._warn("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ Excel: {register_input}")

        receipt_request = Path(self.receipt_request_input.get()) if mode == "receipts" else None
        receipt_login = Path(self.receipt_login_input.get()) if mode == "receipts" else None
        receipt_row_range = self.receipt_row_range.get().strip() or None
        receipt_name_suffix = self.receipt_name_suffix.get().strip() if mode == "receipts" else ""
        receipt_make_folder = bool(self.receipt_make_folder.get()) if mode == "receipts" else False
        receipt_doc_types: list[str] | None = None
        if mode == "receipts":
            if not receipt_request or not receipt_request.exists():
                return self._warn("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ RequestData: {receipt_request}")
            if not receipt_login or not receipt_login.exists():
                return self._warn("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ UsernameLogin: {receipt_login}")
            receipt_doc_types = [
                key for key, var in (
                    ("receipt", self.doc_receipt), ("bt44", self.doc_bt44),
                    ("bt22", self.doc_bt22), ("bt53", self.doc_bt53),
                    ("bt55", self.doc_bt55), ("bt52", self.doc_bt52),
                    ("bt56", self.doc_bt56),
                ) if var.get()
            ]
            if not receipt_doc_types:
                return self._warn("ข้อมูลไม่ครบ", "กรุณาเลือกประเภทเอกสารอย่างน้อย 1 อย่าง")

        result_login = Path(self.result_login_input.get()) if mode == "results" else None
        result_request_type = self.result_request_type.get().strip()
        result_ref_str = self.result_ref_input.get().strip() if mode == "results" else ""
        result_ref = Path(result_ref_str) if result_ref_str else None
        result_doc_keys: list[str] | None = None
        if mode == "results":
            if not result_login or not result_login.exists():
                return self._warn("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ UsernameLogin: {result_login}")
            if result_ref is not None and not result_ref.exists():
                return self._warn(
                    "ไม่พบไฟล์ Ref_number",
                    f"ระบุไฟล์ Ref_number แต่หาไม่พบ:\n{result_ref}\n\n"
                    "ถ้าต้องการใช้โหมด B (e-Tracking) ให้กดปุ่ม 'ล้าง' ที่ช่อง Ref_number",
                )
            result_doc_keys = [
                key for key, var in (
                    ("result_notice", self.doc_result_notice),
                    ("request_receipt", self.doc_request_receipt),
                    ("bt50", self.doc_result_bt50),
                ) if var.get()
            ]
            if not result_doc_keys:
                return self._warn("ข้อมูลไม่ครบ", "กรุณาเลือกเอกสารอย่างน้อย 1 อย่าง (ใบแจ้งผล / ใบรับคำขอ / บต.50)")
            cfg["request_type"] = result_request_type

        inform_excel = Path(self.inform_excel_input.get()) if mode == "inform" else None
        inform_login = Path(self.inform_login_input.get()) if mode == "inform" else None
        inform_row_range = self.inform_row_range.get().strip() or None
        inform_commit = bool(self.inform_commit.get()) if mode == "inform" else False
        if mode == "inform":
            if not inform_excel or not inform_excel.exists():
                return self._warn("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ FormRequestEmployment: {inform_excel}")
            if not inform_login or not inform_login.exists():
                return self._warn("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ UsernameLogin: {inform_login}")
            if inform_commit:
                if not messagebox.askyesno(
                    "ยืนยันส่งคำขอจริง",
                    "โหมด 'COMMIT' จะบันทึกข้อมูลจริงลงระบบ (ข้อ 7+8) — ยืนยันดำเนินการ?",
                    icon="warning",
                ):
                    return None

        bt30_excel = Path(self.bt30_excel_input.get()) if mode == "bt30" else None
        bt30_login = Path(self.bt30_login_input.get()) if mode == "bt30" else None
        bt30_row_range = self.bt30_row_range.get().strip() or None
        bt30_do_step2 = bool(self.bt30_do_step2.get()) if mode == "bt30" else False
        bt30_submit = bool(self.bt30_submit.get()) if mode == "bt30" else False
        if mode == "bt30":
            if not bt30_excel or not bt30_excel.exists():
                return self._warn("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ from_bt30: {bt30_excel}")
            if not bt30_login or not bt30_login.exists():
                return self._warn("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ UsernameLogin: {bt30_login}")
            if bt30_do_step2:
                try:
                    _recs = _read_bt30_excel(bt30_excel)
                    _sel = [_recs[i - 1] for i in _parse_row_range(bt30_row_range, len(_recs))]
                    _oversize = _bt30_preflight_doc_sizes(_sel, log=lambda *a: None)
                except Exception as e:
                    messagebox.showerror("อ่านไฟล์ไม่สำเร็จ", f"ตรวจขนาดไฟล์แนบก่อนรันไม่สำเร็จ:\n{e}")
                    return None
                if _oversize:
                    _lines = "\n".join(
                        f"• แถว {p['row']} {p['name']} | {p['label']}\n     {p['file']} = {p['mb']} MB"
                        for p in _oversize
                    )
                    messagebox.showerror(
                        f"ไฟล์แนบเกิน {_BT30_MAX_DOC_MB:.0f} MB — แก้ไขก่อนจึงจะรันได้",
                        f"พบไฟล์แนบเกิน {_BT30_MAX_DOC_MB:.0f} MB จำนวน {len(_oversize)} ไฟล์\n"
                        f"ระบบ e-WorkPermit จำกัดไฟล์แนบไม่เกิน {_BT30_MAX_DOC_MB:.0f} MB ต่อไฟล์\n\n"
                        f"{_lines}\n\n"
                        f"กรุณาย่อ/บีบอัดไฟล์ให้ ≤ {_BT30_MAX_DOC_MB:.0f} MB แล้วกดรันใหม่",
                    )
                    self._log(f"⛔ ยกเลิกการรัน — พบไฟล์แนบเกิน {_BT30_MAX_DOC_MB:.0f} MB "
                              f"จำนวน {len(_oversize)} ไฟล์ (ต้องแก้ไขก่อน)")
                    return None
            if bt30_submit and not bt30_do_step2:
                return self._warn("ตั้งค่าไม่ถูกต้อง", "ต้องเปิด 'ทำขั้นตอนที่ 2 ต่อ' ก่อน จึงจะส่งคำขอจริง (Step 7-8) ได้")
            if bt30_submit:
                if not messagebox.askyesno(
                    "⚠ ยืนยันส่งคำขอจริง — ย้อนกลับไม่ได้",
                    "เปิดโหมด 'ส่งคำขอจริง + ชำระเงิน' (Step 7-8)\n\n"
                    "การกด 'ถัดไป' หน้าชำระเงินคือการยืนยันส่งคำขอจริงต่อกรมการจัดหางาน "
                    "ย้อนกลับไม่ได้ และจะเกิดค่าธรรมเนียมจริง\n\n"
                    "ระบบจะเก็บข้อมูลหน้าผลสำเร็จและดาวน์โหลดใบแจ้งชำระเงิน "
                    "(ดาวน์โหลดได้ครั้งเดียว) ลงในโฟลเดอร์ reports/bt30_submitted/\n\n"
                    "ยืนยันดำเนินการส่งคำขอจริงหรือไม่?",
                    icon="warning",
                ):
                    return None

        etk_login = Path(self.etk_login_input.get()) if etk_multi else None
        if etk_multi:
            if not etk_login or not etk_login.exists():
                return self._warn("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ UsernameLogin: {etk_login}")

        bt44_excel = Path(self.bt44_excel_input.get()) if mode == "bt44" else None
        bt44_login = Path(self.bt44_login_input.get()) if mode == "bt44" else None
        bt44_row_range = self.bt44_row_range.get().strip() or None
        bt44_dry_run = bool(self.bt44_dry_run.get()) if mode == "bt44" else False
        bt44_dry_stop = self.bt44_dry_stop.get() if mode == "bt44" else "step2"
        bt44_check_docs = bool(self.bt44_check_docs.get()) if mode == "bt44" else True
        if mode == "bt44":
            if not bt44_excel or not bt44_excel.exists():
                return self._warn("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ from_bt44: {bt44_excel}")
            if not bt44_login or not bt44_login.exists():
                return self._warn("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ UsernameLogin: {bt44_login}")

        billpay_login = Path(self.billpay_login_input.get()) if mode == "bill_payment" else None
        billpay_row_range = self.billpay_row_range.get().strip() or None
        billpay_request_types: list[str] | None = None
        if mode == "bill_payment":
            if not billpay_login or not billpay_login.exists():
                return self._warn("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ UsernameLogin: {billpay_login}")
            sel_idx = self.billpay_reqtypes_lb.curselection()
            billpay_request_types = [self._billpay_reqtype_codes[i] for i in sel_idx] or None

        payrcpt_request = Path(self.payrcpt_request_input.get()) if mode == "payment_receipts" else None
        payrcpt_login = Path(self.payrcpt_login_input.get()) if mode == "payment_receipts" else None
        payrcpt_row_range = self.payrcpt_row_range.get().strip() or None
        if mode == "payment_receipts":
            if not payrcpt_request or not payrcpt_request.exists():
                return self._warn("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ Request_Receipt: {payrcpt_request}")
            if not payrcpt_login or not payrcpt_login.exists():
                return self._warn("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ UsernameLogin: {payrcpt_login}")

        appt_login = Path(self.appt_login_input.get()) if mode == "appointment" else None
        appt_row_range = self.appt_row_range.get().strip() or None
        if mode == "appointment":
            if not appt_login or not appt_login.exists():
                return self._warn("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ UsernameLogin: {appt_login}")
            cfg["request_type"] = ""  # โหมดนี้ใช้ 'ทั้งหมด' เสมอ

        bt30ctn_login_str = self.bt30ctn_login_input.get().strip() if mode == "bt30_ctn" else ""
        bt30ctn_login = Path(bt30ctn_login_str) if bt30ctn_login_str else None
        bt30ctn_row_range = self.bt30ctn_row_range.get().strip() or None
        bt30ctn_subfolder = bool(self.bt30ctn_make_subfolder.get()) if mode == "bt30_ctn" else False
        bt30ctn_do_ctn = bool(self.bt30ctn_do_ctn.get()) if mode == "bt30_ctn" else True
        bt30ctn_do_appt = bool(self.bt30ctn_do_appointment.get()) if mode == "bt30_ctn" else False
        bt30ctn_do_bt25 = bool(self.bt30ctn_do_bt25.get()) if mode == "bt30_ctn" else False
        bt30ctn_do_receipt = bool(self.bt30ctn_do_receipt.get()) if mode == "bt30_ctn" else False
        if mode == "bt30_ctn":
            if not (bt30ctn_do_ctn or bt30ctn_do_appt or bt30ctn_do_bt25 or bt30ctn_do_receipt):
                return self._warn("ข้อมูลไม่ครบ", "กรุณาเลือกอย่างน้อย 1 เอกสาร (บต.30 / บต.25 / ใบนัดหมาย / ใบเสร็จ)")
            if bt30ctn_do_bt25 and "MT_59" not in cfg.get("request_types", []):
                _mt59_label = ("การยื่นคำขอรับใบอนุญาตทำงานของคนต่างด้าวซึ่งได้รับอนุญาต"
                               "ให้อยู่ในราชอาณาจักรเป็นการชั่วคราว")
                _ok = messagebox.askyesno(
                    "บต.25 ต้องเลือกรายการคำขอ MT_59",
                    "การดาวน์โหลด 'แบบ บต.25' ต้องกรอง 'รายการคำขอ' เป็น:\n"
                    f"  • {_mt59_label} (MT_59)\n\n"
                    "แต่ตอนนี้ยังไม่ได้เลือกไว้\n\n"
                    "กด 'Yes' เพื่อเพิ่ม MT_59 ให้อัตโนมัติแล้วรันต่อ\n"
                    "กด 'No' เพื่อยกเลิก แล้วกลับไปเลือกเอง",
                )
                if not _ok:
                    return None
                self._req_selected_codes.add("MT_59")
                self._render_request_listbox(self.request_search.get())
                self._update_request_summary()
                cfg["request_types"] = self._get_selected_request_codes()
                self._log("[i] เพิ่มรายการคำขอ MT_59 อัตโนมัติ (จำเป็นสำหรับ บต.25)")
            _has_file = bool(bt30ctn_login and bt30ctn_login.exists())
            _has_creds = bool(self.username.get().strip() and self.password.get())
            if bt30ctn_login_str and not _has_file and not _has_creds:
                return self._warn("ไม่พบไฟล์", f"ไม่พบไฟล์ UsernameLogin: {bt30ctn_login}\n"
                                  "กด 'ล้าง' ที่ช่องไฟล์เพื่อใช้ Username/Password แทน")
            if not _has_file and not _has_creds:
                return self._warn(
                    "ข้อมูลไม่ครบ",
                    "โหมดนี้ต้องมีอย่างใดอย่างหนึ่ง:\n"
                    "  1) ไฟล์ UsernameLogin.xlsx ที่มีข้อมูลบัญชี\n"
                    "  2) กรอก Username/Password ในการ์ด 'บัญชีเข้าสู่ระบบ'",
                )
            if not _has_file:
                bt30ctn_login = None

        namelist_form_type = self.namelist_form_type.get().strip() if mode == "namelist_alien" else ""
        namelist_limit = int(self.namelist_limit.get() or 0) if mode == "namelist_alien" else 0
        if mode == "namelist_alien" and not namelist_form_type:
            return self._warn("ข้อมูลไม่ครบ", "กรุณาระบุ Form Type (เช่น MT_63_2_3103_RENEWAL)")

        booking_login: Path | None = None
        booking_months_ahead = 3
        booking_branch_filter: list[str] | None = None
        if mode == "booking_availability":
            _bl = (self.booking_login_input.get() or "").strip()
            if _bl:
                p = Path(_bl)
                if not p.exists():
                    return self._warn("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ UsernameLogin: {p}")
                booking_login = p
            booking_months_ahead = max(1, int(self.booking_months_ahead.get() or 3))
            self._booking_persist_selection()
            booking_branch_filter = sorted(self._booking_selected_codes) or None

        ticket_excel: Path | None = None
        if mode == "ticket":
            _te = (self.ticket_excel_input.get() or "").strip()
            if not _te:
                return self._warn("ข้อมูลไม่ครบ", "กรุณาเลือกไฟล์ Excel รายการ Ticket ID")
            ticket_excel = Path(_te)
            if not ticket_excel.exists():
                return self._warn("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ '{ticket_excel}'")

        # ไฟล์บัญชีที่งานนี้ใช้ (สำหรับบันทึกการใช้งานว่าทำให้บัญชีไหนบ้าง)
        self._job_login_file = {
            "etracking": (etk_req_login if etk_req_excel is not None else etk_login),
            "permit_report": etk_login, "receipts": receipt_login, "results": result_login,
            "inform": inform_login, "bt30": bt30_login, "bt44": bt44_login,
            "bill_payment": billpay_login, "payment_receipts": payrcpt_login,
            "appointment": appt_login, "bt30_ctn": bt30ctn_login,
            "booking_availability": booking_login,
        }.get(mode)
        return (mode, cfg, out, limit, sub_tabs, register_input, register_row_range,
                receipt_request, receipt_login, receipt_row_range, receipt_doc_types,
                receipt_name_suffix, receipt_make_folder,
                result_login, result_doc_keys, result_ref, etk_multi, etk_login,
                etk_combine, etk_req_excel, etk_req_login,
                inform_excel, inform_login, inform_row_range, inform_commit,
                bt30_excel, bt30_login, bt30_row_range, bt30_do_step2, bt30_submit,
                bt44_excel, bt44_login, bt44_row_range, bt44_dry_run, bt44_dry_stop, bt44_check_docs,
                billpay_login, billpay_row_range, billpay_request_types,
                payrcpt_request, payrcpt_login, payrcpt_row_range,
                appt_login, appt_row_range,
                bt30ctn_login, bt30ctn_row_range, bt30ctn_subfolder,
                bt30ctn_do_ctn, bt30ctn_do_appt,
                bt30ctn_do_bt25, bt30ctn_do_receipt,
                namelist_form_type, namelist_limit,
                booking_login, booking_months_ahead, booking_branch_filter,
                ticket_excel)

    def _on_cancel(self) -> None:
        if messagebox.askyesno("ยืนยัน", "ต้องการยกเลิกการทำงานใช่หรือไม่?"):
            self._cancel.set()
            self._pause_evt.set()
            self.btn_pause.config(state="disabled")
            self.btn_resume.config(state="disabled")
            self._log("[!] ส่งสัญญาณยกเลิก... จะหยุดหลังเสร็จรายการปัจจุบัน")

    def _on_pause(self) -> None:
        """พัก process ค้างไว้กับที่ (ไม่ปิดเบราว์เซอร์) — worker จะหยุดที่รายการถัดไป"""
        self._pause_evt.clear()
        self.btn_pause.config(state="disabled")
        self.btn_resume.config(state="normal")
        self._log("[⏸] หยุดชั่วคราว — จะค้างไว้หลังจบรายการปัจจุบัน กด 'ต่อ' เพื่อทำงานต่อ")

    def _on_resume(self) -> None:
        self._pause_evt.set()
        self.btn_resume.config(state="disabled")
        self.btn_pause.config(state="normal")
        self._log("[▶] ทำงานต่อจากเดิม")

    def _wait_if_paused_or_cancelled(self) -> bool:
        """callback ส่งให้ run_* แทน is_cancelled: พัก → บล็อกจนกด 'ต่อ'/'ยกเลิก'; คืน True เมื่อยกเลิก"""
        self._pause_evt.wait()
        return self._cancel.is_set()
    def _run(
        self, mode: str, cfg: dict, out: Path, limit: int, sub_tabs: list[str],
        register_input: Path | None = None, register_row_range: str | None = None,
        receipt_request: Path | None = None, receipt_login: Path | None = None,
        receipt_row_range: str | None = None, receipt_doc_types: list[str] | None = None,
        receipt_name_suffix: str = "", receipt_make_folder: bool = False,
        result_login: Path | None = None, result_doc_keys: list[str] | None = None,
        result_ref: Path | None = None,
        etk_multi: bool = False, etk_login: Path | None = None,
        etk_combine: bool = False,
        etk_req_excel: Path | None = None, etk_req_login: Path | None = None,
        inform_excel: Path | None = None, inform_login: Path | None = None,
        inform_row_range: str | None = None, inform_commit: bool = False,
        bt30_excel: Path | None = None, bt30_login: Path | None = None,
        bt30_row_range: str | None = None, bt30_do_step2: bool = True,
        bt30_submit: bool = False,
        bt44_excel: Path | None = None, bt44_login: Path | None = None,
        bt44_row_range: str | None = None, bt44_dry_run: bool = False,
        bt44_dry_stop: str = "step2",
        bt44_check_docs: bool = True,
        billpay_login: Path | None = None, billpay_row_range: str | None = None,
        billpay_request_types: list[str] | None = None,
        payrcpt_request: Path | None = None, payrcpt_login: Path | None = None,
        payrcpt_row_range: str | None = None,
        appt_login: Path | None = None, appt_row_range: str | None = None,
        bt30ctn_login: Path | None = None, bt30ctn_row_range: str | None = None,
        bt30ctn_subfolder: bool = False,
        bt30ctn_do_ctn: bool = True, bt30ctn_do_appt: bool = False,
        bt30ctn_do_bt25: bool = False, bt30ctn_do_receipt: bool = False,
        namelist_form_type: str = "", namelist_limit: int = 0,
        booking_login: Path | None = None, booking_months_ahead: int = 3,
        booking_branch_filter: list[str] | None = None,
        ticket_excel: Path | None = None,
    ) -> None:
        try:
            if mode == "aliens":
                count, path = run_aliens_scrape(
                    cfg, out, sub_tabs=sub_tabs,
                    log=self._log,
                    progress=self._set_progress,
                    is_cancelled=self._wait_if_paused_or_cancelled,
                )
            elif mode == "register":
                count, path = run_register(
                    cfg, register_input, out,
                    row_range=register_row_range,
                    log=self._log,
                    progress=self._set_progress,
                    is_cancelled=self._wait_if_paused_or_cancelled,
                )
            elif mode == "receipts":
                count, path = run_receipts(
                    cfg, receipt_request, receipt_login, out,
                    row_range=receipt_row_range,
                    doc_types=receipt_doc_types,
                    name_suffix=receipt_name_suffix,
                    make_subfolder=receipt_make_folder,
                    log=self._log,
                    progress=self._set_progress,
                    is_cancelled=self._wait_if_paused_or_cancelled,
                )
            elif mode == "results":
                if result_ref is not None and result_ref.exists():
                    self._log(f"[i] โหมด results: ค้นหาตามเลขคำขอจาก {result_ref.name}")
                    count, path = run_result_docs_by_ref(
                        cfg, result_ref, result_login, out,
                        doc_keys=result_doc_keys,
                        log=self._log,
                        progress=self._set_progress,
                        is_cancelled=self._wait_if_paused_or_cancelled,
                    )
                else:
                    self._log("[i] โหมด results: ดึงจาก e-Tracking (ไม่ได้ระบุไฟล์ Ref_number)")
                    count, path = run_result_docs(
                        cfg, result_login, out,
                        request_type=cfg.get("request_type", ""),
                        doc_keys=result_doc_keys,
                        log=self._log,
                        progress=self._set_progress,
                        is_cancelled=self._wait_if_paused_or_cancelled,
                    )
            elif mode == "inform":
                count, path = run_inform_employer(
                    cfg, inform_excel, inform_login, out,
                    row_range=inform_row_range, commit=inform_commit,
                    log=self._log,
                    progress=self._set_progress,
                    is_cancelled=self._wait_if_paused_or_cancelled,
                )
            elif mode == "bt30":
                count, path = run_bt30(
                    cfg, bt30_excel, bt30_login, out,
                    row_range=bt30_row_range, do_step2=bt30_do_step2,
                    do_submit=bt30_submit,
                    log=self._log,
                    progress=self._set_progress,
                    is_cancelled=self._wait_if_paused_or_cancelled,
                )
            elif mode == "bt44":
                count, path = run_bt44(
                    cfg, bt44_excel, bt44_login, out,
                    row_range=bt44_row_range,
                    dry_run=bt44_dry_run,
                    dry_stop_at=bt44_dry_stop,
                    check_docs=bt44_check_docs,
                    log=self._log,
                    progress=self._set_progress,
                    is_cancelled=self._wait_if_paused_or_cancelled,
                )
            elif mode == "bill_payment":
                count, path = run_bill_payment(
                    cfg, billpay_login, out,
                    row_range=billpay_row_range,
                    request_types=billpay_request_types,
                    log=self._log,
                    progress=self._set_progress,
                    is_cancelled=self._wait_if_paused_or_cancelled,
                )
            elif mode == "payment_receipts":
                count, path = run_payment_receipts(
                    cfg, payrcpt_request, payrcpt_login, out,
                    row_range=payrcpt_row_range,
                    log=self._log,
                    progress=self._set_progress,
                    is_cancelled=self._wait_if_paused_or_cancelled,
                )
            elif mode == "appointment":
                count, path = run_appointment(
                    cfg, appt_login, out,
                    request_type=cfg.get("request_type", ""),
                    row_range=appt_row_range,
                    log=self._log,
                    progress=self._set_progress,
                    is_cancelled=self._wait_if_paused_or_cancelled,
                )
            elif mode == "bt30_ctn":
                count, path = run_bt30_ctn(
                    cfg, bt30ctn_login, out,
                    row_range=bt30ctn_row_range,
                    make_subfolder=bt30ctn_subfolder,
                    do_ctn=bt30ctn_do_ctn,
                    do_appointment=bt30ctn_do_appt,
                    do_bt25=bt30ctn_do_bt25,
                    do_receipt=bt30ctn_do_receipt,
                    log=self._log,
                    progress=self._set_progress,
                    is_cancelled=self._wait_if_paused_or_cancelled,
                )
            elif mode == "namelist_alien":
                count, path = run_namelist_alien(
                    cfg, out,
                    form_type=namelist_form_type,
                    limit=namelist_limit,
                    log=self._log,
                    progress=self._set_progress,
                    is_cancelled=self._wait_if_paused_or_cancelled,
                )
            elif mode == "booking_availability":
                count, path = run_booking_availability(
                    cfg, booking_login, out,
                    months_ahead=booking_months_ahead,
                    branch_filter=booking_branch_filter,
                    log=self._log,
                    progress=self._set_progress,
                    is_cancelled=self._wait_if_paused_or_cancelled,
                )
            elif mode == "permit_report":
                if etk_multi and etk_login is not None:
                    count, path = run_permit_report_multi(
                        cfg, etk_login, out, limit=limit,
                        log=self._log,
                        progress=self._set_progress,
                        is_cancelled=self._wait_if_paused_or_cancelled,
                    )
                else:
                    count, path = run_permit_report(
                        cfg, out, limit=limit,
                        log=self._log,
                        progress=self._set_progress,
                        is_cancelled=self._wait_if_paused_or_cancelled,
                    )
            elif mode == "ticket":
                count, path = run_ticket_report(
                    cfg, ticket_excel, out,
                    log=self._log,
                    progress=self._set_progress,
                    is_cancelled=self._wait_if_paused_or_cancelled,
                )
            else:
                if etk_req_excel is not None and etk_req_login is not None:
                    count, path = run_scrape_by_ref(
                        cfg, etk_req_excel, etk_req_login, out,
                        log=self._log,
                        progress=self._set_progress,
                        is_cancelled=self._wait_if_paused_or_cancelled,
                    )
                elif etk_multi and etk_login is not None:
                    count, path = run_scrape_multi(
                        cfg, etk_login, out, limit=limit,
                        combine=etk_combine,
                        log=self._log,
                        progress=self._set_progress,
                        is_cancelled=self._wait_if_paused_or_cancelled,
                    )
                else:
                    count, path = run_scrape(
                        cfg, out, limit=limit,
                        log=self._log,
                        progress=self._set_progress,
                        is_cancelled=self._wait_if_paused_or_cancelled,
                    )
            self.after(0, lambda: self._done_ok(count, path))
        except Exception as e:
            # PEP 3110: `e` ถูกลบทิ้งหลังออกจาก except block → capture เป็น local ก่อน
            import traceback
            err_msg = str(e) or e.__class__.__name__
            tb_txt = traceback.format_exc()
            self._log(f"[ผิดพลาด] {err_msg}")
            self._log(tb_txt)
            self.after(0, lambda em=err_msg: self._done_err(em))


    def _usage_end(self, status: str, count: int = 0, path=None, error: str = "") -> None:
        meta, self._job_meta = self._job_meta, None
        try:
            usage_tracker.end_job(meta, status=status, count=count, out_path=path,
                                  progress=self._last_progress, error=error, log=self._log)
        except Exception:
            pass

    def _done_ok(self, count: int, path: Path) -> None:
        self._set_running(False)
        self._usage_end("cancelled" if self._cancel.is_set() else "ok", count=count, path=path)
        self.progress_lbl.config(text=f"เสร็จ {count} รายการ")
        self._log(f"✓ เสร็จสิ้น — {count} รายการ → {path}")
        if messagebox.askyesno(
            "เสร็จสิ้น",
            f"ดึงข้อมูล {count} รายการ\nบันทึกที่:\n{path}\n\nต้องการเปิดไฟล์เลยไหม?",
        ):
            import os
            try:
                os.startfile(str(path))  # type: ignore[attr-defined]
            except Exception as e:
                messagebox.showerror("เปิดไฟล์ไม่ได้", str(e))

    def _done_err(self, msg: str) -> None:
        self._set_running(False)
        self._usage_end("error", error=msg)
        self.progress_lbl.config(text="ผิดพลาด")
        self._log_box.set_open(True)
        messagebox.showerror("เกิดข้อผิดพลาด", msg)

    def _booking_open_monitor(self) -> None:
        cfg = self._collect_cfg_for_branches()
        if not cfg:
            return
        # เก็บ selection ก่อน
        self._booking_persist_selection()
        selected_codes = sorted(self._booking_selected_codes)
        if not selected_codes:
            if not messagebox.askyesno(
                "ยืนยัน",
                "คุณยังไม่ได้เลือกสาขา — จะ scan ทั้ง 89 สาขา (~5-8 นาทีต่อรอบ)\nต้องการเปิด Monitor ต่อไปหรือไม่?",
            ):
                return
        # แปลง code → {code, name} จาก master list
        branches: list[dict]
        if selected_codes:
            code_set = set(selected_codes)
            branches = [
                {"code": b["code"], "name": b["name"]}
                for b in self._booking_all_branches if b["code"] in code_set
            ]
        else:
            branches = [
                {"code": b["code"], "name": b["name"]}
                for b in self._booking_all_branches if b.get("code")
            ]
        if not branches:
            messagebox.showerror("ไม่มีสาขา", "รายชื่อสาขาว่าง — กรุณาโหลดรายชื่อก่อน")
            return
        try:
            months_ahead = max(1, int(self.booking_months_ahead.get() or 3))
        except Exception:
            months_ahead = 3
        BookingMonitor(self, cfg, branches, months_ahead)


class BookingMonitor(tk.Toplevel):
    """หน้าจอ Real-Time Monitor แสดงที่ว่างสด ๆ + auto-refresh"""

    def __init__(self, parent: tk.Misc, cfg: dict,
                 branches: list[dict], months_ahead: int):
        super().__init__(parent)
        self._cfg = cfg
        self._branches = branches
        self._months_ahead = months_ahead
        self._stop_evt = threading.Event()
        self._msg_q: queue.Queue = queue.Queue()
        self._last_slots: dict[tuple, dict] = {}  # (code,date,round) → slot dict
        self._last_scan_ts: float = 0.0
        self._next_scan_ts: float = 0.0
        self._worker: threading.Thread | None = None

        self.title(
            f"🖥️ Monitor: คิวถ่ายบัตร ({len(branches)} สาขา × {months_ahead} เดือน)"
        )
        self.geometry("1400x750")
        # Booking browser (persistent Chrome — เปิด/ปิดผ่านปุ่มใน toolbar)
        self._booking_browser = BookingBrowser(log_cb=self._log_via_status)
        self._build_ui()
        # เริ่ม auto-poll queue
        self.after(200, self._poll_queue)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        # เริ่ม scan อัตโนมัติ
        self._start_worker()

    def _log_via_status(self, msg: str) -> None:
        """สำหรับให้ BookingBrowser ยิง log ไปที่ status label"""
        self._msg_q.put({"kind": "status", "text": msg})

    def _build_ui(self) -> None:
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text="ทุก").pack(side="left")
        self._interval = tk.IntVar(value=120)  # 2 นาที
        ttk.Spinbox(
            top, from_=30, to=3600, increment=30,
            textvariable=self._interval, width=6,
        ).pack(side="left", padx=4)
        ttk.Label(top, text="วินาที").pack(side="left")

        ttk.Label(top, text=" | ").pack(side="left", padx=6)
        ttk.Label(top, text="เดือนที่ตรวจ:").pack(side="left")
        self._months_var = tk.IntVar(value=self._months_ahead)
        ttk.Spinbox(
            top, from_=1, to=18, increment=1,
            textvariable=self._months_var, width=4,
            command=self._on_months_changed,
        ).pack(side="left", padx=4)
        ttk.Label(top, text="เดือน (ปรับได้ทันที)", foreground="gray").pack(side="left")

        self._btn_stop = ttk.Button(top, text="⏹  หยุด", command=self._stop_worker)
        self._btn_stop.pack(side="left", padx=(10, 4))
        self._btn_start = ttk.Button(
            top, text="▶  เริ่มใหม่", command=self._start_worker, state="disabled",
        )
        self._btn_start.pack(side="left", padx=4)
        ttk.Button(top, text="🔄  scan เดี๋ยวนี้", command=self._trigger_now).pack(
            side="left", padx=4,
        )

        # แสดงทุกรอบ (รวมเต็ม)
        self._include_full = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            top, text="แสดงทุกรอบ (รวมที่เต็มแล้ว)",
            variable=self._include_full,
            command=self._render_tree,
        ).pack(side="left", padx=(12, 4))

        # จัดลำดับ: ว่างขึ้นก่อน (priority sort)
        self._priority_sort = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            top, text="🟢 ว่าง/ใหม่ ขึ้นก่อน",
            variable=self._priority_sort,
            command=self._render_tree,
        ).pack(side="left", padx=(4, 4))

        # ปุ่มเปิด Chrome สำหรับจอง
        ttk.Label(top, text=" | ").pack(side="left", padx=4)
        self._btn_open_chrome = ttk.Button(
            top, text="🌐  เปิด Chrome จอง (login 1 ครั้ง)",
            command=self._toggle_booking_browser,
        )
        self._btn_open_chrome.pack(side="left", padx=(4, 0))
        ttk.Button(
            top, text="📊  ประวัติจอง",
            command=self._open_history,
        ).pack(side="left", padx=(4, 0))
        ttk.Button(
            top, text="⏹  ยกเลิก queue จอง",
            command=self._cancel_booking_queue,
        ).pack(side="left", padx=(4, 0))

        ttk.Label(top, text=" | ").pack(side="left", padx=6)
        ttk.Label(top, text="ค้นหา:").pack(side="left")
        self._search = tk.StringVar(value="")
        _e = ttk.Entry(top, textvariable=self._search, width=25)
        _e.pack(side="left", padx=4)
        self._search.trace_add("write", lambda *_: self._render_tree())

        # status
        self._status = tk.StringVar(value="พร้อม — จะเริ่ม scan รอบแรกทันที")
        ttk.Label(self, textvariable=self._status, padding=6, foreground="#333").pack(
            fill="x",
        )

        # Summary bar
        self._summary = tk.StringVar(value="")
        ttk.Label(
            self, textvariable=self._summary, padding=(8, 0),
            foreground="#1a5490", font=("TkDefaultFont", 10, "bold"),
        ).pack(fill="x")

        # Treeview
        mid = ttk.Frame(self, padding=(6, 4))
        mid.pack(fill="both", expand=True)
        cols = ("date", "code", "name", "round", "left", "cap", "pct", "state")
        self._tree = ttk.Treeview(
            mid, columns=cols, show="headings", height=24,
        )
        headings = [
            ("date", "วันที่", 90),
            ("code", "รหัสสาขา", 130),
            ("name", "ชื่อสาขา", 420),
            ("round", "ช่วงเวลา", 110),
            ("left", "ที่ว่าง", 70),
            ("cap", "รวม", 60),
            ("pct", "% เหลือ", 80),
            ("state", "สถานะ", 100),
        ]
        for key, text, width in headings:
            self._tree.heading(
                key, text=text,
                command=lambda k=key: self._sort_by(k),
            )
            anchor = "center" if key in ("left", "cap", "pct", "state", "round") else "w"
            self._tree.column(key, width=width, anchor=anchor)
        vsb = ttk.Scrollbar(mid, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._tree.pack(side="left", fill="both", expand=True)
        # ตกแต่งสี
        self._tree.tag_configure("new", background="#C6EFCE")     # เขียว = ใหม่
        self._tree.tag_configure("gone", background="#F5B7B1")    # แดง = หายแล้ว
        self._tree.tag_configure("full", background="#FFC7CE")    # แดงเข้ม = เต็ม
        self._tree.tag_configure("low", background="#FFEB9C")     # เหลือง = เหลือน้อย
        self._tree.tag_configure("ok", background="#EAF4EA")      # เขียวอ่อน = ปกติ

        # เก็บ mapping tree_iid → slot dict (สำหรับ double-click)
        self._row_data: dict[str, dict] = {}

        # Double-click / Enter → เปิด popup จอง
        self._tree.bind("<Double-1>", self._on_row_activate)
        self._tree.bind("<Return>", self._on_row_activate)

        # Right-click context menu
        self._ctx_menu = tk.Menu(self, tearoff=0)
        self._ctx_menu.add_command(
            label="🎯 เตรียมจอง (แสดงข้อมูล + copy)",
            command=self._on_row_activate,
        )
        self._ctx_menu.add_command(
            label="📋 คัดลอกทั้งแถว",
            command=self._copy_row_short,
        )
        self._tree.bind("<Button-3>", self._on_right_click)

        self._sort_key = "date"
        self._sort_reverse = False
        self._snapshot: list[dict] = []
        self._prev_keys: set[tuple] = set()
        # ป้ายเวลาถัดไป
        self.after(1000, self._tick_countdown)

    def _sort_by(self, key: str) -> None:
        if self._sort_key == key:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_key = key
            self._sort_reverse = False
        self._render_tree()

    def _tick_countdown(self) -> None:
        if not self._stop_evt.is_set() and self._next_scan_ts > 0:
            import time as _t
            remain = int(self._next_scan_ts - _t.time())
            if remain > 0:
                self._status.set(
                    self._status.get().split("|")[0].strip()
                    + f"  |  รอบถัดไปใน {remain} วิ"
                )
        self.after(1000, self._tick_countdown)

    def _poll_queue(self) -> None:
        try:
            while True:
                msg = self._msg_q.get_nowait()
                self._handle_msg(msg)
        except queue.Empty:
            pass
        self.after(300, self._poll_queue)

    def _handle_msg(self, msg: dict) -> None:
        kind = msg.get("kind")
        if kind == "status":
            self._status.set(msg.get("text", ""))
        elif kind == "scan_start":
            import time as _t
            self._status.set(f"⏳ กำลัง scan… (เริ่ม {_t.strftime('%H:%M:%S')})")
        elif kind == "scan_progress":
            done = msg.get("done", 0)
            total = msg.get("total", 0)
            code = msg.get("code", "")
            month = msg.get("month", "")
            if code:
                self._status.set(
                    f"⏳ scan {done}/{total}  →  {code} ({month})"
                )
            else:
                self._status.set(f"⏳ กำลังรวบรวมผล… ({done}/{total})")
        elif kind == "scan_done":
            import time as _t
            slots = msg.get("slots") or []
            self._last_scan_ts = _t.time()
            self._next_scan_ts = self._last_scan_ts + max(30, int(self._interval.get()))
            # เก็บเฉพาะ "ว่าง" ใน keys (สำหรับ new/gone comparison)
            avail_slots = [s for s in slots if int(s.get("left", 0)) > 0]
            keys_now = {(s["branch_code"], s["date"], s["round"]) for s in avail_slots}
            self._new_keys = keys_now - self._prev_keys
            self._gone_keys = self._prev_keys - keys_now
            self._prev_keys = keys_now
            self._snapshot = slots
            total = len(slots)
            avail = len(avail_slots)
            full = total - avail
            branches_hit = len({s["branch_code"] for s in avail_slots})
            dates_hit = len({s["date"] for s in avail_slots})
            self._summary.set(
                f"🟢 {avail:,} ว่าง  |  🔴 {full:,} เต็ม  |  "
                f"{branches_hit} สาขามีที่ว่าง  |  {dates_hit} วันมีที่ว่าง  |  "
                f"🆕 ใหม่ {len(self._new_keys)}  |  ❌ หาย {len(self._gone_keys)}"
            )
            self._status.set(f"✓ อัปเดตล่าสุด {_t.strftime('%H:%M:%S')}")
            self._render_tree()
        elif kind == "scan_err":
            self._status.set(f"✗ error: {msg.get('text', '')[:200]}")
        elif kind == "worker_done":
            self._btn_stop.config(state="disabled")
            self._btn_start.config(state="normal")
            self._status.set("⏹ หยุดแล้ว")

    def _render_tree(self) -> None:
        # เคลียร์
        for iid in self._tree.get_children():
            self._tree.delete(iid)
        self._row_data.clear()
        show_full = bool(self._include_full.get()) if hasattr(self, "_include_full") else True
        # รวม gone rows (สร้าง virtual entries)
        rows: list[tuple[dict, str]] = []
        for s in self._snapshot:
            key = (s["branch_code"], s["date"], s["round"])
            left = int(s.get("left", 0))
            cap = int(s.get("capacity", 0))
            if key in getattr(self, "_new_keys", set()):
                tag = "new"
            elif left <= 0:
                if not show_full:
                    continue
                tag = "full"
            elif cap > 0 and left <= max(1, cap * 0.2):
                tag = "low"
            else:
                tag = "ok"
            rows.append((s, tag))
        # gone (แสดงเป็นแดง — ยกเว้นถ้าปิด show_full)
        if show_full:
            for key in getattr(self, "_gone_keys", set()):
                code, date, rnd = key
                name = next(
                    (b["name"] for b in self._branches if b["code"] == code), ""
                )
                rows.append((
                    {
                        "branch_code": code, "branch_name": name,
                        "date": date, "round": rnd, "left": 0, "capacity": 0,
                    }, "gone",
                ))
        # filter search
        filt = self._search.get().strip().lower()
        if filt:
            rows = [
                (s, t) for (s, t) in rows
                if filt in s["branch_code"].lower()
                or filt in s["branch_name"].lower()
                or filt in s["date"].lower()
                or filt in s["round"].lower()
            ]
        # sort
        sk = self._sort_key
        rev = self._sort_reverse
        # priority ของสถานะ (ยิ่งเลขน้อย = ขึ้นบน)
        _state_priority = {"new": 0, "ok": 1, "low": 2, "full": 3, "gone": 4}
        priority_first = bool(self._priority_sort.get()) if hasattr(self, "_priority_sort") else True

        def _keyf(item):
            s = item[0]
            tag = item[1]
            # ถ้าเปิด priority sort → ให้ state priority มาก่อน
            prio = _state_priority.get(tag, 9) if priority_first else 0
            # แล้วค่อย sort ตาม column
            if sk == "date": sub = (s.get("date", ""),)
            elif sk == "code": sub = (s.get("branch_code", ""),)
            elif sk == "name": sub = (s.get("branch_name", ""),)
            elif sk == "round": sub = (s.get("round", ""),)
            elif sk == "left": sub = (-int(s.get("left", 0)),)
            elif sk == "cap": sub = (-int(s.get("capacity", 0)),)
            elif sk == "pct":
                cap = int(s.get("capacity", 0))
                sub = (-((int(s.get("left", 0)) / cap * 100) if cap else 0),)
            elif sk == "state":
                sub = (_state_priority.get(tag, 9),)
            else:
                sub = (s.get("date", ""),)
            return (prio,) + sub

        rows.sort(key=_keyf, reverse=rev)
        # insert
        for s, tag in rows:
            cap = int(s.get("capacity", 0))
            left = int(s.get("left", 0))
            pct = round(left / cap * 100, 1) if cap else 0
            state = {"new": "🆕 ใหม่", "gone": "❌ หายแล้ว",
                     "low": "⚠ เหลือน้อย", "ok": "ว่าง",
                     "full": "🔴 เต็ม"}.get(tag, "")
            iid = self._tree.insert(
                "", "end", values=(
                    s.get("date", ""),
                    s.get("branch_code", ""),
                    s.get("branch_name", "") or "",
                    s.get("round", ""),
                    left, cap, f"{pct}%", state,
                ),
                tags=(tag,),
            )
            self._row_data[iid] = {**s, "_tag": tag}

    def _trigger_now(self) -> None:
        # ตั้ง next_scan = 0 เพื่อบังคับ scan ทันที
        self._next_scan_ts = 0
        self._status.set("⚡ บังคับ scan ทันที…")

    def _on_months_changed(self) -> None:
        """ปรับ months_ahead กลางคัน — จะมีผลใน scan รอบถัดไป"""
        try:
            m = max(1, int(self._months_var.get()))
            self._months_ahead = m
            self._status.set(f"✓ ตั้งเดือนที่ตรวจ = {m} เดือน (จะมีผลใน scan รอบถัดไป)")
            # อัปเดต title
            self.title(
                f"🖥️ Monitor: คิวถ่ายบัตร ({len(self._branches)} สาขา × {m} เดือน)"
            )
        except Exception:
            pass

    # ---- Row actions (Copy/Book) ----
    def _get_selected_slot(self) -> dict | None:
        sel = self._tree.selection()
        if not sel:
            return None
        return self._row_data.get(sel[0])

    def _on_right_click(self, event) -> None:
        iid = self._tree.identify_row(event.y)
        if iid:
            self._tree.selection_set(iid)
            try:
                self._ctx_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self._ctx_menu.grab_release()

    def _on_row_activate(self, event=None) -> None:
        slot = self._get_selected_slot()
        if not slot:
            return
        BookingActionDialog(self, slot, self._booking_browser, log_cb=self._log_via_status)

    def _toggle_booking_browser(self) -> None:
        """เปิด/ปิด persistent Chrome สำหรับจอง"""
        if self._booking_browser.is_running:
            if messagebox.askyesno(
                "ปิด Chrome จอง?", "จะปิด Chrome สำหรับจอง? (login/cookies จะยังถูกเก็บไว้)",
                parent=self,
            ):
                self._booking_browser.stop()
                self._btn_open_chrome.config(text="🌐  เปิด Chrome จอง (login 1 ครั้ง)")
                self._status.set("ปิด Chrome จองแล้ว")
        else:
            self._status.set("⏳ กำลังเปิด Chrome สำหรับจอง...")
            self._booking_browser.start()
            self._btn_open_chrome.config(text="🛑  ปิด Chrome จอง")
            self._status.set("✓ Chrome เปิดแล้ว — login บัญชีจองใน Chrome นั้น (ครั้งแรก)")

    def _open_history(self) -> None:
        """เปิดหน้าต่างประวัติจอง"""
        BookingHistoryDialog(self)

    def _cancel_booking_queue(self) -> None:
        """ยกเลิก queue จองที่ค้างอยู่ (ไม่ปิด browser)"""
        if not self._booking_browser.is_running:
            messagebox.showinfo("ไม่มี queue", "Chrome จอง ยังไม่ได้เปิด", parent=self)
            return
        if messagebox.askyesno(
            "ยกเลิก queue?",
            "จะข้ามคำสั่งจองที่ค้างอยู่ทั้งหมด?\n(Chrome ยังเปิดอยู่)",
            parent=self,
        ):
            self._booking_browser.cancel_queue()
            self._status.set("⏹ ยกเลิก queue จอง — Chrome ยังเปิดอยู่")

    def _copy_row_short(self) -> None:
        slot = self._get_selected_slot()
        if not slot:
            return
        left = int(slot.get("left", 0))
        cap = int(slot.get("capacity", 0))
        text = (
            f"{slot.get('branch_code','')} {slot.get('branch_name','')} | "
            f"{slot.get('date','')} {slot.get('round','')} | ว่าง {left}/{cap}"
        )
        try:
            self.clipboard_clear()
            self.clipboard_append(text)
            self._status.set(f"📋 คัดลอกแล้ว: {text[:80]}")
        except Exception as e:
            self._status.set(f"❌ copy ไม่ได้: {e}")

    def _start_worker(self) -> None:
        if self._worker and self._worker.is_alive():
            return
        self._stop_evt.clear()
        self._btn_stop.config(state="normal")
        self._btn_start.config(state="disabled")
        self._worker = threading.Thread(target=self._run_worker, daemon=True)
        self._worker.start()

    def _stop_worker(self) -> None:
        self._stop_evt.set()
        self._btn_stop.config(state="disabled")
        self._status.set("⏹ กำลังหยุด… รอ scan รอบปัจจุบันเสร็จ")

    def _on_close(self) -> None:
        self._stop_evt.set()
        try:
            self._booking_browser.stop()
        except Exception:
            pass
        self.after(500, self.destroy)

    def _run_worker(self) -> None:
        """Background: keep browser open + scan periodically.
        - รองรับ session timeout: ถ้า scan fail 3 ครั้งต่อกัน → re-login + ดึง token ใหม่
        """
        import time as _t
        try:
            from scrape_wa import (
                login as _login,
                _booking_get_token,
                _launch_chromium,
            )
            from playwright.sync_api import sync_playwright
            from datetime import date as _date

            self._msg_q.put({"kind": "status", "text": "🔐 กำลัง login..."})
            with sync_playwright() as pw:
                browser = _launch_chromium(
                    pw, self._cfg,
                    ["--ignore-certificate-errors", "--start-maximized"],
                )
                ctx = browser.new_context(
                    locale="th-TH", ignore_https_errors=True,
                    viewport={"width": 1600, "height": 1000},
                )
                page = ctx.new_page()

                def _do_login_and_token() -> str | None:
                    """Login + get token — return token or None on failure"""
                    try:
                        _login(page, {
                            "username": self._cfg["username"],
                            "password": self._cfg["password"],
                            "user_type": self._cfg.get("user_type", "ผู้กระทำการแทน"),
                            "method": self._cfg.get("method", "E-Workpermit"),
                        })
                        page.wait_for_timeout(1200)
                        return _booking_get_token(page, log=lambda *a: None)
                    except Exception as e:
                        self._msg_q.put({"kind": "status",
                                         "text": f"❌ login/token err: {str(e)[:100]}"})
                        return None

                try:
                    token = _do_login_and_token()
                    if not token:
                        self._msg_q.put({
                            "kind": "scan_err",
                            "text": "ไม่พบ Bearer token (ต้องมี record สถานะ AP/SS)",
                        })
                        return

                    # เตรียม year_month_pairs
                    today = _date.today()

                    def _make_pairs(n: int) -> list[tuple[int, int]]:
                        pairs: list[tuple[int, int]] = []
                        y, m = today.year, today.month
                        for _ in range(max(1, int(n))):
                            pairs.append((y, m))
                            m += 1
                            if m > 12:
                                m = 1; y += 1
                        return pairs

                    # loop
                    consecutive_failures = 0
                    while not self._stop_evt.is_set():
                        # อ่านค่า months ล่าสุด (ผู้ใช้อาจปรับกลางคัน)
                        year_month_pairs = _make_pairs(self._months_ahead)
                        self._msg_q.put({"kind": "scan_start"})

                        def _prog(done, total, code, name, month_str):
                            self._msg_q.put({
                                "kind": "scan_progress",
                                "done": done, "total": total,
                                "code": code, "name": name, "month": month_str,
                            })

                        try:
                            # scan รวมทุกรอบเสมอ — filter ตอนแสดงผลใน UI
                            slots = booking_scan_once(
                                page, token, self._branches, year_month_pairs,
                                log=lambda *a: None,
                                is_cancelled=lambda: self._stop_evt.is_set(),
                                include_full=True,
                                on_progress=_prog,
                            )
                            # ตรวจ session timeout: ถ้าไม่มี slot เลย + scan ก่อนหน้ามี = อาจ session หมด
                            if not slots and consecutive_failures >= 1:
                                raise RuntimeError("empty result — session may have expired")
                            self._msg_q.put({"kind": "scan_done", "slots": slots})
                            consecutive_failures = 0  # reset
                        except Exception as e:
                            consecutive_failures += 1
                            self._msg_q.put({
                                "kind": "scan_err",
                                "text": f"scan #{consecutive_failures} fail: {str(e)[:120]}",
                            })
                            # 2 ครั้งต่อกัน → ลอง re-login
                            if consecutive_failures >= 2:
                                self._msg_q.put({
                                    "kind": "status",
                                    "text": "🔄 session อาจหมดอายุ — กำลัง re-login...",
                                })
                                new_token = _do_login_and_token()
                                if new_token:
                                    token = new_token
                                    self._msg_q.put({
                                        "kind": "status",
                                        "text": "✓ re-login สำเร็จ — scan รอบใหม่",
                                    })
                                    consecutive_failures = 0
                                else:
                                    self._msg_q.put({
                                        "kind": "status",
                                        "text": "❌ re-login ล้มเหลว — รอ scan รอบถัดไป",
                                    })
                        # รอ interval
                        interval = max(30, int(self._interval.get()))
                        for _ in range(interval):
                            if self._stop_evt.is_set():
                                break
                            if self._next_scan_ts == 0:  # trigger now
                                break
                            _t.sleep(1)
                finally:
                    try: ctx.close()
                    except Exception: pass
                    try: browser.close()
                    except Exception: pass
        except Exception as e:
            self._msg_q.put({"kind": "scan_err", "text": f"worker crashed: {e}"})
        finally:
            self._msg_q.put({"kind": "worker_done"})


class BookingBrowser:
    """Persistent Chrome (Playwright) สำหรับจอง — login ครั้งเดียว ใช้ตลอด"""

    TRACKING_URL = "https://eworkpermit.doe.go.th/Permit/Tracking"
    PROFILE_DIR = ROOT / "booking_profile"
    HISTORY_FILE = ROOT / "_booking_history.jsonl"
    SLIPS_DIR = ROOT / "reports" / "booking_slips"

    def __init__(self, log_cb=print):
        self._log = log_cb
        self._cmd_q: queue.Queue = queue.Queue()
        self._thread: threading.Thread | None = None
        self._running = False
        self._ready_evt = threading.Event()

    @property
    def is_running(self) -> bool:
        return self._running and self._thread is not None and self._thread.is_alive()

    def start(self) -> None:
        if self.is_running:
            self._log("[booking-browser] running อยู่แล้ว")
            return
        self._running = True
        self._ready_evt.clear()
        self._queue_stats = {"done": 0, "total": 0, "success": 0, "failed": 0}
        self._cancel_queue = False
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._running = False
        self._cancel_queue = True
        self._cmd_q.put({"kind": "close"})

    def cancel_queue(self) -> None:
        """ยกเลิก queue จองที่ค้างอยู่ (ไม่ปิด browser)"""
        self._cancel_queue = True
        # เคลียร์ queue (drain แต่ leave close command)
        drained = 0
        try:
            while True:
                cmd = self._cmd_q.get_nowait()
                if cmd.get("kind") == "close":
                    self._cmd_q.put(cmd)
                    break
                drained += 1
        except queue.Empty:
            pass
        self._log(f"[booking-browser] ⏹ ยกเลิก queue — เคลียร์ {drained} คำสั่งค้างอยู่")

    def goto_record(self, req_no: str, slot: dict, auto_confirm: bool = False) -> None:
        """ให้ browser ไปที่ record + open การนัดหมาย tab + pre-select slot
        req_no: 1 เลข หรือหลายเลขคั่นด้วย , หรือ newline
        auto_confirm=True → กด 'ยืนยัน' ให้ด้วย (Full-Auto)
        """
        if not self.is_running:
            self._log("[booking-browser] ยังไม่เปิด — กด '🌐 เปิด Chrome จอง' ก่อน")
            return
        # split เป็น list ของ req_no
        req_list = [
            r.strip() for r in req_no.replace(",", "\n").replace(";", "\n").split("\n")
            if r.strip()
        ]
        if not req_list:
            return
        self._cancel_queue = False
        # reset stats ถ้าไม่มี queue ค้าง
        if self._cmd_q.empty() or self._queue_stats["done"] >= self._queue_stats["total"]:
            self._queue_stats = {"done": 0, "total": len(req_list),
                                 "success": 0, "failed": 0}
        else:
            self._queue_stats["total"] += len(req_list)
        for r in req_list:
            self._cmd_q.put({
                "kind": "goto_record", "req_no": r, "slot": slot,
                "auto_confirm": auto_confirm,
            })

    def _log_history(self, entry: dict) -> None:
        """บันทึกประวัติจองลง _booking_history.jsonl"""
        import time as _t
        entry = {"ts": _t.strftime("%Y-%m-%d %H:%M:%S"), **entry}
        try:
            with self.HISTORY_FILE.open("a", encoding="utf-8") as f:
                f.write(json.dumps(entry, ensure_ascii=False) + "\n")
        except Exception:
            pass

    def _run(self) -> None:
        try:
            from playwright.sync_api import sync_playwright
            self.PROFILE_DIR.mkdir(parents=True, exist_ok=True)
            with sync_playwright() as pw:
                ctx = pw.chromium.launch_persistent_context(
                    user_data_dir=str(self.PROFILE_DIR),
                    headless=False,
                    args=["--start-maximized", "--disable-blink-features=AutomationControlled"],
                    viewport=None,
                )
                page = ctx.pages[0] if ctx.pages else ctx.new_page()
                # ไปที่ tracking (ถ้ายังไม่ login จะ redirect ไปหน้า login)
                try:
                    page.goto(self.TRACKING_URL, wait_until="domcontentloaded", timeout=30_000)
                except Exception as e:
                    self._log(f"[booking-browser] goto tracking err: {e}")
                self._log("[booking-browser] ✓ เปิด Chrome แล้ว — กรุณา login ถ้ายังไม่ได้ login")
                self._ready_evt.set()

                while self._running:
                    try:
                        cmd = self._cmd_q.get(timeout=1.0)
                    except queue.Empty:
                        # ตรวจว่า ctx ยังเปิดอยู่ไหม
                        try:
                            _ = len(ctx.pages)
                        except Exception:
                            self._log("[booking-browser] Chrome ถูกปิดแล้ว")
                            break
                        continue

                    kind = cmd.get("kind")
                    if kind == "close":
                        break
                    if kind == "goto_record":
                        if self._cancel_queue:
                            self._log(f"[booking-browser] ⏭ ข้าม {cmd.get('req_no','')} (cancelled)")
                            self._queue_stats["done"] += 1
                            self._queue_stats["failed"] += 1
                            continue
                        try:
                            ok = self._do_goto_record(
                                page, ctx, cmd["req_no"], cmd["slot"],
                                auto_confirm=bool(cmd.get("auto_confirm", False)),
                            )
                            self._queue_stats["done"] += 1
                            if ok: self._queue_stats["success"] += 1
                            else:  self._queue_stats["failed"] += 1
                        except Exception as e:
                            self._log(f"[booking-browser] goto_record error: {e}")
                            self._queue_stats["done"] += 1
                            self._queue_stats["failed"] += 1
                        # แสดง queue progress
                        s = self._queue_stats
                        if s["total"] > 1:
                            self._log(
                                f"[booking-browser] 📊 queue {s['done']}/{s['total']} "
                                f"(✅ {s['success']} ❌ {s['failed']})"
                            )
                        # ถ้าเสร็จ queue → notify
                        if s["done"] >= s["total"] and s["total"] > 1:
                            self._log(
                                f"[booking-browser] ✓ เสร็จ queue: {s['success']}/{s['total']} สำเร็จ, "
                                f"{s['failed']} ล้มเหลว"
                            )
                try:
                    ctx.close()
                except Exception:
                    pass
        except Exception as e:
            self._log(f"[booking-browser] crashed: {e}")
        finally:
            self._running = False

    def _do_goto_record(self, page, ctx, req_no: str, slot: dict,
                        auto_confirm: bool = False) -> bool:
        """Flow: search req_no → click record → tab การนัดหมาย → pre-select slot
        return True = สำเร็จ, False = ล้มเหลว (จะยังต่อ req_no ถัดไปเสมอ)
        """
        self._log(f"[booking-browser] → เปิด record {req_no}")
        history_entry: dict = {
            "req_no": req_no,
            "branch_code": slot.get("branch_code", ""),
            "branch_name": slot.get("branch_name", ""),
            "date": slot.get("date", ""),
            "round": slot.get("round", ""),
            "capacity": int(slot.get("capacity", 0)),
            "left": int(slot.get("left", 0)),
            "auto_confirm": auto_confirm,
            "result": "pending",
            "error": "",
        }
        # เก็บ iframe tab reference ให้ปิดได้ end-of-flow
        qp_tab = None
        success = False
        # 1. Focus tracking page (ถ้าอยู่ที่อื่น → กลับไป Tracking)
        try:
            cur = page.url or ""
            if "/Permit/Tracking" not in cur:
                page.goto(self.TRACKING_URL, wait_until="domcontentloaded", timeout=30_000)
                page.wait_for_timeout(1500)
        except Exception:
            pass
        # ตรวจ login
        if "login" in (page.url or "").lower() or page.locator('input[type="password"]').count():
            self._log("[booking-browser] ⚠ ยังไม่ได้ login — กรุณา login ในหน้า Chrome แล้วสั่ง 🚀 อีกครั้ง")
            history_entry["result"] = "failed"
            history_entry["error"] = "not logged in"
            self._log_history(history_entry)
            return False
        # 2. พิมพ์ req_no ในช่องค้นหา
        try:
            inp = page.locator(
                'input[placeholder*="ค้นหา"], input[type="search"], input.form-control'
            ).first
            inp.click(timeout=8000)
            inp.fill(req_no, timeout=8000)
            self._log(f"[booking-browser] พิมพ์เลขคำขอ {req_no} ในช่องค้นหา")
            page.wait_for_timeout(400)
            # กด Enter หรือ ค้นหา ปุ่ม
            try:
                inp.press("Enter")
            except Exception:
                pass
            btn = page.get_by_role("button", name="ค้นหา").first
            try:
                btn.click(timeout=3000)
            except Exception:
                pass
            page.wait_for_timeout(2500)
        except Exception as e:
            self._log(f"[booking-browser] search error: {e}")
            history_entry["result"] = "failed"
            history_entry["error"] = f"search: {e}"[:200]
            self._log_history(history_entry)
            return False
        # 3. คลิก record — หา a[onclick*="openDetail"] ที่มี req_no
        try:
            # Row มี text = req_no → ในนั้นมีปุ่ม 'ดูรายละเอียด' หรือ link
            row = page.locator("tr", has_text=req_no).first
            row.wait_for(timeout=8000)
            # ลอง click ปุ่ม 'ดูรายละเอียด' ก่อน
            btn_detail = row.locator("a, button").filter(has_text="ดูรายละเอียด").first
            if btn_detail.count():
                btn_detail.click(timeout=5000)
            else:
                # fallback: คลิก link แรก
                row.locator("a").first.click(timeout=5000)
            self._log("[booking-browser] คลิก record แล้ว → รอ Detail")
            page.wait_for_load_state("domcontentloaded", timeout=20_000)
            page.wait_for_timeout(2500)
        except Exception as e:
            self._log(f"[booking-browser] คลิก record error: {e}")
            history_entry["result"] = "failed"
            history_entry["error"] = f"click record: {e}"[:200]
            self._log_history(history_entry)
            return False
        # 4. คลิก tab 'การนัดหมาย'
        try:
            tab = page.locator('a[href="#tab_default_5"]').first
            tab.wait_for(timeout=8000)
            tab.click(timeout=5000)
            self._log("[booking-browser] เปิด tab การนัดหมาย")
            page.wait_for_timeout(3000)
        except Exception as e:
            self._log(f"[booking-browser] คลิก tab error: {e}")
            history_entry["result"] = "failed"
            history_entry["error"] = f"click tab: {e}"[:200]
            self._log_history(history_entry)
            return False
        # 5. อ่าน iframe src
        try:
            iframe_src = page.evaluate(
                "() => (document.querySelector('#link_appointment') || {}).src || ''"
            )
            if not iframe_src or iframe_src == "about:blank":
                self._log("[booking-browser] ⚠ ไม่มี iframe หรือยังไม่โหลด (record นี้อาจไม่มีสิทธิ์จอง)")
                history_entry["result"] = "failed"
                history_entry["error"] = "no iframe (no permission?)"
                self._log_history(history_entry)
                return False
        except Exception as e:
            self._log(f"[booking-browser] อ่าน iframe error: {e}")
            history_entry["result"] = "failed"
            history_entry["error"] = f"iframe: {e}"[:200]
            self._log_history(history_entry)
            return False
        # 6. เปิด iframe ใน tab ใหม่ → auto-select
        try:
            qp_tab = ctx.new_page()
            qp_tab.goto(iframe_src, wait_until="networkidle", timeout=30_000)
            qp_tab.wait_for_timeout(3500)
            self._log("[booking-browser] เปิดหน้าจองใน tab ใหม่ — เริ่ม auto-select")
            success = self._auto_select(qp_tab, slot, auto_confirm=auto_confirm)
            # 7. ถ้า auto_confirm สำเร็จ → download ใบนัดหมาย
            slip_path = ""
            if auto_confirm and success:
                slip_path = self._download_appointment_slip(page, req_no)
                if slip_path:
                    history_entry["slip_path"] = slip_path
                    self._log(f"[booking-browser] 📄 บันทึกใบนัดหมาย: {Path(slip_path).name}")
            history_entry["result"] = "confirmed" if (auto_confirm and success) else (
                "pre-selected" if success else "partial"
            )
            self._log_history(history_entry)
            # cleanup: ปิด iframe tab เฉพาะกรณี auto_confirm success (จองเสร็จแล้ว)
            # กรณี semi-auto ให้ tab ค้างไว้เพื่อ user กด "ยืนยัน" เอง
            if auto_confirm and success:
                try:
                    qp_tab.close()
                except Exception:
                    pass
            return success
        except Exception as e:
            self._log(f"[booking-browser] เปิด iframe error: {e}")
            history_entry["result"] = "failed"
            history_entry["error"] = f"iframe open: {e}"[:200]
            self._log_history(history_entry)
            # cleanup iframe tab ถ้าเปิดค้าง
            if qp_tab is not None:
                try:
                    qp_tab.close()
                except Exception:
                    pass
            return False

    def _download_appointment_slip(self, page, req_no: str) -> str:
        """หลังจองสำเร็จ → กลับหน้า Detail → download ใบนัดหมาย
        return: path ของ PDF ที่บันทึก หรือ '' ถ้าล้มเหลว
        """
        try:
            from scrape_wa import _bt30_ctn_download_appointment
        except Exception as e:
            self._log(f"[booking-browser] import error: {e}")
            return ""
        try:
            # กลับหน้า Detail ของ record (page ยังอยู่ที่ Detail)
            # reload เพื่อให้ iframe show ใบนัดหมายใหม่ (หลัง booking)
            self._log("[booking-browser] กำลังโหลดใบนัดหมาย...")
            page.reload(wait_until="domcontentloaded", timeout=25_000)
            page.wait_for_timeout(2500)
            # เรียก function download
            pdf_bytes, passport, name, err = _bt30_ctn_download_appointment(
                page, log=lambda *a: None,
            )
            if err or not pdf_bytes:
                self._log(f"[booking-browser] ⚠ ดาวน์โหลดใบนัดหมาย: {err or 'PDF ว่าง'}")
                return ""
            # save
            self.SLIPS_DIR.mkdir(parents=True, exist_ok=True)
            import time as _t
            ts = _t.strftime("%Y%m%d_%H%M%S")
            safe_pp = (passport or "").replace(" ", "").replace("/", "")[:20] or "unknown"
            fn = f"{ts}_{req_no}_{safe_pp}.pdf"
            path = self.SLIPS_DIR / fn
            path.write_bytes(pdf_bytes)
            return str(path)
        except Exception as e:
            self._log(f"[booking-browser] download slip error: {e}")
            return ""

    def _auto_select(self, qp, slot: dict, auto_confirm: bool = False) -> bool:
        """เลือกสาขา → คลิก 'เพิ่มวันนัดหมาย' → เปลี่ยน month → คลิกวันที่ → คลิกช่วงเวลา
        return True = สำเร็จทุก step, False = มี step ที่ error
        """
        code = slot.get("branch_code", "")
        date_str = slot.get("date", "")  # YYYY-MM-DD
        rnd = slot.get("round", "")       # "09:00 - 09:30"
        # 1. เปิด dropdown สาขา
        try:
            qp.get_by_role("button", name="กรุณาเลือกสถานที่").first.click(timeout=8000)
            qp.wait_for_timeout(1500)
            # หา LI ที่มีรหัสสาขา (title/text = ชื่อสาขา) — ใช้ text match กับส่วนที่ตรง
            branch_name = slot.get("branch_name", "")
            if branch_name:
                li = qp.locator("li", has_text=branch_name).first
            else:
                li = qp.locator("li").first
            li.click(timeout=8000)
            self._log(f"[booking-browser] เลือกสาขา {code} แล้ว")
            qp.wait_for_timeout(3000)
        except Exception as e:
            self._log(f"[booking-browser] เลือกสาขา error: {e}")
            return False
        # 2. คลิก 'เพิ่มวันนัดหมาย'
        try:
            qp.get_by_role("button", name="เพิ่มวันนัดหมาย").first.click(timeout=8000)
            qp.wait_for_timeout(3500)
        except Exception as e:
            self._log(f"[booking-browser] คลิก 'เพิ่มวันนัดหมาย' error: {e}")
            return False
        # 3. เปลี่ยน month/year ให้ตรง date_str
        try:
            y, m, d = (int(x) for x in date_str.split("-"))
            # SELECT: option value ของเดือน = 0..11 (index)
            month_idx = m - 1
            qp.evaluate(
                """(args) => {
                    const [monthIdx, year] = args;
                    const sels = [...document.querySelectorAll('select')].filter(el => el.offsetParent);
                    for (const s of sels) {
                        const vals = [...s.options].map(o => o.value);
                        if (vals.length <= 13 && vals.includes(String(monthIdx))) {
                            s.value = String(monthIdx);
                            s.dispatchEvent(new Event('change', {bubbles:true}));
                        } else if (vals.includes(String(year))) {
                            s.value = String(year);
                            s.dispatchEvent(new Event('change', {bubbles:true}));
                        }
                    }
                }""",
                [month_idx, y],
            )
            qp.wait_for_timeout(3500)  # รอ calendar refresh
        except Exception as e:
            self._log(f"[booking-browser] เปลี่ยน month/year error: {e}")
        # 4. คลิกวันที่
        try:
            y, m, d = (int(x) for x in date_str.split("-"))
            # หา DIV ที่มี text = day + cursor=pointer
            cell = qp.evaluate(
                """(day) => {
                    const cells = [...document.querySelectorAll('.grid.grid-cols-7 > div')];
                    for (const el of cells) {
                        const s = getComputedStyle(el);
                        const txt = (el.textContent||'').replace(/\\s+/g,' ').trim();
                        if (txt === String(day) && s.cursor === 'pointer') {
                            const r = el.getBoundingClientRect();
                            return {x: r.x + r.width/2, y: r.y + r.height/2};
                        }
                    }
                    return null;
                }""",
                d,
            )
            if cell:
                qp.mouse.click(cell["x"], cell["y"])
                self._log(f"[booking-browser] คลิกวัน {d}")
                qp.wait_for_timeout(3000)
            else:
                self._log(f"[booking-browser] ⚠ ไม่พบวันที่ {d} ที่ enable ในปฏิทิน")
                return False
        except Exception as e:
            self._log(f"[booking-browser] คลิกวันที่ error: {e}")
            return False
        # 5. คลิก time slot (ช่วงเวลา)
        #    React/Tailwind: หา element ที่ "คลิกได้" (cursor:pointer / button) และ text ตรงรอบเวลา
        #    แล้วคลิกที่พิกัดกลาง — เลี่ยง locator('*', has_text=rnd) ที่ไปชน ancestor (html/body)
        try:
            slot_hit = qp.evaluate(
                r"""(rnd) => {
                    const norm = s => (s || '').replace(/\s+/g, '').replace(/[\u2013\u2014]/g, '-');
                    const want = norm(rnd);                             // "09:00-09:30"
                    const times = (rnd.match(/\d{1,2}:\d{2}/g) || []);  // ["09:00","09:30"]
                    const vis = el => !!(el && el.offsetParent !== null);
                    const all = Array.from(document.querySelectorAll(
                        'button, [role="button"], a, li, td, div, span')).filter(vis);

                    function dumpSlots() {
                        const out = [];
                        for (const el of all) {
                            const raw = (el.textContent || '').replace(/\s+/g, ' ').trim();
                            if (/\d{1,2}:\d{2}/.test(raw) && raw.length < 60) {
                                out.push({tag: el.tagName,
                                          cls: (el.className || '').toString().slice(0, 120),
                                          text: raw, cursor: getComputedStyle(el).cursor});
                            }
                        }
                        return out.slice(0, 80);
                    }

                    const cands = [];
                    for (const el of all) {
                        const t = norm(el.textContent);
                        if (!t) continue;
                        const hit = t.includes(want) ||
                            (times.length === 2 && t.includes(times[0]) && t.includes(times[1]));
                        if (!hit) continue;
                        const cs = getComputedStyle(el);
                        const clickable = cs.cursor === 'pointer' ||
                            el.tagName === 'BUTTON' || el.getAttribute('role') === 'button' ||
                            typeof el.onclick === 'function';
                        cands.push({el, len: t.length, clickable});
                    }
                    if (!cands.length) return {ok: false, dump: dumpSlots()};
                    // clickable ก่อน แล้วเลือก text สั้นสุด (innermost = ปุ่มช่วงเวลาจริง)
                    cands.sort((a, b) =>
                        (a.clickable === b.clickable) ? (a.len - b.len) : (a.clickable ? -1 : 1));
                    const best = cands[0].el;
                    try { best.scrollIntoView({block: 'center', inline: 'center'}); } catch (e) {}
                    const r = best.getBoundingClientRect();
                    return {ok: true, x: r.x + r.width / 2, y: r.y + r.height / 2,
                            tag: best.tagName,
                            cls: (best.className || '').toString().slice(0, 80),
                            text: (best.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 60)};
                }""",
                rnd,
            )
            if not slot_hit or not slot_hit.get("ok"):
                # เซฟ DOM ช่วงเวลาไว้ debug (เผื่อ selector ยังไม่ตรงกับหน้าเว็บจริง)
                try:
                    dump = (slot_hit or {}).get("dump") or []
                    dbg = ROOT / "_booking_timeslot_dom.json"
                    dbg.write_text(
                        json.dumps({"want": rnd, "candidates": dump}, ensure_ascii=False, indent=2),
                        encoding="utf-8",
                    )
                    self._log(
                        f"[booking-browser] ⚠ ไม่พบช่วงเวลา '{rnd}' — เซฟ DOM → {dbg.name} "
                        f"(มี {len(dump)} element ที่มีเวลา HH:MM)"
                    )
                except Exception:
                    self._log(f"[booking-browser] ⚠ ไม่พบช่วงเวลา '{rnd}' (เซฟ DOM ไม่สำเร็จ)")
                return False
            qp.mouse.click(slot_hit["x"], slot_hit["y"])
            self._log(
                f"[booking-browser] เลือกช่วงเวลา {rnd} "
                f"(→ {slot_hit.get('tag')} .{(slot_hit.get('cls') or '')[:40]})"
            )
            qp.wait_for_timeout(1500)
        except Exception as e:
            self._log(f"[booking-browser] คลิก time slot error: {e}")
            return False

        # 6. auto_confirm → คลิก "ยืนยัน"
        if auto_confirm:
            try:
                self._log("[booking-browser] 🚨 auto_confirm=True — จะกด 'ยืนยัน' ใน 3 วิ")
                qp.wait_for_timeout(3000)
                confirm_btn = qp.get_by_role("button", name="ยืนยัน").first
                confirm_btn.click(timeout=8000)
                self._log("[booking-browser] ✓✓✓ กดยืนยันแล้ว — จองสำเร็จ (ตรวจสอบใน Chrome)")
                qp.wait_for_timeout(2000)
            except Exception as e:
                self._log(f"[booking-browser] คลิก 'ยืนยัน' error: {e}")
                return False
        else:
            self._log("[booking-browser] ✓ Pre-select เสร็จ — โปรดตรวจในหน้า Chrome + กด 'ยืนยัน' เอง")
        return True


class BookingActionDialog(tk.Toplevel):
    """Popup เล็ก ๆ ยืนยันเลขคำขอ + สั่ง Chrome ให้ไปหน้าจอง"""

    SHORTCUTS_FILE = ROOT / "_booking_shortcuts.json"  # เก็บเลขคำขอล่าสุด/mapping สาขา→req

    def __init__(self, parent: tk.Misc, slot: dict, browser: "BookingBrowser", log_cb=None):
        super().__init__(parent)
        self._slot = slot
        self._browser = browser
        self._log_cb = log_cb or (lambda *a: None)
        self._shortcuts = self._load_shortcuts()
        self.title("🎯 เตรียมจอง")
        self.geometry("560x460")
        self.transient(parent)
        self.resizable(False, False)
        self._build_ui()
        self.after(100, self._focus_center)

    def _load_shortcuts(self) -> dict:
        try:
            if self.SHORTCUTS_FILE.exists():
                data = json.loads(self.SHORTCUTS_FILE.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    return data
        except Exception:
            pass
        return {}

    def _save_shortcuts(self) -> None:
        try:
            self.SHORTCUTS_FILE.write_text(
                json.dumps(self._shortcuts, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception:
            pass

    def _focus_center(self) -> None:
        self.grab_set()
        self.lift()
        self.focus_set()

    def _build_ui(self) -> None:
        s = self._slot
        code = s.get("branch_code", "")
        name = s.get("branch_name", "")
        date = s.get("date", "")
        rnd = s.get("round", "")
        left = int(s.get("left", 0))
        cap = int(s.get("capacity", 0))

        # แปลงวันไทย
        thai_days = ["จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์", "อาทิตย์"]
        thai_months = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
                       "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
        date_th = date
        try:
            from datetime import date as _d
            y, m, dd = (int(x) for x in date.split("-"))
            dt = _d(y, m, dd)
            date_th = f"{thai_days[dt.weekday()]} {dd} {thai_months[m-1]} {y + 543}"
        except Exception:
            pass

        # Header
        tk.Label(self, text="🎯 ไปหน้าจอง", font=("TkDefaultFont", 14, "bold"),
                 fg="#0b7c3f", pady=8).pack(fill="x")

        # Info
        info = ttk.LabelFrame(self, text="รอบที่จะจอง", padding=10)
        info.pack(fill="x", padx=12, pady=6)
        for i, (k, v) in enumerate([
            ("สาขา:", f"{code} — {name}"),
            ("วัน:", date_th),
            ("เวลา:", rnd),
            ("ที่ว่าง:", f"{left}/{cap}"),
        ]):
            tk.Label(info, text=k, font=("TkDefaultFont", 10, "bold")).grid(row=i, column=0, sticky="w")
            tk.Label(info, text=v, anchor="w", wraplength=380).grid(row=i, column=1, sticky="w", padx=6)
        info.columnconfigure(1, weight=1)

        # เลขคำขอ
        req_frm = ttk.LabelFrame(self, text="เลขคำขอที่ต้องการจอง (หลายอันได้ — บรรทัดละ 1)", padding=10)
        req_frm.pack(fill="x", padx=12, pady=6)
        last_req = self._shortcuts.get("_last_req_no", "")
        self._req_text = tk.Text(
            req_frm, height=4, wrap="none", font=("TkDefaultFont", 11),
            undo=True,
        )
        self._req_text.pack(fill="x")
        if last_req:
            self._req_text.insert("1.0", last_req)
        tk.Label(
            req_frm,
            text="(หลายเลขคั่นด้วยขึ้นบรรทัดใหม่ หรือ comma — ระบบจะจองทีละอัน)",
            foreground="gray",
        ).pack(anchor="w", pady=(4, 0))

        # ตัวเลือกยืนยันอัตโนมัติ
        self._auto_confirm = tk.BooleanVar(value=False)
        confirm_frm = ttk.Frame(self, padding=(12, 4))
        confirm_frm.pack(fill="x")
        ttk.Checkbutton(
            confirm_frm,
            text="⚠ กด 'ยืนยัน' อัตโนมัติเลย (Full-Auto — จองทันที ไม่ต้องกดเอง)",
            variable=self._auto_confirm,
        ).pack(anchor="w")
        tk.Label(
            confirm_frm,
            text="   default = ไม่ติ๊ก → หยุดที่ปุ่มยืนยันให้ตรวจก่อน (ปลอดภัยกว่า)",
            foreground="gray", font=("TkDefaultFont", 9),
        ).pack(anchor="w")

        # ปุ่มหลัก
        btns = ttk.Frame(self, padding=(12, 8))
        btns.pack(fill="x")
        self._btn_launch = ttk.Button(
            btns, text="🚀  ไปหน้าจองเลย",
            command=self._go,
        )
        self._btn_launch.pack(side="left", padx=4, ipady=6, ipadx=12)
        # ปุ่มเปิด Chrome (ถ้ายังไม่เปิด)
        if not self._browser.is_running:
            self._btn_open_chrome = ttk.Button(
                btns, text="🌐  เปิด Chrome (setup)",
                command=self._open_chrome,
            )
            self._btn_open_chrome.pack(side="left", padx=4, ipady=6, ipadx=8)
        ttk.Button(btns, text="ปิด", command=self.destroy).pack(side="right", padx=4, ipady=4)

        # tip
        tip = ttk.LabelFrame(self, text="ทำครั้งแรก", padding=10)
        tip.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        if not self._browser.is_running:
            hint = (
                "⚠ ยังไม่ได้เปิด Chrome สำหรับจอง\n\n"
                "👉 กดปุ่ม 🌐 'เปิด Chrome (setup)' ด้านบน\n"
                "     → Chrome ใหม่จะเปิดขึ้นมา\n"
                "     → Login บัญชีจอง 1 ครั้ง\n"
                "     → กลับมา double-click row + กด 🚀\n"
                "\n"
                "ครั้งต่อไป: Chrome จำ login ให้ → คลิก 🚀 = auto ไปหน้าจอง"
            )
        else:
            hint = (
                "✓ Chrome พร้อมใช้งาน\n\n"
                "คลิก 🚀 → Chrome จะ:\n"
                "  1. ไปหน้า Tracking\n"
                "  2. ค้นหาเลขคำขอ\n"
                "  3. คลิก record → tab การนัดหมาย\n"
                "  4. เลือกสาขา / วัน / เวลา ให้ auto\n"
                "  5. หยุดที่ปุ่มยืนยัน → คุณตรวจ + กด 'ยืนยัน' เอง"
            )
        tk.Label(tip, text=hint, justify="left", anchor="nw",
                 font=("TkDefaultFont", 9)).pack(fill="both", expand=True)

    def _go(self) -> None:
        req_raw = self._req_text.get("1.0", "end").strip()
        req_list = [
            r.strip() for r in req_raw.replace(",", "\n").replace(";", "\n").split("\n")
            if r.strip()
        ]
        if not req_list:
            messagebox.showwarning("ยังไม่กรอกเลขคำขอ", "กรอก 'เลขคำขอ' อย่างน้อย 1 อัน", parent=self)
            return
        if not self._browser.is_running:
            messagebox.showwarning(
                "Chrome ยังไม่เปิด",
                "กดปุ่ม 🌐 'เปิด Chrome (setup)' ด้านล่างก่อน\n"
                "แล้ว login บัญชีจองใน Chrome นั้น 1 ครั้ง",
                parent=self,
            )
            return
        auto_confirm = bool(self._auto_confirm.get())
        # confirm ถ้าหลายอัน + full-auto
        if auto_confirm and len(req_list) > 1:
            if not messagebox.askyesno(
                "ยืนยัน Full-Auto หลายเลข",
                f"จะจอง {len(req_list)} เลขคำขอ ในสาขา/วัน/เวลาเดียวกัน\n"
                f"+ กด 'ยืนยัน' ให้อัตโนมัติทุกเลข\n\n"
                f"ต้องการดำเนินการต่อ?",
                parent=self,
            ):
                return
        # เก็บ req_no ล่าสุด (รวมทั้ง list)
        self._shortcuts["_last_req_no"] = "\n".join(req_list)
        self._save_shortcuts()
        # สั่ง browser
        for req in req_list:
            self._browser.goto_record(req, self._slot, auto_confirm=auto_confirm)
        mode_txt = "🚨 FULL-AUTO" if auto_confirm else "SEMI-AUTO (หยุดก่อนยืนยัน)"
        self._log_cb(
            f"[monitor] → สั่ง Chrome จอง {len(req_list)} record ({mode_txt})"
        )
        self.destroy()

    def _open_chrome(self) -> None:
        """เปิด Chrome จองจากใน popup"""
        if self._browser.is_running:
            messagebox.showinfo("Chrome เปิดอยู่แล้ว", "Chrome สำหรับจองกำลังทำงานอยู่", parent=self)
            return
        self._browser.start()
        messagebox.showinfo(
            "กำลังเปิด Chrome",
            "Chrome สำหรับจองกำลังเปิด...\n\n"
            "ครั้งแรกให้ login บัญชีจองใน Chrome นั้น 1 ครั้ง\n"
            "หลังจากนั้นค่อยกด 🚀 'ไปหน้าจองเลย'",
            parent=self,
        )
        # ปิด popup เพื่อให้ user login ได้สะดวก
        self.destroy()


class BookingHistoryDialog(tk.Toplevel):
    """หน้าต่างแสดงประวัติจองทั้งหมด (อ่านจาก _booking_history.jsonl)"""

    HISTORY_FILE = ROOT / "_booking_history.jsonl"

    def __init__(self, parent: tk.Misc):
        super().__init__(parent)
        self.title("📊 ประวัติจอง")
        self.geometry("1100x600")
        self.transient(parent)
        self._build_ui()
        self._load_history()

    def _build_ui(self) -> None:
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="🔄  Refresh", command=self._load_history).pack(side="left")
        ttk.Label(top, text="  ค้นหา:").pack(side="left", padx=(12, 4))
        self._search = tk.StringVar(value="")
        ttk.Entry(top, textvariable=self._search, width=25).pack(side="left", padx=4)
        self._search.trace_add("write", lambda *_: self._load_history())
        ttk.Button(
            top, text="📁 เปิดไฟล์",
            command=lambda: self._open_file(),
        ).pack(side="right")
        ttk.Button(top, text="ปิด", command=self.destroy).pack(side="right", padx=8)

        # summary
        self._summary = tk.StringVar(value="")
        ttk.Label(self, textvariable=self._summary, padding=(8, 4),
                  foreground="#1a5490", font=("TkDefaultFont", 10, "bold")).pack(fill="x")

        # tree
        cols = ("ts", "req_no", "code", "date", "round", "cap", "left", "mode", "result", "slip", "error")
        self._tree = ttk.Treeview(self, columns=cols, show="headings", height=22)
        headings = [
            ("ts", "เวลา", 130),
            ("req_no", "เลขคำขอ", 130),
            ("code", "รหัสสาขา", 120),
            ("date", "วันที่", 90),
            ("round", "ช่วงเวลา", 110),
            ("cap", "รวม", 60),
            ("left", "เหลือ", 60),
            ("mode", "โหมด", 80),
            ("result", "ผล", 90),
            ("slip", "ใบนัดหมาย", 90),
            ("error", "หมายเหตุ", 200),
        ]
        for k, t, w in headings:
            self._tree.heading(k, text=t)
            anchor = "center" if k in ("cap", "left", "mode", "slip") else "w"
            self._tree.column(k, width=w, anchor=anchor)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._tree.pack(side="left", fill="both", expand=True)
        # double-click → เปิด PDF ใบนัดหมาย
        self._row_slips: dict[str, str] = {}  # iid → slip_path
        self._tree.bind("<Double-1>", self._open_slip)
        # tags
        self._tree.tag_configure("confirmed", background="#C6EFCE")
        self._tree.tag_configure("pre-selected", background="#E8F0EA")
        self._tree.tag_configure("failed", background="#F5B7B1")
        self._tree.tag_configure("partial", background="#FFEB9C")

    def _load_history(self) -> None:
        for iid in self._tree.get_children():
            self._tree.delete(iid)
        self._row_slips = {}
        if not self.HISTORY_FILE.exists():
            self._summary.set("ไม่มีประวัติ")
            return
        entries: list[dict] = []
        try:
            with self.HISTORY_FILE.open("r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    try:
                        entries.append(json.loads(line))
                    except Exception:
                        pass
        except Exception as e:
            self._summary.set(f"อ่านไฟล์ error: {e}")
            return
        # filter
        filt = (self._search.get() or "").strip().lower()
        if filt:
            entries = [
                e for e in entries
                if filt in (e.get("req_no", "") or "").lower()
                or filt in (e.get("branch_code", "") or "").lower()
                or filt in (e.get("branch_name", "") or "").lower()
                or filt in (e.get("date", "") or "").lower()
                or filt in (e.get("result", "") or "").lower()
            ]
        # sort desc by ts
        entries.sort(key=lambda e: e.get("ts", ""), reverse=True)
        n_conf = sum(1 for e in entries if e.get("result") == "confirmed")
        n_pre = sum(1 for e in entries if e.get("result") == "pre-selected")
        n_fail = sum(1 for e in entries if e.get("result") == "failed")
        self._summary.set(
            f"รวม {len(entries)} รายการ  |  ✅ ยืนยัน {n_conf}  |  📝 pre-select {n_pre}  |  ❌ error {n_fail}"
        )
        for e in entries:
            result = e.get("result", "")
            mode = "🚨 auto" if e.get("auto_confirm") else "manual"
            slip = e.get("slip_path", "")
            slip_txt = "📄 คลิก" if slip else ""
            iid = self._tree.insert(
                "", "end", values=(
                    e.get("ts", ""),
                    e.get("req_no", ""),
                    e.get("branch_code", ""),
                    e.get("date", ""),
                    e.get("round", ""),
                    e.get("capacity", ""),
                    e.get("left", ""),
                    mode,
                    result,
                    slip_txt,
                    (e.get("error", "") or "")[:80],
                ),
                tags=(result,),
            )
            if slip:
                self._row_slips[iid] = slip

    def _open_slip(self, event=None) -> None:
        """Double-click row → เปิด PDF ใบนัดหมาย"""
        import os
        sel = self._tree.selection()
        if not sel:
            return
        iid = sel[0]
        path = self._row_slips.get(iid)
        if not path:
            return
        p = Path(path)
        if not p.exists():
            messagebox.showwarning("ไม่พบไฟล์", f"ไฟล์หาย: {path}", parent=self)
            return
        try:
            os.startfile(str(p))
        except Exception as e:
            messagebox.showerror("เปิด PDF ไม่ได้", str(e), parent=self)

    def _open_file(self) -> None:
        import os
        if self.HISTORY_FILE.exists():
            try:
                os.startfile(str(self.HISTORY_FILE))
            except Exception as e:
                messagebox.showerror("เปิดไฟล์ไม่ได้", str(e), parent=self)
        else:
            messagebox.showinfo("ไม่มีไฟล์", "ยังไม่มีประวัติจอง", parent=self)


if __name__ == "__main__":
    App().mainloop()