"""GUI สำหรับดึงรายงาน 'รอยื่นเอกสารเพิ่มเติม' จาก e-WorkPermit
รัน: python gui_app.py
"""
from __future__ import annotations

import json
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

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


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("ดึงรายงาน e-WorkPermit (รอยื่นเอกสารเพิ่มเติม)")
        self.geometry("920x820")
        self.minsize(780, 700)
        try:
            self.tk.call("tk", "scaling", 1.2)
        except Exception:
            pass

        self._log_q: queue.Queue[str] = queue.Queue()
        self._worker: threading.Thread | None = None
        self._cancel = threading.Event()
        # _pause_evt: set = ทำงานปกติ, clear = พักค้างไว้ (worker จะรอที่จุดปลอดภัย)
        self._pause_evt = threading.Event()
        self._pause_evt.set()

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

        self._build_ui()
        self.after(120, self._drain_log)

    # ---------- UI ----------
    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 6}

        # ---- Scrollable outer container ----
        # ห่อ content ทั้งหมดด้วย Canvas + Scrollbar เพื่อให้ผู้ใช้ scroll
        # ลงไปเห็นปุ่ม/log ได้ เมื่อจอไม่พอ (ตัวเลือกเยอะ)
        outer = ttk.Frame(self)
        outer.pack(fill="both", expand=True)
        self._scroll_canvas = tk.Canvas(outer, borderwidth=0, highlightthickness=0)
        _vsb = ttk.Scrollbar(outer, orient="vertical", command=self._scroll_canvas.yview)
        self._scroll_canvas.configure(yscrollcommand=_vsb.set)
        _vsb.pack(side="right", fill="y")
        self._scroll_canvas.pack(side="left", fill="both", expand=True)
        self._body = ttk.Frame(self._scroll_canvas)
        self._body_window = self._scroll_canvas.create_window(
            (0, 0), window=self._body, anchor="nw",
        )

        def _on_body_configure(_evt=None):
            self._scroll_canvas.configure(scrollregion=self._scroll_canvas.bbox("all"))

        def _on_canvas_configure(evt):
            # ให้ inner frame มีความกว้างเท่า canvas เสมอ (fill x)
            self._scroll_canvas.itemconfig(self._body_window, width=evt.width)

        self._body.bind("<Configure>", _on_body_configure)
        self._scroll_canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(evt):
            # Windows/macOS: evt.delta คูณ 120; Linux: ใช้ Button-4/5 (ไม่ครอบคลุมที่นี่)
            # ถ้าเมาส์อยู่บน widget ที่ scroll เองได้ (Listbox / Text) → ปล่อยให้มัน scroll เอง
            try:
                w = evt.widget
                p = w
                while p is not None:
                    if isinstance(p, (tk.Listbox, tk.Text)):
                        return
                    p = getattr(p, "master", None)
            except Exception:
                pass
            try:
                self._scroll_canvas.yview_scroll(int(-1 * (evt.delta / 120)), "units")
            except Exception:
                pass

        # bind mouse wheel ให้ทำงานทั่วหน้าต่าง (เฉพาะตอน hover บน canvas/inner)
        self._scroll_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # ---- โหมดการทำงาน ----
        mode_frm = ttk.LabelFrame(self._body, text="โหมดการทำงาน", padding=10)
        mode_frm.pack(fill="x", **pad)
        self.source_mode = tk.StringVar(value="etracking")
        ttk.Radiobutton(
            mode_frm, text="e-Tracking — ดึงรายการคำขอ (รายงาน WA)",
            variable=self.source_mode, value="etracking",
            command=self._on_mode_changed,
        ).pack(anchor="w")
        ttk.Radiobutton(
            mode_frm, text="จัดการบัญชี → ข้อมูลคนต่างด้าว (รายชื่อคนต่างด้าวทั้งหมด)",
            variable=self.source_mode, value="aliens",
            command=self._on_mode_changed,
        ).pack(anchor="w")
        ttk.Radiobutton(
            mode_frm, text="ลงทะเบียนคนต่างด้าวจาก Excel (Register + ดึงข้อมูลใบอนุญาต)",
            variable=self.source_mode, value="register",
            command=self._on_mode_changed,
        ).pack(anchor="w")
        ttk.Radiobutton(
            mode_frm, text="ดาวน์โหลดใบเสร็จ",
            variable=self.source_mode, value="receipts",
            command=self._on_mode_changed,
        ).pack(anchor="w")
        ttk.Radiobutton(
            mode_frm, text="ดาวน์โหลดใบแจ้งผล / ใบรับคำขอ (ตามรายการคำขอ)",
            variable=self.source_mode, value="results",
            command=self._on_mode_changed,
        ).pack(anchor="w")
        ttk.Radiobutton(
            mode_frm, text="แจ้งเข้านายจ้าง (แบบ บต.52) — INFORM_ENTER_EXIT",
            variable=self.source_mode, value="inform",
            command=self._on_mode_changed,
        ).pack(anchor="w")
        ttk.Radiobutton(
            mode_frm, text="ยื่นต่อใบอนุญาตทำงาน (แบบ บต.30) — MoU: กรอกค้นหาข้อมูลคนต่างด้าว",
            variable=self.source_mode, value="bt30",
            command=self._on_mode_changed,
        ).pack(anchor="w")
        ttk.Radiobutton(
            mode_frm, text="เปลี่ยนย้ายนายจ้างในระบบ (บต.44)",
            variable=self.source_mode, value="bt44",
            command=self._on_mode_changed,
        ).pack(anchor="w")
        ttk.Radiobutton(
            mode_frm, text="ดาวน์โหลดใบแจ้งชำระเงิน (รอจ่ายค่าธรรมเนียม) — หลายบัญชีจาก Excel",
            variable=self.source_mode, value="bill_payment",
            command=self._on_mode_changed,
        ).pack(anchor="w")
        ttk.Radiobutton(
            mode_frm, text="ดาวน์โหลดใบเสร็จรับเงินทั้งชุด (900+100+อื่นๆ) — ตามเลขคำขอจาก Excel",
            variable=self.source_mode, value="payment_receipts",
            command=self._on_mode_changed,
        ).pack(anchor="w")
        ttk.Radiobutton(
            mode_frm, text="นัดหมายถ่ายบัตร — เก็บที่อยู่จากใบเสร็จค่าธรรมเนียมใบอนุญาตทำงาน (Status=AP)",
            variable=self.source_mode, value="appointment",
            command=self._on_mode_changed,
        ).pack(anchor="w")
        ttk.Radiobutton(
            mode_frm, text="ดาวน์โหลด (ใบตอบรับ / ใบนัดหมาย)",
            variable=self.source_mode, value="bt30_ctn",
            command=self._on_mode_changed,
        ).pack(anchor="w")
        ttk.Radiobutton(
            mode_frm, text="ดึงรายชื่อคนต่างด้าวที่ยื่นคำขอต่ออายุแล้ว (NameListAlien — ทุก pagination)",
            variable=self.source_mode, value="namelist_alien",
            command=self._on_mode_changed,
        ).pack(anchor="w")
        ttk.Radiobutton(
            mode_frm, text="ตรวจวันว่างจอง (คิวถ่ายบัตร) — Real-Time Check วัน+ช่วงเวลาที่ยังว่างของทุกสาขา",
            variable=self.source_mode, value="booking_availability",
            command=self._on_mode_changed,
        ).pack(anchor="w")
        ttk.Radiobutton(
            mode_frm, text="ดึงรายงานข้อมูลการขออนุญาต — สถานะคำขอ (อนุมัติคำขอ) + สถานประกอบการ (บริษัท/จังหวัด)",
            variable=self.source_mode, value="permit_report",
            command=self._on_mode_changed,
        ).pack(anchor="w")

        frm = ttk.LabelFrame(self._body, text="ข้อมูลเข้าสู่ระบบ", padding=10)
        frm.pack(fill="x", **pad)

        ttk.Label(frm, text="Username (อีเมล):").grid(row=0, column=0, sticky="w", pady=4)
        self.username = tk.StringVar()
        ttk.Entry(frm, textvariable=self.username, width=44).grid(row=0, column=1, sticky="we", padx=8)

        ttk.Label(frm, text="Password:").grid(row=1, column=0, sticky="w", pady=4)
        self.password = tk.StringVar()
        self.pw_entry = ttk.Entry(frm, textvariable=self.password, width=44, show="•")
        self.pw_entry.grid(row=1, column=1, sticky="we", padx=8)
        self.show_pw = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            frm, text="แสดงรหัสผ่าน", variable=self.show_pw,
            command=self._toggle_pw,
        ).grid(row=1, column=2, sticky="w")

        ttk.Label(frm, text="ประเภทผู้ใช้:").grid(row=2, column=0, sticky="w", pady=4)
        self.user_type = tk.StringVar(value="ผู้กระทำการแทน")
        ttk.Combobox(
            frm, textvariable=self.user_type, width=42, state="readonly",
            values=[
                "ผู้กระทำการแทน",
                "นายจ้าง-คนไทย",
                "นายจ้าง-คนต่างชาติ",
                "บริษัทนำเข้า (บนจ.)",
                "คนต่างด้าว",
            ],
        ).grid(row=2, column=1, sticky="we", padx=8)

        ttk.Label(frm, text="ระบบ:").grid(row=3, column=0, sticky="w", pady=4)
        self.method = tk.StringVar(value="E-Workpermit")
        ttk.Combobox(
            frm, textvariable=self.method, width=42, state="readonly",
            values=["E-Workpermit", "E-Service"],
        ).grid(row=3, column=1, sticky="we", padx=8)

        # ซ่อนหน้าต่างเบราว์เซอร์ — ใช้ได้กับทุกโหมด
        # หมายเหตุ: headless จริงถูกซ่อนจาก UI (กันผู้ใช้สับสน) — ใช้โหมด "ซ่อนนอกจอ" แทน
        self.headless = tk.BooleanVar(value=False)

        # ซ่อนแบบเปิดเบราว์เซอร์จริง (ไม่ headless) แต่ย้ายหน้าต่างออกนอกจอ
        # เหมาะกับเว็บที่ตรวจจับ/บล็อก headless — ทำงานเหมือน browser จริงทุกอย่าง
        self.hide_window = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            frm, text="ซ่อนหน้าต่างเบราว์เซอร์ขณะทำงาน — ไม่รบกวนการใช้งานเครื่อง",
            variable=self.hide_window,
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(8, 0))

        frm.columnconfigure(1, weight=1)

        # ---- ตัวเลือก (e-Tracking) ----
        opt = ttk.LabelFrame(self._body, text="ตัวเลือก", padding=10)
        opt.pack(fill="x", **pad)
        self._etracking_frames: list[ttk.LabelFrame] = [opt]

        # ---- รูปแบบรายงาน (Template) — เฉพาะโหมด e-Tracking ----
        self.etk_template_frame = ttk.LabelFrame(
            self._body, text="รูปแบบรายงาน (Template)", padding=10
        )
        self.etk_template_frame.pack(fill="x", **pad)
        self.etk_template = tk.StringVar(value="main")
        ttk.Radiobutton(
            self.etk_template_frame,
            text="Template หลัก — รายละเอียดครบทุกคอลัมน์ (ค่าเริ่มต้น)",
            variable=self.etk_template, value="main",
        ).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(
            self.etk_template_frame,
            text=("Template นัดหมาย — ลำดับ | Username | เลขคำขอ | ชื่อบริษัท | "
                  "ชื่อแรงงาน | วันที่นัดหมาย | เวลานัด | สถานที่"),
            variable=self.etk_template, value="appointment",
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))
        ttk.Label(
            self.etk_template_frame,
            text=("Template นัดหมาย จะเข้าแท็บ 'การนัดหมาย' ของแต่ละรายการ"
                  "เพื่อดึงวันที่/เวลา/สถานที่ (ทำงานช้าลงเล็กน้อย) — เลือกกรองสถานะ "
                  "'รอนัดหมาย/นัดหมายแล้ว' ในตัวกรองด้านล่างเอง"),
            foreground="gray", wraplength=780, justify="left",
        ).grid(row=2, column=0, sticky="w", pady=(4, 0))

        # รายการคำขอ (เลือกได้หลายรายการ)
        # default = MT_59_MOU_RENEWAL (รายการที่ใช้บ่อยที่สุด)
        _default_code = "MT_59_MOU_RENEWAL"
        # เก็บ StringVar สำหรับ backward compat (ยังใช้กับ mode อื่นเช่น appointment ที่รับ request_type เดียว)
        self.request_type = tk.StringVar(value=_default_code)
        self.request_label = tk.StringVar(value="")

        # header row: label + ปุ่มเลือกทั้งหมด/ล้าง
        rt_head = ttk.Frame(opt)
        rt_head.grid(row=0, column=0, columnspan=3, sticky="we", pady=(0, 2))
        ttk.Label(
            rt_head, text="รายการคำขอ (เลือกได้หลายรายการ — Ctrl/Shift + คลิก):",
        ).pack(side="left")
        ttk.Button(
            rt_head, text="เลือกทั้งหมด", command=self._req_select_all, width=12,
        ).pack(side="right", padx=(4, 0))
        ttk.Button(
            rt_head, text="ล้าง", command=self._req_clear, width=8,
        ).pack(side="right")

        # Listbox + scrollbars — ข้าม entry แรกที่ code=="" (ทั้งหมด/ไม่กรอง)
        self._req_codes: list[str] = [c for c, _ in REQUEST_TYPES if c]
        self._req_labels: list[str] = [l for c, l in REQUEST_TYPES if c]
        self._req_visible_idx: list[int] = []                 # row-in-listbox → master idx
        self._req_selected_codes: set[str] = {_default_code}  # เก็บ selection ข้าม filter
        rt_box = ttk.Frame(opt)
        rt_box.grid(row=1, column=0, columnspan=3, sticky="we", pady=(0, 4))
        rt_box.columnconfigure(0, weight=1)
        # ช่องค้นหา (กรองรายการในลิสต์ตามชื่อ/โค้ด)
        rt_search_row = ttk.Frame(rt_box)
        rt_search_row.grid(row=0, column=0, columnspan=2, sticky="we", pady=(0, 3))
        ttk.Label(rt_search_row, text="🔍 ค้นหา:").pack(side="left")
        self.request_search = tk.StringVar(value="")
        ttk.Entry(rt_search_row, textvariable=self.request_search).pack(
            side="left", fill="x", expand=True, padx=(4, 4),
        )
        self.request_search.trace_add("write", lambda *_: self._req_apply_search())
        ttk.Button(
            rt_search_row, text="✕", width=3,
            command=lambda: self.request_search.set(""),
        ).pack(side="left")
        vsb = ttk.Scrollbar(rt_box, orient="vertical")
        hsb = ttk.Scrollbar(rt_box, orient="horizontal")
        self.request_listbox = tk.Listbox(
            rt_box, selectmode=tk.EXTENDED, height=8, exportselection=False,
            yscrollcommand=vsb.set, xscrollcommand=hsb.set,
            activestyle="dotbox",
        )
        vsb.config(command=self.request_listbox.yview)
        hsb.config(command=self.request_listbox.xview)
        self.request_listbox.grid(row=1, column=0, sticky="we")
        vsb.grid(row=1, column=1, sticky="ns")
        hsb.grid(row=2, column=0, sticky="we")
        self.request_listbox.bind("<<ListboxSelect>>", self._on_request_changed)
        # เติมรายการครั้งแรก (renderer จะ re-select default ให้) + เลื่อนไปหาที่เลือก
        self._render_request_listbox("")
        try:
            _idx = next(
                (r for r, m in enumerate(self._req_visible_idx)
                 if self._req_codes[m] == _default_code),
                None,
            )
            if _idx is not None:
                self.request_listbox.see(_idx)
        except Exception:
            pass

        # แสดงจำนวนที่เลือก + hint
        self.request_summary = tk.StringVar(value="")
        ttk.Label(opt, textvariable=self.request_summary, foreground="#1a7000").grid(
            row=2, column=0, columnspan=3, sticky="w",
        )
        ttk.Label(
            opt, text="(ไม่เลือกอะไรเลย = ดึงทุกรายการคำขอ)", foreground="gray",
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(0, 4))

        # ---- ตัวกรองวันที่ยื่นคำขอ (จาก → ถึง) ----
        date_row = ttk.Frame(opt)
        date_row.grid(row=4, column=0, columnspan=3, sticky="we", pady=(4, 0))
        ttk.Label(date_row, text="วันที่ยื่นคำขอ (จาก):").pack(side="left")
        self.date_from = tk.StringVar(value="")
        ttk.Entry(date_row, textvariable=self.date_from, width=14).pack(
            side="left", padx=(4, 12),
        )
        ttk.Label(date_row, text="ถึง:").pack(side="left")
        self.date_to = tk.StringVar(value="")
        ttk.Entry(date_row, textvariable=self.date_to, width=14).pack(
            side="left", padx=(4, 12),
        )
        ttk.Label(
            date_row,
            text="รูปแบบ วว/ดด/ปปปป (เช่น 01/07/2026) — เว้นว่างทั้งสองช่อง = ทุกวันที่",
            foreground="gray",
        ).pack(side="left")

        # ---- หลายบัญชี (multi-user) ----
        self.etk_multi = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            opt, text="ดึงหลายบัญชีจากไฟล์ UsernameLogin.xlsx (ไม่ต้องกรอก Username/Password ด้านบน)",
            variable=self.etk_multi, command=self._on_etk_multi_changed,
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(6, 0))
        self.etk_login_lbl = ttk.Label(opt, text="ไฟล์ UsernameLogin.xlsx:")
        self.etk_login_lbl.grid(row=6, column=0, sticky="w", pady=4)
        self.etk_login_input = tk.StringVar(value=str(ROOT / "UsernameLogin.xlsx"))
        self.etk_login_entry = ttk.Entry(opt, textvariable=self.etk_login_input)
        self.etk_login_entry.grid(row=6, column=1, sticky="we", padx=8)
        self.etk_login_btn = ttk.Button(
            opt, text="เลือก...",
            command=lambda: self._choose_into(self.etk_login_input),
        )
        self.etk_login_btn.grid(row=6, column=2, padx=4)

        # ตัวเลือก: รวมทุกบัญชีเป็นไฟล์เดียว หรือ แยกไฟล์ต่อบัญชี (เฉพาะโหมดหลายบัญชี)
        self.etk_combine = tk.BooleanVar(value=False)
        self.etk_combine_cb = ttk.Checkbutton(
            opt,
            text="รวมทุกบัญชีเป็นรายงานไฟล์เดียว (มีคอลัมน์ 'บัญชี (Username)' แยกแถว) — ไม่ติ๊ก = แยกไฟล์ต่อบัญชี",
            variable=self.etk_combine,
        )
        self.etk_combine_cb.grid(row=7, column=0, columnspan=3, sticky="w", pady=(0, 2))

        ttk.Label(opt, text="Login ใหม่ทุก (รายการ):").grid(row=8, column=0, sticky="w", pady=4)
        self.relogin_every = tk.StringVar(value="500")
        # state="normal" = พิมพ์ตัวเลขเองได้ (ไม่จำกัดแค่ค่าใน dropdown)
        ttk.Combobox(
            opt, textvariable=self.relogin_every, width=10, state="normal",
            values=["ปิด (ไม่ login ใหม่)", "100", "200", "300", "500", "1000", "2000"],
        ).grid(row=8, column=1, sticky="w", padx=8)
        ttk.Label(opt, text="พิมพ์จำนวนเองได้ (0/ว่าง = ปิด) — กัน session timeout ในงานยาว",
                  foreground="gray").grid(
            row=8, column=2, sticky="w",
        )

        ttk.Label(opt, text="จำกัดจำนวน (0 = ทั้งหมด):").grid(row=9, column=0, sticky="w", pady=4)
        self.limit = tk.IntVar(value=0)
        ttk.Spinbox(opt, from_=0, to=10000, textvariable=self.limit, width=10).grid(
            row=9, column=1, sticky="w", padx=8,
        )

        ttk.Label(opt, text="ไฟล์ที่บันทึก:").grid(row=10, column=0, sticky="w", pady=4)
        self.out_path = tk.StringVar(value=str(REPORTS_DIR / "WA_report.xlsx"))
        ttk.Entry(opt, textvariable=self.out_path).grid(row=10, column=1, sticky="we", padx=8)
        ttk.Button(opt, text="เลือก...", command=self._choose_out).grid(row=10, column=2, padx=4)
        opt.columnconfigure(1, weight=1)

        # ---- ตัวกรองสถานะ (advanced) ----
        adv = ttk.LabelFrame(self._body, text="ตัวกรองสถานะ (ตามที่เห็นในเว็บ)", padding=10)
        adv.pack(fill="x", **pad)
        self._etracking_frames.append(adv)

        ttk.Label(
            adv, text="ติ๊ก checkbox สถานะที่จะดึง (หน้า e-Tracking):", foreground="#444"
        ).grid(row=0, column=0, columnspan=5, sticky="w")
        self.filter_vars: dict[str, tk.BooleanVar] = {}
        for i, (code, label) in enumerate(STATUS_FILTERS):
            v = tk.BooleanVar(value=False)
            self.filter_vars[code] = v
            ttk.Checkbutton(adv, text=f"{label}", variable=v).grid(
                row=1, column=i, sticky="w", padx=4, pady=2
            )

        ttk.Separator(adv, orient="horizontal").grid(
            row=2, column=0, columnspan=5, sticky="we", pady=(8, 4)
        )

        ttk.Label(
            adv, text="กดเข้าไป scrape detail เฉพาะสถานะเหล่านี้ (match แบบ substring):",
            foreground="#444",
        ).grid(row=3, column=0, columnspan=5, sticky="w")
        self.whitelist_vars: dict[str, tk.BooleanVar] = {}
        for i, status in enumerate(WHITELIST_OPTIONS):
            v = tk.BooleanVar(value=True)
            self.whitelist_vars[status] = v
            ttk.Checkbutton(adv, text=status, variable=v).grid(
                row=4 + (i // 2), column=(i % 2) * 3, columnspan=3,
                sticky="w", padx=4, pady=1,
            )

        ttk.Separator(adv, orient="horizontal").grid(
            row=4 + (len(WHITELIST_OPTIONS) + 1) // 2, column=0, columnspan=5,
            sticky="we", pady=(8, 4),
        )

        self.capture_extra_notes = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            adv, text="เก็บ 'บันทึกเพิ่มเติม' จากหน้าแก้ไข (ปิดถ้าหน้าไม่มีปุ่มแก้ไข)",
            variable=self.capture_extra_notes,
        ).grid(
            row=5 + (len(WHITELIST_OPTIONS) + 1) // 2, column=0, columnspan=5,
            sticky="w", padx=4, pady=(2, 0),
        )

        # ---- ดึงเฉพาะเลขคำขอ (ถ้าระบุ) — ข้ามฟิลเตอร์สถานะ/วันที่/whitelist ด้านบนทั้งหมด ----
        # (เฉพาะโหมด e-Tracking บัญชีเดียว/หลายบัญชี — ไม่ใช้กับ bt30_ctn/permit_report จึงไม่รวมใน _etracking_frames)
        self.etk_req_no_frame = ttk.LabelFrame(
            self._body, text="ดึงเฉพาะเลขคำขอ (ทางเลือก — เว้นว่าง = ดึงตามฟิลเตอร์สถานะปกติด้านบน)",
            padding=10,
        )
        ttk.Label(
            self.etk_req_no_frame,
            text="โหมด A: ไฟล์ Excel เลขคำขอ (รองรับหลาย user พร้อมกัน — จัดกลุ่มตาม username แล้ว login ทีละบัญชีอัตโนมัติ)",
            foreground="#444",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 2))
        ttk.Label(self.etk_req_no_frame, text="ไฟล์ Ref_number.xlsx:").grid(
            row=1, column=0, sticky="w", pady=4,
        )
        self.etk_req_excel_input = tk.StringVar(value="")
        ttk.Entry(self.etk_req_no_frame, textvariable=self.etk_req_excel_input).grid(
            row=1, column=1, sticky="we", padx=8,
        )
        etk_req_excel_btns = ttk.Frame(self.etk_req_no_frame)
        etk_req_excel_btns.grid(row=1, column=2, padx=4)
        ttk.Button(
            etk_req_excel_btns, text="เลือก...",
            command=lambda: self._choose_into(self.etk_req_excel_input),
        ).pack(side="left")
        ttk.Button(
            etk_req_excel_btns, text="ล้าง",
            command=lambda: self.etk_req_excel_input.set(""),
        ).pack(side="left", padx=(4, 0))
        ttk.Label(
            self.etk_req_no_frame,
            text="คอลัมน์ที่ต้องมี: Ref_number (เลขคำขอ), user (Username เจ้าของคำขอ)",
            foreground="gray",
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(0, 6))
        ttk.Label(self.etk_req_no_frame, text="ไฟล์ UsernameLogin.xlsx:").grid(
            row=3, column=0, sticky="w", pady=4,
        )
        self.etk_req_login_input = tk.StringVar(value=str(ROOT / "UsernameLogin.xlsx"))
        ttk.Entry(self.etk_req_no_frame, textvariable=self.etk_req_login_input).grid(
            row=3, column=1, sticky="we", padx=8,
        )
        ttk.Button(
            self.etk_req_no_frame, text="เลือก...",
            command=lambda: self._choose_into(self.etk_req_login_input),
        ).grid(row=3, column=2, padx=4)
        ttk.Label(
            self.etk_req_no_frame,
            text="ต้องมีคอลัมน์: Username, Password, Type — ใช้ตอนมีไฟล์ Ref_number.xlsx เท่านั้น",
            foreground="gray",
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(0, 8))

        ttk.Separator(self.etk_req_no_frame, orient="horizontal").grid(
            row=5, column=0, columnspan=3, sticky="we", pady=(0, 8),
        )
        ttk.Label(
            self.etk_req_no_frame,
            text="โหมด B: พิมพ์/วางเลขคำขอเอง (ใช้บัญชีเดียวจาก Username/Password ด้านบน — ใช้เมื่อไม่ระบุไฟล์ Excel โหมด A)",
            foreground="#444",
        ).grid(row=6, column=0, columnspan=3, sticky="w", pady=(0, 2))
        self.etk_req_no_text = tk.Text(self.etk_req_no_frame, height=4, width=60)
        self.etk_req_no_text.grid(row=7, column=0, columnspan=3, sticky="we")
        self.etk_req_no_frame.columnconfigure(1, weight=1)

        # โหลด default ตาม request_type แรก
        self._apply_profile_defaults(self.request_type.get())
        self._update_request_summary()

        # ---- ตัวเลือก (โหมด aliens) ----
        self.aliens_frame = ttk.LabelFrame(
            self._body, text="ตัวเลือก — ข้อมูลคนต่างด้าว", padding=10,
        )
        ttk.Label(
            self.aliens_frame, text="เลือกหมวดข้อมูลที่จะดึง:", foreground="#444",
        ).pack(anchor="w")
        self.alien_subtab_vars: dict[str, tk.BooleanVar] = {}
        for name in ALIEN_SUB_TABS:
            v = tk.BooleanVar(value=True)
            self.alien_subtab_vars[name] = v
            ttk.Checkbutton(self.aliens_frame, text=name, variable=v).pack(
                anchor="w", padx=8, pady=2,
            )
        ttk.Label(
            self.aliens_frame,
            text="(ดึงเฉพาะข้อมูลในตาราง ไม่กดเข้าไปดูรายละเอียดแต่ละคน — เร็วมาก)",
            foreground="gray",
        ).pack(anchor="w", pady=(6, 0))

        # ---- ตัวเลือก (โหมด register) ----
        self.register_frame = ttk.LabelFrame(
            self._body, text="ตัวเลือก — ลงทะเบียนคนต่างด้าว", padding=10,
        )
        ttk.Label(
            self.register_frame,
            text="ไฟล์ Excel ต้องมีคอลัมน์: No., Name, TaxID, PassportNo., Password, Email",
            foreground="#444",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))
        ttk.Label(self.register_frame, text="ไฟล์ Excel input:").grid(row=1, column=0, sticky="w", pady=4)
        self.register_input = tk.StringVar(value=str(ROOT / "Exm.xlsx"))
        ttk.Entry(self.register_frame, textvariable=self.register_input).grid(
            row=1, column=1, sticky="we", padx=8,
        )
        ttk.Button(
            self.register_frame, text="เลือก...", command=self._choose_register_input,
        ).grid(row=1, column=2, padx=4)
        ttk.Label(self.register_frame, text="เลือกแถวที่จะลงทะเบียน:").grid(
            row=2, column=0, sticky="w", pady=4,
        )
        self.register_row_range = tk.StringVar(value="")
        ttk.Entry(self.register_frame, textvariable=self.register_row_range).grid(
            row=2, column=1, sticky="we", padx=8,
        )
        ttk.Label(
            self.register_frame,
            text='เช่น "4-13" หรือ "1,3,5-8" (เว้นว่าง = ทุกแถว)',
            foreground="gray",
        ).grid(row=2, column=2, sticky="w")
        ttk.Label(
            self.register_frame,
            text='ตัวเลขคือลำดับแถวข้อมูลใน Excel (ไม่นับหัวตาราง) — แถวแรก = 1',
            foreground="gray",
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(2, 0))
        ttk.Label(
            self.register_frame,
            text="(ใช้บัญชีของคนต่างด้าวแต่ละคนจาก Excel — ไม่ต้องใส่ Username/Password ด้านบน)",
            foreground="gray",
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(8, 0))
        self.register_frame.columnconfigure(1, weight=1)

        # ---- ตัวเลือก (โหมด receipts) ----
        self.receipt_frame = ttk.LabelFrame(
            self._body, text="ตัวเลือก — ดาวน์โหลดเอกสาร (ใบเสร็จทุกราคา / บต.44 / บต.22 / บต.53 / บต.56)", padding=10,
        )
        ttk.Label(
            self.receipt_frame,
            text="RequestData ต้องมีคอลัมน์: ลำดับ, ชื่อคนต่างด้าว(Eng), เลขที่คำขอ, Username, PASSPORT NUMBER",
            foreground="#444",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 2))
        ttk.Label(
            self.receipt_frame,
            text="UsernameLogin ต้องมีคอลัมน์: Username, Password, Type",
            foreground="#444",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 6))

        ttk.Label(self.receipt_frame, text="ไฟล์ RequestData.xlsx:").grid(
            row=2, column=0, sticky="w", pady=4,
        )
        self.receipt_request_input = tk.StringVar(value=str(ROOT / "RequestData.xlsx"))
        ttk.Entry(self.receipt_frame, textvariable=self.receipt_request_input).grid(
            row=2, column=1, sticky="we", padx=8,
        )
        ttk.Button(
            self.receipt_frame, text="เลือก...",
            command=lambda: self._choose_into(self.receipt_request_input),
        ).grid(row=2, column=2, padx=4)

        ttk.Label(self.receipt_frame, text="ไฟล์ UsernameLogin.xlsx:").grid(
            row=3, column=0, sticky="w", pady=4,
        )
        self.receipt_login_input = tk.StringVar(value=str(ROOT / "UsernameLogin.xlsx"))
        ttk.Entry(self.receipt_frame, textvariable=self.receipt_login_input).grid(
            row=3, column=1, sticky="we", padx=8,
        )
        ttk.Button(
            self.receipt_frame, text="เลือก...",
            command=lambda: self._choose_into(self.receipt_login_input),
        ).grid(row=3, column=2, padx=4)

        ttk.Label(self.receipt_frame, text="เลือกแถวที่จะดาวน์โหลด:").grid(
            row=4, column=0, sticky="w", pady=4,
        )
        self.receipt_row_range = tk.StringVar(value="")
        ttk.Entry(self.receipt_frame, textvariable=self.receipt_row_range).grid(
            row=4, column=1, sticky="we", padx=8,
        )
        ttk.Label(
            self.receipt_frame,
            text='เช่น "1-50" หรือ "2,5,8" (เว้นว่าง = ทุกแถว)',
            foreground="gray",
        ).grid(row=4, column=2, sticky="w")

        # เลือกประเภทเอกสารที่จะดาวน์โหลด
        ttk.Label(self.receipt_frame, text="เอกสารที่จะดาวน์โหลด:").grid(
            row=5, column=0, sticky="w", pady=(6, 2),
        )
        doc_box = ttk.Frame(self.receipt_frame)
        doc_box.grid(row=5, column=1, columnspan=2, sticky="w", pady=(6, 2))
        ttk.Checkbutton(
            doc_box, text="ใบเสร็จ (ทุกใบ)", variable=self.doc_receipt,
        ).pack(side="left", padx=(0, 14))
        ttk.Checkbutton(
            doc_box, text="แบบ บต.44", variable=self.doc_bt44,
        ).pack(side="left", padx=(0, 14))
        ttk.Checkbutton(
            doc_box, text="แบบ บต.22", variable=self.doc_bt22,
        ).pack(side="left", padx=(0, 14))
        ttk.Checkbutton(
            doc_box, text="แบบ บต.53", variable=self.doc_bt53,
        ).pack(side="left", padx=(0, 14))
        ttk.Checkbutton(
            doc_box, text="แบบ บต.55", variable=self.doc_bt55,
        ).pack(side="left", padx=(0, 14))
        ttk.Checkbutton(
            doc_box, text="แบบ บต.52", variable=self.doc_bt52,
        ).pack(side="left", padx=(0, 14))
        ttk.Checkbutton(
            doc_box, text="แบบ บต.56", variable=self.doc_bt56,
        ).pack(side="left")

        # ต่อท้ายชื่อไฟล์เอง (optional) + สร้างโฟลเดอร์แยกตาม PASSPORT
        self.receipt_name_suffix = tk.StringVar(value="")
        self.receipt_make_folder = tk.BooleanVar(value=False)
        ttk.Label(self.receipt_frame, text="ต่อท้ายชื่อไฟล์ (ถ้าต้องการ):").grid(
            row=6, column=0, sticky="w", pady=(6, 2),
        )
        ttk.Entry(self.receipt_frame, textvariable=self.receipt_name_suffix, width=18).grid(
            row=6, column=1, sticky="w", padx=8, pady=(6, 2),
        )
        ttk.Label(
            self.receipt_frame,
            text='เว้นว่าง = ไม่ต่อท้าย | ใส่ "_IO" → {PASSPORT}_BT22_IO.pdf',
            foreground="gray",
        ).grid(row=6, column=2, sticky="w")
        ttk.Checkbutton(
            self.receipt_frame,
            text="สร้างโฟลเดอร์แยกตาม PASSPORT (เช่น receipts/MD1524123/MD1524123_BT44.pdf)",
            variable=self.receipt_make_folder,
        ).grid(row=7, column=0, columnspan=3, sticky="w", pady=(2, 4))

        ttk.Label(
            self.receipt_frame,
            text="ไฟล์ PDF เก็บในโฟลเดอร์ receipts — ใบเสร็จดาวน์ทุกใบ ตั้งชื่อตามราคา: {PASSPORT}_RECEIPT100.pdf / _RECEIPT1800.pdf / _BT44.pdf / _BT22.pdf / _BT53.pdf / _BT55.pdf / _BT52.pdf / _BT56.pdf",
            foreground="gray",
        ).grid(row=8, column=0, columnspan=3, sticky="w", pady=(2, 0))
        ttk.Label(
            self.receipt_frame,
            text="(จัดกลุ่มตาม Username → login ตาม Type ในไฟล์ — ไม่ต้องกรอก Username/Password ด้านบน)",
            foreground="gray",
        ).grid(row=9, column=0, columnspan=3, sticky="w", pady=(2, 0))
        ttk.Label(
            self.receipt_frame,
            text="ทำซ้ำได้: เอกสารที่มีไฟล์ PDF อยู่แล้วจะถูกข้าม / รองรับ session หมดอายุ",
            foreground="gray",
        ).grid(row=10, column=0, columnspan=3, sticky="w", pady=(2, 0))
        self.receipt_frame.columnconfigure(1, weight=1)

        # ---- ตัวเลือก (โหมด results: ใบแจ้งผล / ใบรับคำขอ) ----
        self.result_frame = ttk.LabelFrame(
            self._body, text="ตัวเลือก — ดาวน์โหลดเอกสารผลอนุญาต (ใบแจ้งผล / ใบรับคำขอ)", padding=10,
        )
        ttk.Label(
            self.result_frame,
            text="โหมด A: ใส่ไฟล์ Ref_number.xlsx → ค้นหา-ดาวน์โหลดตามเลขคำขอ (Ref_number + user)",
            foreground="#444",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 2))
        ttk.Label(
            self.result_frame,
            text="โหมด B: เว้นว่าง/ล้างไฟล์ Ref_number → ดึงจาก e-Tracking ตามรายการคำขอ (สถานะ 'รอนัดหมาย')",
            foreground="#444",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 6))

        # ---- โหมด A: ไฟล์ Ref_number.xlsx (ค้นหาตามเลขคำขอ) ----
        ttk.Label(self.result_frame, text="ไฟล์ Ref_number.xlsx:").grid(
            row=2, column=0, sticky="w", pady=4,
        )
        self.result_ref_input = tk.StringVar(value=str(ROOT / "Ref_number.xlsx"))
        ttk.Entry(self.result_frame, textvariable=self.result_ref_input).grid(
            row=2, column=1, sticky="we", padx=8,
        )
        rref_btns = ttk.Frame(self.result_frame)
        rref_btns.grid(row=2, column=2, padx=4)
        ttk.Button(
            rref_btns, text="เลือก...",
            command=lambda: self._choose_into(self.result_ref_input),
        ).pack(side="left")
        ttk.Button(
            rref_btns, text="ล้าง",
            command=lambda: self.result_ref_input.set(""),
        ).pack(side="left", padx=(4, 0))
        ttk.Label(
            self.result_frame,
            text="คอลัมน์ที่ต้องมี: Ref_number (เลขคำขอ), user (Username เจ้าของคำขอ)",
            foreground="gray",
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(0, 6))

        ttk.Label(self.result_frame, text="ไฟล์ UsernameLogin.xlsx:").grid(
            row=4, column=0, sticky="w", pady=4,
        )
        self.result_login_input = tk.StringVar(value=str(ROOT / "UsernameLogin.xlsx"))
        ttk.Entry(self.result_frame, textvariable=self.result_login_input).grid(
            row=4, column=1, sticky="we", padx=8,
        )
        ttk.Button(
            self.result_frame, text="เลือก...",
            command=lambda: self._choose_into(self.result_login_input),
        ).grid(row=4, column=2, padx=4)
        ttk.Label(
            self.result_frame,
            text="UsernameLogin ต้องมีคอลัมน์: Username, Password, Type (จัดกลุ่มตาม Username → login อัตโนมัติ)",
            foreground="gray",
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(0, 6))

        # ---- โหมด B: รายการคำขอ (ใช้เมื่อไม่มีไฟล์ Ref_number) ----
        ttk.Label(self.result_frame, text="รายการคำขอ (โหมด B):").grid(row=6, column=0, sticky="w", pady=4)
        self.result_request_type = tk.StringVar(value="")
        self.result_request_label = tk.StringVar(value=REQUEST_TYPES[0][1])
        rr_combo = ttk.Combobox(
            self.result_frame, textvariable=self.result_request_label, state="readonly",
            values=[lbl for _, lbl in REQUEST_TYPES],
        )
        rr_combo.grid(row=6, column=1, columnspan=2, sticky="we", padx=8)
        rr_combo.bind("<<ComboboxSelected>>", self._on_result_request_changed)

        ttk.Label(self.result_frame, text="หรือใส่ code เอง:").grid(row=7, column=0, sticky="w", pady=4)
        ttk.Entry(self.result_frame, textvariable=self.result_request_type).grid(
            row=7, column=1, sticky="we", padx=8,
        )
        ttk.Label(self.result_frame, text="(เว้นว่าง = ทั้งหมด)", foreground="gray").grid(
            row=7, column=2, sticky="w",
        )

        ttk.Label(self.result_frame, text="เอกสารที่จะดาวน์โหลด:").grid(
            row=8, column=0, sticky="w", pady=(6, 2),
        )
        rdoc_box = ttk.Frame(self.result_frame)
        rdoc_box.grid(row=8, column=1, columnspan=2, sticky="w", pady=(6, 2))
        ttk.Checkbutton(
            rdoc_box, text="ใบแจ้งผลใบอนุญาตทำงาน", variable=self.doc_result_notice,
        ).pack(side="left", padx=(0, 14))
        ttk.Checkbutton(
            rdoc_box, text="ใบรับคำขอใบอนุญาตทำงาน", variable=self.doc_request_receipt,
        ).pack(side="left", padx=(0, 14))
        ttk.Checkbutton(
            rdoc_box, text="แบบ บต.50 อ.6", variable=self.doc_result_bt50,
        ).pack(side="left")

        ttk.Label(
            self.result_frame,
            text="ไฟล์ PDF เก็บในโฟลเดอร์ result_docs — ตั้งชื่อ {ชื่อคนต่างด้าว(Eng)}_{ชื่อเอกสารบนเว็บ}.pdf",
            foreground="gray",
        ).grid(row=9, column=0, columnspan=3, sticky="w", pady=(2, 0))
        ttk.Label(
            self.result_frame,
            text="ทำซ้ำได้: เอกสารที่มีไฟล์ PDF อยู่แล้วจะถูกข้าม / รองรับ session หมดอายุ",
            foreground="gray",
        ).grid(row=10, column=0, columnspan=3, sticky="w", pady=(2, 0))
        self.result_frame.columnconfigure(1, weight=1)

        # ---- ตัวเลือก (โหมด inform: แจ้งเข้านายจ้าง บต.52) ----
        self.inform_frame = ttk.LabelFrame(
            self._body, text="ตัวเลือก — แจ้งเข้านายจ้าง (แบบ บต.52)", padding=10,
        )
        ttk.Label(
            self.inform_frame,
            text="FormRequestEmployment ต้องมีคอลัมน์: ประเภทนายจ้าง, รหัสนายจ้าง, เลขใบอนุญาตทำงาน, วันที่เริ่มงาน, Files",
            foreground="#444",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 2))
        ttk.Label(
            self.inform_frame,
            text="UsernameLogin ใช้บัญชีแรกที่เจอ (ควรเป็นบัญชี บนจ. ที่ตรงกับฟอร์มนี้)",
            foreground="#444",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 6))

        ttk.Label(self.inform_frame, text="ไฟล์ FormRequestEmployment.xlsx:").grid(
            row=2, column=0, sticky="w", pady=4,
        )
        self.inform_excel_input = tk.StringVar(value=str(ROOT / "FormRequestEmployment.xlsx"))
        ttk.Entry(self.inform_frame, textvariable=self.inform_excel_input).grid(
            row=2, column=1, sticky="we", padx=8,
        )
        ttk.Button(
            self.inform_frame, text="เลือก...",
            command=lambda: self._choose_into(self.inform_excel_input),
        ).grid(row=2, column=2, padx=4)

        ttk.Label(self.inform_frame, text="ไฟล์ UsernameLogin.xlsx:").grid(
            row=3, column=0, sticky="w", pady=4,
        )
        self.inform_login_input = tk.StringVar(value=str(ROOT / "UsernameLogin.xlsx"))
        ttk.Entry(self.inform_frame, textvariable=self.inform_login_input).grid(
            row=3, column=1, sticky="we", padx=8,
        )
        ttk.Button(
            self.inform_frame, text="เลือก...",
            command=lambda: self._choose_into(self.inform_login_input),
        ).grid(row=3, column=2, padx=4)

        ttk.Label(self.inform_frame, text="เลือกแถวที่จะทำ:").grid(
            row=4, column=0, sticky="w", pady=4,
        )
        self.inform_row_range = tk.StringVar(value="")
        ttk.Entry(self.inform_frame, textvariable=self.inform_row_range).grid(
            row=4, column=1, sticky="we", padx=8,
        )
        ttk.Label(
            self.inform_frame,
            text='เช่น "1-5" หรือ "2,4" (เว้นว่าง = ทุกแถว)',
            foreground="gray",
        ).grid(row=4, column=2, sticky="w")

        self.inform_commit = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            self.inform_frame,
            text="ยืนยันส่งคำขอจริง (ทำข้อ 7+8) — ถ้าไม่ติ๊กจะหยุดก่อนข้อ 7 (โหมดทดสอบ)",
            variable=self.inform_commit,
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(8, 2))
        ttk.Label(
            self.inform_frame,
            text="โหมดทดสอบ: เก็บ screenshot หน้าสรุปคำขอใน inform_screenshots/ — ไม่ส่งจริง",
            foreground="#a04a00",
        ).grid(row=6, column=0, columnspan=3, sticky="w", pady=(2, 0))
        ttk.Label(
            self.inform_frame,
            text="ไฟล์แนบ (คอลัมน์ Files) ใช้ relative path จากโฟลเดอร์เดียวกับ Excel",
            foreground="gray",
        ).grid(row=7, column=0, columnspan=3, sticky="w", pady=(2, 0))
        self.inform_frame.columnconfigure(1, weight=1)

        # ---- ตัวเลือก (โหมด bt30: ยื่นต่อใบอนุญาตทำงานตาม MoU แบบ บต.30) ----
        self.bt30_frame = ttk.LabelFrame(
            self._body, text="ตัวเลือก — ยื่นต่อใบอนุญาตทำงาน (แบบ บต.30 / MoU)", padding=10,
        )
        ttk.Label(
            self.bt30_frame,
            text="from_bt30.xlsx ต้องมีคอลัมน์: No., คำนำหน้า, ชื่อ, สัญชาติ, เพศ, วันเกิด",
            foreground="#444",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 2))
        ttk.Label(
            self.bt30_frame,
            text="ขั้นตอน 1: Login → เมนูบริการ → ต่ออายุตาม MoU → 'ค้นหาข้อมูลคนต่างด้าว' → บันทึก",
            foreground="#444",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 6))

        ttk.Label(self.bt30_frame, text="ไฟล์ from_bt30.xlsx:").grid(
            row=2, column=0, sticky="w", pady=4,
        )
        self.bt30_excel_input = tk.StringVar(value=str(ROOT / "from_bt30.xlsx"))
        ttk.Entry(self.bt30_frame, textvariable=self.bt30_excel_input).grid(
            row=2, column=1, sticky="we", padx=8,
        )
        ttk.Button(
            self.bt30_frame, text="เลือก...",
            command=lambda: self._choose_into(self.bt30_excel_input),
        ).grid(row=2, column=2, padx=4)

        ttk.Label(self.bt30_frame, text="ไฟล์ UsernameLogin.xlsx:").grid(
            row=3, column=0, sticky="w", pady=4,
        )
        self.bt30_login_input = tk.StringVar(value=str(ROOT / "UsernameLogin.xlsx"))
        ttk.Entry(self.bt30_frame, textvariable=self.bt30_login_input).grid(
            row=3, column=1, sticky="we", padx=8,
        )
        ttk.Button(
            self.bt30_frame, text="เลือก...",
            command=lambda: self._choose_into(self.bt30_login_input),
        ).grid(row=3, column=2, padx=4)

        ttk.Label(self.bt30_frame, text="เลือกแถวที่จะทำ:").grid(
            row=4, column=0, sticky="w", pady=4,
        )
        self.bt30_row_range = tk.StringVar(value="")
        ttk.Entry(self.bt30_frame, textvariable=self.bt30_row_range).grid(
            row=4, column=1, sticky="we", padx=8,
        )
        ttk.Label(
            self.bt30_frame,
            text='เช่น "1-5" หรือ "2,4" (เว้นว่าง = ทุกแถว)',
            foreground="gray",
        ).grid(row=4, column=2, sticky="w")
        self.bt30_do_step2 = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            self.bt30_frame,
            text="ทำขั้นตอนที่ 2 ต่อ (2.1–2.9 → Step 4 เอกสารนายจ้าง → 5.1/5.2 สรุปคำขอ → 6.1 ยืนยันตัวตน)",
            variable=self.bt30_do_step2,
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ttk.Label(
            self.bt30_frame,
            text="โหมดขั้นตอน 1+2 = 1 แถวต่อ 1 คำขอ (เปิดฟอร์มใหม่ทุกคน) และหยุดหลัง Step 6.1 (ยืนยันตัวตน) — ไม่ส่งคำขอ/ไม่ชำระเงิน",
            foreground="#a04a00",
        ).grid(row=6, column=0, columnspan=3, sticky="w", pady=(4, 0))
        ttk.Label(
            self.bt30_frame,
            text="วันที่ในไฟล์ใช้รูปแบบ ค.ศ. (เช่น 01/01/1995) — เก็บ screenshot ไว้ใน reports/bt30_screenshots/",
            foreground="#a04a00",
        ).grid(row=7, column=0, columnspan=3, sticky="w", pady=(2, 0))
        ttk.Separator(self.bt30_frame, orient="horizontal").grid(
            row=8, column=0, columnspan=3, sticky="we", pady=(8, 4),
        )
        self.bt30_submit = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            self.bt30_frame,
            text="⚠ ส่งคำขอจริง + ชำระเงิน (Step 7) แล้วเก็บข้อมูล+ดาวน์โหลดใบชำระเงิน (Step 8)",
            variable=self.bt30_submit,
        ).grid(row=9, column=0, columnspan=3, sticky="w", pady=(0, 0))
        ttk.Label(
            self.bt30_frame,
            text="*** อันตราย: กด 'ถัดไป' หน้าชำระเงิน = ยืนยันส่งคำขอจริง ย้อนกลับไม่ได้ และมีค่าธรรมเนียม ***",
            foreground="#c00000", font=("Segoe UI", 9, "bold"),
        ).grid(row=10, column=0, columnspan=3, sticky="w", pady=(2, 0))
        ttk.Label(
            self.bt30_frame,
            text="ต้องทำขั้นตอนที่ 2 ก่อน • ใบชำระเงินดาวน์โหลดได้ครั้งเดียว → เก็บใน reports/bt30_submitted/",
            foreground="#c00000",
        ).grid(row=11, column=0, columnspan=3, sticky="w", pady=(2, 0))
        self.bt30_frame.columnconfigure(1, weight=1)

        # ---- ตัวเลือก (โหมด bt44: แจ้งการทำงาน/เปลี่ยนรายการในใบอนุญาต ซึ่งไม่กระทบ แบบ บต.44) ----
        self.bt44_frame = ttk.LabelFrame(
            self._body, text="ตัวเลือก — เปลี่ยนย้ายนายจ้างในระบบ (บต.44)", padding=10,
        )
        ttk.Label(
            self.bt44_frame,
            text="from_bt44 ต้องมีคอลัมน์: No., คำนำหน้า, หมายเลขอ้างอิงของคนต่างด้าว, ชื่อ, สัญชาติ, เพศ, เกิดวันที่",
            foreground="#444",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 2))
        ttk.Label(
            self.bt44_frame,
            text="ขั้นตอน: Login → เมนูบริการ → การยื่นขอเปลี่ยนรายการฯ ซึ่งไม่กระทบ → 'ค้นหาข้อมูลคนต่างด้าว' → บันทึก",
            foreground="#444",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 6))

        ttk.Label(self.bt44_frame, text="ไฟล์ from_bt44:").grid(
            row=2, column=0, sticky="w", pady=4,
        )
        self.bt44_excel_input = tk.StringVar(value=str(ROOT / "from_bt44.xlxs.xlsx"))
        ttk.Entry(self.bt44_frame, textvariable=self.bt44_excel_input).grid(
            row=2, column=1, sticky="we", padx=8,
        )
        ttk.Button(
            self.bt44_frame, text="เลือก...",
            command=lambda: self._choose_into(self.bt44_excel_input),
        ).grid(row=2, column=2, padx=4)

        ttk.Label(self.bt44_frame, text="ไฟล์ UsernameLogin.xlsx:").grid(
            row=3, column=0, sticky="w", pady=4,
        )
        self.bt44_login_input = tk.StringVar(value=str(ROOT / "UsernameLogin.xlsx"))
        ttk.Entry(self.bt44_frame, textvariable=self.bt44_login_input).grid(
            row=3, column=1, sticky="we", padx=8,
        )
        ttk.Button(
            self.bt44_frame, text="เลือก...",
            command=lambda: self._choose_into(self.bt44_login_input),
        ).grid(row=3, column=2, padx=4)

        ttk.Label(self.bt44_frame, text="เลือกแถวที่จะทำ:").grid(
            row=4, column=0, sticky="w", pady=4,
        )
        self.bt44_row_range = tk.StringVar(value="")
        ttk.Entry(self.bt44_frame, textvariable=self.bt44_row_range).grid(
            row=4, column=1, sticky="we", padx=8,
        )
        ttk.Label(
            self.bt44_frame,
            text='เช่น "1-5" หรือ "2,4" (เว้นว่าง = ทุกแถว)',
            foreground="gray",
        ).grid(row=4, column=2, sticky="w")
        ttk.Label(
            self.bt44_frame,
            text="หมายเหตุ: หมายเลขอ้างอิงของคนต่างด้าว หากไม่มีระบุใน Excel ก็ไม่ต้องกรอก",
            foreground="#444",
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ttk.Label(
            self.bt44_frame,
            text="หยุดหลังกด 'บันทึก' (ไม่กดถัดไป/ไม่ส่งคำขอ) — ถ้า Modal แจ้ง Error จะเก็บข้อความไว้ในรายงานแล้วข้ามไปคนถัดไป",
            foreground="#a04a00",
        ).grid(row=6, column=0, columnspan=3, sticky="w", pady=(2, 0))

        # 'ทดลองยื่น' (Dry-Run) — หยุดที่ Step 2.3 + เก็บข้อมูลคนต่างด้าวเต็ม + ไม่ส่งคำขอ
        self.bt44_dry_run = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            self.bt44_frame,
            text="🧪 ทดลองยื่น (Dry-Run) — ไม่ส่งคำขอจริง + เก็บข้อมูลคนต่างด้าวเต็มลงรายงาน",
            variable=self.bt44_dry_run,
            command=self._on_bt44_dryrun_toggle,
        ).grid(row=7, column=0, columnspan=3, sticky="w", pady=(8, 0))

        # ระดับการหยุด — เลือกได้เมื่อ Dry-Run เปิด
        self.bt44_dry_stop = tk.StringVar(value="step2")
        self._bt44_dry_stop_frame = ttk.Frame(self.bt44_frame)
        self._bt44_dry_stop_frame.grid(row=8, column=0, columnspan=3, sticky="w", padx=(24, 0), pady=(2, 0))
        ttk.Label(self._bt44_dry_stop_frame, text="หยุดที่:").pack(side="left")
        self._bt44_dry_stop_rb2 = ttk.Radiobutton(
            self._bt44_dry_stop_frame,
            text="Step 2 (ตรวจเลขใบอนุญาต) — ไม่แก้ที่อยู่",
            variable=self.bt44_dry_stop, value="step2",
        )
        self._bt44_dry_stop_rb2.pack(side="left", padx=(6, 12))
        self._bt44_dry_stop_rb3 = ttk.Radiobutton(
            self._bt44_dry_stop_frame,
            text="Step 3 (เลือกนายจ้าง+ประเภทกิจการ+งาน) — รวมแก้ที่อยู่",
            variable=self.bt44_dry_stop, value="step3",
        )
        self._bt44_dry_stop_rb3.pack(side="left")
        ttk.Label(
            self.bt44_frame,
            text="• Step 2 = หยุดหลังตรวจเลขใบอนุญาต (Step 2.3) เพื่อเช็คข้อมูลคนต่างด้าวอย่างเดียว\n"
                 "• Step 3 = ทำ Step 2.2 แก้ที่อยู่ + Step 3.1 เปลี่ยนนายจ้าง + Step 3.3-3.4 เลือกประเภทกิจการ/งาน "
                 "แล้วหยุดก่อนแนบเอกสาร (Step 4)",
            foreground="#1a7000", justify="left",
        ).grid(row=9, column=0, columnspan=3, sticky="w", pady=(2, 0))
        # ปิด radio buttons ตอน dry-run ยังไม่เปิด
        self._on_bt44_dryrun_toggle()

        # ตรวจไฟล์แนบ Step 4 ก่อนรัน (ค่าเริ่มต้น: ON)
        self.bt44_check_docs = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            self.bt44_frame,
            text="📎 ตรวจไฟล์แนบ Step 4 ก่อนรัน (ถ้าเอกสารไม่ครบ/เปิดไม่ได้ → ข้าม record นั้น)",
            variable=self.bt44_check_docs,
        ).grid(row=10, column=0, columnspan=3, sticky="w", pady=(8, 0))
        ttk.Label(
            self.bt44_frame,
            text="ปิดได้เมื่อยังไม่ต้องการตรวจไฟล์ (เช่น ระหว่างทดสอบ flow Dry-Run) — ในการใช้งานจริงควรเปิดไว้",
            foreground="#a04a00",
        ).grid(row=11, column=0, columnspan=3, sticky="w", pady=(2, 0))
        self.bt44_frame.columnconfigure(1, weight=1)

        # ----- โหมด: ดาวน์โหลดใบแจ้งชำระเงิน (รอจ่ายค่าธรรมเนียม) -----
        self.billpay_frame = ttk.LabelFrame(
            self._body, text="ตัวเลือก — ดาวน์โหลดใบแจ้งชำระเงิน (รอจ่ายค่าธรรมเนียม)", padding=10,
        )
        ttk.Label(
            self.billpay_frame,
            text="ขั้นตอน: Login → Tracking → กรองสถานะ 'รอชำระเงิน / รอจ่ายเงินค่าธรรมเนียม' → เปิดคำขอ → แท็บ 'การชำระเงิน' → พิมพ์แบบฟอร์มการชำระเงิน",
            foreground="#444",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 2))
        ttk.Label(
            self.billpay_frame,
            text="รองรับหลายบัญชี — วนทุก Username/Password/Type ในไฟล์ UsernameLogin.xlsx",
            foreground="#444",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 6))

        ttk.Label(self.billpay_frame, text="ไฟล์ UsernameLogin.xlsx:").grid(
            row=2, column=0, sticky="w", pady=4,
        )
        self.billpay_login_input = tk.StringVar(value=str(ROOT / "UsernameLogin.xlsx"))
        ttk.Entry(self.billpay_frame, textvariable=self.billpay_login_input).grid(
            row=2, column=1, sticky="we", padx=8,
        )
        ttk.Button(
            self.billpay_frame, text="เลือก...",
            command=lambda: self._choose_into(self.billpay_login_input),
        ).grid(row=2, column=2, padx=4)

        ttk.Label(self.billpay_frame, text="เลือกแถวที่จะดาวน์โหลด:").grid(
            row=3, column=0, sticky="w", pady=4,
        )
        self.billpay_row_range = tk.StringVar(value="")
        ttk.Entry(self.billpay_frame, textvariable=self.billpay_row_range).grid(
            row=3, column=1, sticky="we", padx=8,
        )
        ttk.Label(
            self.billpay_frame,
            text='เช่น "1-5" หรือ "2,4" (เว้นว่าง = ทุกแถวในแต่ละบัญชี)',
            foreground="gray",
        ).grid(row=3, column=2, sticky="w")

        # รายการคำขอ (เลือกได้หลายรายการ) — เหมือนโหมด e-Tracking แต่ติ๊กได้หลายอัน
        ttk.Label(self.billpay_frame, text="รายการคำขอ:").grid(
            row=4, column=0, sticky="nw", pady=4,
        )
        bp_lb_box = ttk.Frame(self.billpay_frame)
        bp_lb_box.grid(row=4, column=1, columnspan=2, sticky="we", padx=8, pady=4)
        bp_sb = ttk.Scrollbar(bp_lb_box, orient="vertical")
        # selectmode=multiple → คลิกเลือก/ยกเลิกได้ทีละอัน ไม่ต้องกด Ctrl
        self.billpay_reqtypes_lb = tk.Listbox(
            bp_lb_box, selectmode="multiple", height=7,
            exportselection=False, yscrollcommand=bp_sb.set, activestyle="none",
        )
        bp_sb.config(command=self.billpay_reqtypes_lb.yview)
        bp_sb.pack(side="right", fill="y")
        self.billpay_reqtypes_lb.pack(side="left", fill="both", expand=True)
        # เก็บเฉพาะรายการที่มีรหัส (ตัด "" = ทั้งหมด ออก) — ไม่เลือก = ทั้งหมด
        self._billpay_reqtype_codes = [code for code, _ in REQUEST_TYPES if code]
        for code in self._billpay_reqtype_codes:
            lbl = next((l for c, l in REQUEST_TYPES if c == code), code)
            self.billpay_reqtypes_lb.insert("end", lbl)
        ttk.Label(
            self.billpay_frame,
            text="คลิกเลือกได้หลายรายการ (คลิกซ้ำ = ยกเลิก) — ไม่เลือก = ทุกรายการคำขอ",
            foreground="gray",
        ).grid(row=5, column=1, columnspan=2, sticky="w", padx=8)

        ttk.Label(
            self.billpay_frame,
            text="ปลอดภัย: เป็นการสร้าง 'ใบแจ้งชำระเงิน' เพื่อพิมพ์เก็บไว้ (จ่ายภายหลังผ่านธนาคาร) — ไม่มีการตัดเงิน",
            foreground="#1a7000",
        ).grid(row=6, column=0, columnspan=3, sticky="w", pady=(4, 0))
        ttk.Label(
            self.billpay_frame,
            text="ตั้งชื่อไฟล์: {เลขที่คำขอ}_{รายชื่อคนต่างด้าว} → reports/bill_payment/",
            foreground="#444",
        ).grid(row=7, column=0, columnspan=3, sticky="w", pady=(2, 0))
        self.billpay_frame.columnconfigure(1, weight=1)

        # ----- โหมด: ดาวน์โหลดใบเสร็จรับเงินทั้งชุด -----
        self.payrcpt_frame = ttk.LabelFrame(
            self._body, text="ตัวเลือก — ดาวน์โหลดใบเสร็จรับเงินทั้งชุด (900 + 100 + อื่นๆ)",
            padding=10,
        )
        ttk.Label(
            self.payrcpt_frame,
            text="ขั้นตอน: อ่านเลขคำขอจาก Excel → Login ต่อ Username → ค้นหาคำขอ → แท็บ 'การชำระเงิน' → 'ดูใบเสร็จรับเงิน' → วนทุกแถว ทุกหน้า pagination → ดาวน์โหลด → แยก PDF ผลายหน้าเป็น 1 ใบ/1 คน/1 ไฟล์",
            foreground="#444", wraplength=900, justify="left",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))

        ttk.Label(self.payrcpt_frame, text="ไฟล์ Request_Receipt.xlsx:").grid(
            row=1, column=0, sticky="w", pady=4,
        )
        self.payrcpt_request_input = tk.StringVar(value=str(ROOT / "Request_Receipt.xlsx"))
        ttk.Entry(self.payrcpt_frame, textvariable=self.payrcpt_request_input).grid(
            row=1, column=1, sticky="we", padx=8,
        )
        ttk.Button(
            self.payrcpt_frame, text="เลือก...",
            command=lambda: self._choose_into(self.payrcpt_request_input),
        ).grid(row=1, column=2, padx=4)
        ttk.Label(
            self.payrcpt_frame,
            text="คอลัมน์: ลำดับ, เลขที่คำขอ, Username",
            foreground="gray",
        ).grid(row=2, column=1, columnspan=2, sticky="w", padx=8)

        ttk.Label(self.payrcpt_frame, text="ไฟล์ UsernameLogin.xlsx:").grid(
            row=3, column=0, sticky="w", pady=4,
        )
        self.payrcpt_login_input = tk.StringVar(value=str(ROOT / "UsernameLogin.xlsx"))
        ttk.Entry(self.payrcpt_frame, textvariable=self.payrcpt_login_input).grid(
            row=3, column=1, sticky="we", padx=8,
        )
        ttk.Button(
            self.payrcpt_frame, text="เลือก...",
            command=lambda: self._choose_into(self.payrcpt_login_input),
        ).grid(row=3, column=2, padx=4)

        ttk.Label(self.payrcpt_frame, text="เลือกแถวที่จะทำ:").grid(
            row=4, column=0, sticky="w", pady=4,
        )
        self.payrcpt_row_range = tk.StringVar(value="")
        ttk.Entry(self.payrcpt_frame, textvariable=self.payrcpt_row_range).grid(
            row=4, column=1, sticky="we", padx=8,
        )
        ttk.Label(
            self.payrcpt_frame,
            text='เช่น "1-5" หรือ "2,4" (เว้นว่าง = ทุกแถว)',
            foreground="gray",
        ).grid(row=4, column=2, sticky="w")

        ttk.Label(
            self.payrcpt_frame,
            text="ตั้งชื่อไฟล์: {เลขที่คำขอ}_{ชื่อคนต่างด้าว}_{ชื่อนายจ้าง}_{เลขอ้างอิง}_RECEIPT{Amount}.pdf → reports/payment_receipts/",
            foreground="#444", wraplength=900, justify="left",
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ttk.Label(
            self.payrcpt_frame,
            text="ปลอดภัย: เป็นการดาวน์โหลดใบเสร็จที่ชำระเงินไปแล้วเท่านั้น — ไม่มีการส่งคำขอหรือตัดเงิน",
            foreground="#1a7000",
        ).grid(row=6, column=0, columnspan=3, sticky="w", pady=(2, 0))
        self.payrcpt_frame.columnconfigure(1, weight=1)

        # ----- โหมด: นัดหมายถ่ายบัตร -----
        self.appt_frame = ttk.LabelFrame(
            self._body, text="ตัวเลือก — นัดหมายถ่ายบัตร (เก็บที่อยู่จากใบเสร็จค่าธรรมเนียม)",
            padding=10,
        )
        ttk.Label(
            self.appt_frame,
            text="ขั้นตอน: วน Login ทุกบัญชีใน UsernameLogin.xlsx → Tracking → กรองสถานะ 'รอนัดหมาย' (AP) → "
                 "ทุกคำขอ: tab 'ข้อมูลคนต่างด้าว' (เก็บชื่อสถานประกอบการ) → tab 'การชำระเงิน' → "
                 "ดาวน์โหลดใบเสร็จ → อ่านที่อยู่ (จังหวัด + อำเภอ/เขต) จาก PDF",
            foreground="#444", wraplength=900, justify="left",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))

        ttk.Label(self.appt_frame, text="ไฟล์ UsernameLogin.xlsx:").grid(
            row=1, column=0, sticky="w", pady=4,
        )
        self.appt_login_input = tk.StringVar(value=str(ROOT / "UsernameLogin.xlsx"))
        ttk.Entry(self.appt_frame, textvariable=self.appt_login_input).grid(
            row=1, column=1, sticky="we", padx=8,
        )
        ttk.Button(
            self.appt_frame, text="เลือก...",
            command=lambda: self._choose_into(self.appt_login_input),
        ).grid(row=1, column=2, padx=4)

        ttk.Label(self.appt_frame, text="เลือกแถวที่จะทำ (ของแต่ละบัญชี):").grid(
            row=2, column=0, sticky="w", pady=4,
        )
        self.appt_row_range = tk.StringVar(value="")
        ttk.Entry(self.appt_frame, textvariable=self.appt_row_range).grid(
            row=2, column=1, sticky="we", padx=8,
        )
        ttk.Label(
            self.appt_frame,
            text='เช่น "1-5" หรือ "2,4" (เว้นว่าง = ทุกคำขอ AP ของบัญชีนั้น)',
            foreground="gray",
        ).grid(row=2, column=2, sticky="w")

        ttk.Label(
            self.appt_frame,
            text="ปลอดภัย: โหมดอ่านอย่างเดียว — ดาวน์โหลด PDF ใบเสร็จที่ชำระเงินแล้ว ไม่มีการส่งคำขอ/ตัดเงิน",
            foreground="#1a7000",
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ttk.Label(
            self.appt_frame,
            text="เซฟ PDF → reports/appointment_receipts/  |  รายงาน Excel มี hyperlink ไปที่ไฟล์ PDF",
            foreground="#444",
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(2, 0))
        self.appt_frame.columnconfigure(1, weight=1)

        # ---- ตัวเลือก (โหมด bt30_ctn) ----
        self.bt30ctn_frame = ttk.LabelFrame(
            self._body,
            text="ตัวเลือก — ดาวน์โหลด (ใบตอบรับ / ใบนัดหมาย)",
            padding=10,
        )
        ttk.Label(
            self.bt30ctn_frame,
            text="ขั้นตอน: วน Login ทุกบัญชีใน UsernameLogin.xlsx (หรือใช้ Username/Password จาก 'ข้อมูลเข้าสู่ระบบ' ด้านบน) → "
                 "Tracking → กรอง Status ตาม checkbox ด้านบน → ทุกคำขอ: เปิด detail → \n"
                 "  • ใบตอบรับ (บต.30) — tab 'เอกสารตอบรับจากระบบ' → PDF → ตั้งชื่อ {PASSPORT}_BT{XX}_CTN.pdf (สถานะ WP2)\n"
                 "  • แบบ บต.25 — tab 'เอกสารตอบรับจากระบบ' → PDF → ตั้งชื่อ {PASSPORT}_BT25.pdf\n"
                 "  • ใบนัดหมาย — tab 'การนัดหมาย' → iframe queue → PDF → ตั้งชื่อ {PASSPORT}_APPOINTMENT.pdf (สถานะ AP/APSS)\n"
                 "  • Receipt ค่าธรรมเนียมฯ — tab 'การชำระเงิน' → PDF → ตั้งชื่อ {PASSPORT}_RECEIPT.pdf",
            foreground="#444", wraplength=900, justify="left",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))

        ttk.Label(self.bt30ctn_frame, text="ไฟล์ UsernameLogin.xlsx:").grid(
            row=1, column=0, sticky="w", pady=4,
        )
        self.bt30ctn_login_input = tk.StringVar(value=str(ROOT / "UsernameLogin.xlsx"))
        ttk.Entry(self.bt30ctn_frame, textvariable=self.bt30ctn_login_input).grid(
            row=1, column=1, sticky="we", padx=8,
        )
        ttk.Button(
            self.bt30ctn_frame, text="เลือก...",
            command=lambda: self._choose_into(self.bt30ctn_login_input),
        ).grid(row=1, column=2, padx=4)

        ttk.Label(self.bt30ctn_frame, text="เลือกแถวที่จะทำ (ต่อบัญชี):").grid(
            row=2, column=0, sticky="w", pady=4,
        )
        self.bt30ctn_row_range = tk.StringVar(value="")
        ttk.Entry(self.bt30ctn_frame, textvariable=self.bt30ctn_row_range).grid(
            row=2, column=1, sticky="we", padx=8,
        )
        ttk.Label(
            self.bt30ctn_frame,
            text='เช่น "1-10,15" (เว้นว่าง = ทั้งหมด)',
            foreground="gray",
        ).grid(row=2, column=2, sticky="w")

        # date filter (ใช้ตัวแปรร่วมกับโหมด e-Tracking — self.date_from / self.date_to)
        # แสดง label เพื่อบอกว่าจะใช้ค่าจากด้านบน (ถ้ามี)
        ttk.Label(
            self.bt30ctn_frame,
            text="💡 ค่ากรอง 'วันที่ยื่นคำขอ (จาก → ถึง)' + 'รายการคำขอ' + 'สถานะ (checkbox)' ในกล่อง e-Tracking ด้านบนจะถูกใช้ร่วมกัน",
            foreground="#0056b3", wraplength=900, justify="left",
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(6, 2))

        # เลือกประเภทเอกสารที่จะดาวน์โหลด
        ttk.Label(self.bt30ctn_frame, text="เอกสารที่จะดาวน์โหลด:").grid(
            row=4, column=0, sticky="w", pady=(6, 2),
        )
        _doc_box = ttk.Frame(self.bt30ctn_frame)
        _doc_box.grid(row=4, column=1, columnspan=2, sticky="w", pady=(6, 2))
        self.bt30ctn_do_ctn = tk.BooleanVar(value=True)
        self.bt30ctn_do_appointment = tk.BooleanVar(value=False)
        self.bt30ctn_do_bt25 = tk.BooleanVar(value=False)
        self.bt30ctn_do_receipt = tk.BooleanVar(value=False)
        # แถวที่ 1
        _doc_row1 = ttk.Frame(_doc_box)
        _doc_row1.pack(anchor="w")
        ttk.Checkbutton(
            _doc_row1, text="ใบตอบรับ บต.30 (CTN) — สถานะ WP2",
            variable=self.bt30ctn_do_ctn,
        ).pack(side="left", padx=(0, 20))
        ttk.Checkbutton(
            _doc_row1, text="ใบนัดหมาย (APPOINTMENT) — สถานะ AP/APSS",
            variable=self.bt30ctn_do_appointment,
        ).pack(side="left")
        # แถวที่ 2 (เอกสารเพิ่มเติม)
        _doc_row2 = ttk.Frame(_doc_box)
        _doc_row2.pack(anchor="w", pady=(4, 0))
        ttk.Checkbutton(
            _doc_row2, text="แบบ บต.25 (เอกสารตอบรับจากระบบ)",
            variable=self.bt30ctn_do_bt25,
        ).pack(side="left", padx=(0, 20))
        ttk.Checkbutton(
            _doc_row2, text="Receipt ค่าธรรมเนียมใบอนุญาตทำงาน (การชำระเงิน)",
            variable=self.bt30ctn_do_receipt,
        ).pack(side="left")
        # หมายเหตุ: บต.25 ผูกกับรายการคำขอ MT_59 เท่านั้น
        ttk.Label(
            _doc_box,
            text="⚠ 'แบบ บต.25' ต้องกรอง 'รายการคำขอ' เป็น 'การยื่นคำขอรับใบอนุญาตทำงานของ"
                 "คนต่างด้าวซึ่งได้รับอนุญาตให้อยู่ในราชอาณาจักรเป็นการชั่วคราว (MT_59)'",
            foreground="#b06f00", wraplength=680, justify="left",
        ).pack(anchor="w", pady=(4, 0))

        self.bt30ctn_make_subfolder = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            self.bt30ctn_frame,
            text="สร้างโฟลเดอร์แยกตาม PASSPORT (เช่น bt30_ctn/MH788309/MH788309_BT30_CTN.pdf)",
            variable=self.bt30ctn_make_subfolder,
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(6, 2))

        ttk.Label(
            self.bt30ctn_frame,
            text="เซฟ PDF → reports/bt30_ctn/  |  รายงาน Excel มี hyperlink ไปที่ไฟล์ PDF (CTN + บต.25 + APPOINTMENT + Receipt)",
            foreground="#444",
        ).grid(row=6, column=0, columnspan=3, sticky="w", pady=(4, 0))
        ttk.Label(
            self.bt30ctn_frame,
            text="ปลอดภัย: โหมดอ่านอย่างเดียว — ดาวน์โหลด PDF ที่ระบบสร้างแล้ว ไม่มีการส่งคำขอ/ตัดเงิน",
            foreground="#1a7000",
        ).grid(row=7, column=0, columnspan=3, sticky="w", pady=(2, 0))
        self.bt30ctn_frame.columnconfigure(1, weight=1)

        # ---- ตัวเลือก (โหมด namelist_alien) ----
        self.namelist_frame = ttk.LabelFrame(
            self._body,
            text="ตัวเลือก — ดึงรายชื่อคนต่างด้าว (NameListAlien)",
            padding=10,
        )
        ttk.Label(
            self.namelist_frame,
            text="ขั้นตอน: Login ผ่านช่อง 'ข้อมูลเข้าสู่ระบบ' ด้านบน (บัญชีเดียว) → "
                 "เปิด /Requtst63_2/NameListAlien?form_type=<code> → "
                 "วน DataTable ทุกหน้า (100 แถว/หน้า) → บันทึกลง Excel",
            foreground="#444", wraplength=900, justify="left",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))

        ttk.Label(self.namelist_frame, text="Form Type:").grid(
            row=1, column=0, sticky="w", pady=4,
        )
        self.namelist_form_type = tk.StringVar(value="MT_63_2_3103_RENEWAL")
        ttk.Entry(self.namelist_frame, textvariable=self.namelist_form_type).grid(
            row=1, column=1, sticky="we", padx=8,
        )
        ttk.Label(
            self.namelist_frame,
            text="เช่น MT_63_2_3103_RENEWAL, MT_63_2_1302_RENEWAL, MT_63_2_AGN_RENEWAL",
            foreground="gray",
        ).grid(row=1, column=2, sticky="w")

        ttk.Label(self.namelist_frame, text="จำกัดจำนวน (0 = ทั้งหมด):").grid(
            row=2, column=0, sticky="w", pady=4,
        )
        self.namelist_limit = tk.IntVar(value=0)
        ttk.Spinbox(
            self.namelist_frame, from_=0, to=99999,
            textvariable=self.namelist_limit, width=10,
        ).grid(row=2, column=1, sticky="w", padx=8)
        ttk.Label(
            self.namelist_frame,
            text="ใส่ 100 เพื่อทดสอบเร็ว (แต่ละหน้ามี 100 แถว)",
            foreground="gray",
        ).grid(row=2, column=2, sticky="w")

        ttk.Label(
            self.namelist_frame,
            text="ปลอดภัย: โหมดอ่านอย่างเดียว — ไม่มีการส่งคำขอ/แก้ข้อมูลใดๆ",
            foreground="#1a7000",
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(6, 0))
        self.namelist_frame.columnconfigure(1, weight=1)

        # ----- โหมด: ตรวจวันว่างจอง (Real-Time Booking Availability) -----
        self.booking_frame = ttk.LabelFrame(
            self._body,
            text="ตัวเลือก — ตรวจวันว่างจอง (คิวถ่ายบัตร)",
            padding=10,
        )
        ttk.Label(
            self.booking_frame,
            text="ขั้นตอน: Login (ผ่านช่อง 'ข้อมูลเข้าสู่ระบบ' ด้านบน หรือ UsernameLogin.xlsx) → "
                 "หา Bearer token จาก 1 record ที่มีสถานะ AP/SS → "
                 "ดึงรายชื่อสาขาทั้งหมด → วน calendar × time-slot ทุกวัน\n"
                 "รายงาน Excel 3 sheet: ตารางว่างจอง (รายวัน) / สรุปตามสาขา / ช่วงเวลาที่ว่าง (รายรอบ 30 นาที)",
            foreground="#444", wraplength=900, justify="left",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))

        ttk.Label(
            self.booking_frame,
            text="ไฟล์ UsernameLogin.xlsx (ไม่จำเป็น — ปล่อยว่างจะใช้บัญชีจาก 'ข้อมูลเข้าสู่ระบบ'):",
        ).grid(row=1, column=0, sticky="w", pady=4)
        self.booking_login_input = tk.StringVar(value="")
        ttk.Entry(self.booking_frame, textvariable=self.booking_login_input).grid(
            row=1, column=1, sticky="we", padx=8,
        )
        ttk.Button(
            self.booking_frame, text="เลือก...",
            command=lambda: self._choose_into(self.booking_login_input),
        ).grid(row=1, column=2, padx=4)

        ttk.Label(self.booking_frame, text="จำนวนเดือนที่ตรวจ (นับจากเดือนนี้):").grid(
            row=2, column=0, sticky="w", pady=4,
        )
        self.booking_months_ahead = tk.IntVar(value=3)
        ttk.Spinbox(
            self.booking_frame, from_=1, to=12,
            textvariable=self.booking_months_ahead, width=8,
        ).grid(row=2, column=1, sticky="w", padx=8)
        ttk.Label(
            self.booking_frame,
            text="แนะนำ 1-3 เดือน (มากกว่านั้นใช้เวลานานขึ้น ~2-3 นาที/เดือน/89 สาขา)",
            foreground="gray",
        ).grid(row=2, column=2, sticky="w")

        ttk.Label(self.booking_frame, text="กรองรหัสสาขา (เว้นว่าง = ทุกสาขา):").grid(
            row=3, column=0, sticky="nw", pady=4,
        )
        self.booking_branch_filter = tk.StringVar(value="")  # เก็บ compat; ใช้ Listbox จริง
        bk_lb_box = ttk.Frame(self.booking_frame)
        bk_lb_box.grid(row=3, column=1, columnspan=2, sticky="we", padx=8, pady=4)
        # ช่องค้นหา (filter listbox)
        _search_row = ttk.Frame(bk_lb_box)
        _search_row.pack(fill="x")
        ttk.Label(_search_row, text="ค้นหา:").pack(side="left")
        self.booking_search = tk.StringVar(value="")
        _ent = ttk.Entry(_search_row, textvariable=self.booking_search)
        _ent.pack(side="left", fill="x", expand=True, padx=(4, 8))
        self.booking_search.trace_add("write", lambda *_: self._booking_apply_search())
        ttk.Button(
            _search_row, text="โหลด/รีเฟรชรายชื่อสาขา",
            command=self._booking_load_branches,
        ).pack(side="left", padx=(0, 4))
        ttk.Button(
            _search_row, text="ล้างเลือก",
            command=lambda: self.booking_branches_lb.selection_clear(0, "end"),
        ).pack(side="left")
        # Listbox
        _lb_row = ttk.Frame(bk_lb_box)
        _lb_row.pack(fill="both", expand=True, pady=(4, 0))
        bk_sb = ttk.Scrollbar(_lb_row, orient="vertical")
        self.booking_branches_lb = tk.Listbox(
            _lb_row, selectmode="multiple", height=10,
            exportselection=False, yscrollcommand=bk_sb.set, activestyle="none",
        )
        bk_sb.config(command=self.booking_branches_lb.yview)
        bk_sb.pack(side="right", fill="y")
        self.booking_branches_lb.pack(side="left", fill="both", expand=True)
        # โครงข้อมูล: master list (ไม่กรอง) + visible mapping
        self._booking_all_branches: list[dict] = []       # [{code, name}, ...]
        self._booking_visible_idx: list[int] = []          # เชื่อม row-in-listbox → all-branches idx
        self._booking_selected_codes: set[str] = set()     # เก็บ selection ข้าม filter
        ttk.Label(
            self.booking_frame,
            text="คลิกเลือกได้หลายสาขา (คลิกซ้ำ = ยกเลิก) — ไม่เลือก = ทุกสาขา",
            foreground="gray",
        ).grid(row=4, column=1, columnspan=2, sticky="w", padx=8)

        ttk.Label(
            self.booking_frame,
            text="ปลอดภัย: โหมดอ่านอย่างเดียว — ไม่มีการจอง/แก้ข้อมูลใดๆ (เฉพาะเรียก API เพื่ออ่านข้อมูลว่าง)",
            foreground="#1a7000",
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(6, 0))

        # ปุ่ม Monitor (real-time)
        _mon_row = ttk.Frame(self.booking_frame)
        _mon_row.grid(row=6, column=0, columnspan=3, sticky="we", pady=(8, 0))
        ttk.Button(
            _mon_row, text="🖥️  เปิดหน้าจอ Monitor (Real-Time)",
            command=self._booking_open_monitor,
        ).pack(side="left")
        ttk.Label(
            _mon_row,
            text="เปิดหน้าต่างใหม่ — แสดงที่ว่างสด ๆ + auto-refresh (แนะนำเลือก 5-10 สาขาเพื่อรอบ scan สั้น)",
            foreground="gray",
        ).pack(side="left", padx=(10, 0))
        self.booking_frame.columnconfigure(1, weight=1)

        # โหลด cache ถ้ามี
        self._booking_load_branches_from_cache()

        # ปกติซ่อนไว้ — แสดงเมื่อเลือกโหมดอื่น
        self._on_etk_multi_changed()
        self._on_mode_changed()

        # ---- ปุ่ม ----
        btns = ttk.Frame(self._body)
        btns.pack(fill="x", **pad)
        self._btns_frame = btns
        self.btn_start = ttk.Button(
            btns, text="⬇  ดาวน์โหลด Report",
            command=self._on_start,
        )
        self.btn_start.pack(side="left", padx=4, ipadx=12, ipady=4)
        self.btn_pause = ttk.Button(
            btns, text="⏸  หยุด", command=self._on_pause, state="disabled",
        )
        self.btn_pause.pack(side="left", padx=4, ipady=4)
        self.btn_resume = ttk.Button(
            btns, text="▶  ต่อ", command=self._on_resume, state="disabled",
        )
        self.btn_resume.pack(side="left", padx=4, ipady=4)
        self.btn_cancel = ttk.Button(
            btns, text="ยกเลิก", command=self._on_cancel, state="disabled",
        )
        self.btn_cancel.pack(side="left", padx=4, ipady=4)
        self.btn_open = ttk.Button(
            btns, text="เปิดโฟลเดอร์", command=self._open_folder,
        )
        self.btn_open.pack(side="right", padx=4, ipady=4)

        # ---- Progress ----
        prog_frame = ttk.Frame(self._body)
        prog_frame.pack(fill="x", **pad)
        self._prog_frame = prog_frame
        self.progress = ttk.Progressbar(prog_frame, mode="determinate")
        self.progress.pack(fill="x", side="left", expand=True)
        self.progress_lbl = ttk.Label(prog_frame, text="0 / 0", width=12)
        self.progress_lbl.pack(side="left", padx=8)

        # ---- Log ----
        logf = ttk.LabelFrame(self._body, text="บันทึกการทำงาน", padding=6)
        logf.pack(fill="both", expand=True, **pad)
        self._log_frame = logf
        self.log_text = tk.Text(logf, height=14, wrap="word", state="disabled", font=("Consolas", 10))
        self.log_text.pack(fill="both", expand=True, side="left")
        sb = ttk.Scrollbar(logf, command=self.log_text.yview)
        sb.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=sb.set)

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

    def _on_mode_changed(self) -> None:
        """แสดง/ซ่อน frame ตามโหมดที่เลือก"""
        mode = self.source_mode.get()
        # ซ่อนทุก frame ก่อน
        for f in self._etracking_frames:
            f.pack_forget()
        self.etk_req_no_frame.pack_forget()
        self.aliens_frame.pack_forget()
        self.register_frame.pack_forget()
        self.receipt_frame.pack_forget()
        self.result_frame.pack_forget()
        self.inform_frame.pack_forget()
        self.bt30_frame.pack_forget()
        self.bt44_frame.pack_forget()
        self.billpay_frame.pack_forget()
        self.payrcpt_frame.pack_forget()
        self.appt_frame.pack_forget()
        self.etk_template_frame.pack_forget()
        self.bt30ctn_frame.pack_forget()
        self.namelist_frame.pack_forget()
        self.booking_frame.pack_forget()
        # default output path ตามโหมด
        cur = self.out_path.get()
        defaults = {
            "WA_report.xlsx", "WA_aliens.xlsx", "WA_register_report.xlsx",
            "WA_receipts_report.xlsx", "WA_result_docs_report.xlsx",
            "WA_inform_report.xlsx", "WA_bt30_report.xlsx",
            "WA_bt44_report.xlsx",
            "WA_bill_payment_report.xlsx",
            "WA_payment_receipts_report.xlsx",
            "WA_appointment_report.xlsx",
            "WA_bt30_ctn_report.xlsx",
            "WA_namelist_alien_report.xlsx",
            "WA_booking_availability.xlsx",
            "WA_permit_report.xlsx",
        }
        cur_basename = Path(cur).name if cur else ""
        if mode == "aliens":
            self.aliens_frame.pack(fill="x", padx=10, pady=6)
            if cur_basename in defaults:
                self.out_path.set(str(REPORTS_DIR / "WA_aliens.xlsx"))
        elif mode == "register":
            self.register_frame.pack(fill="x", padx=10, pady=6)
            if cur_basename in defaults:
                self.out_path.set(str(REPORTS_DIR / "WA_register_report.xlsx"))
        elif mode == "receipts":
            self.receipt_frame.pack(fill="x", padx=10, pady=6)
            if cur_basename in defaults:
                self.out_path.set(str(REPORTS_DIR / "WA_receipts_report.xlsx"))
        elif mode == "results":
            self.result_frame.pack(fill="x", padx=10, pady=6)
            if cur_basename in defaults:
                self.out_path.set(str(REPORTS_DIR / "WA_result_docs_report.xlsx"))
        elif mode == "inform":
            self.inform_frame.pack(fill="x", padx=10, pady=6)
            if cur_basename in defaults:
                self.out_path.set(str(REPORTS_DIR / "WA_inform_report.xlsx"))
        elif mode == "bt30":
            self.bt30_frame.pack(fill="x", padx=10, pady=6)
            if cur_basename in defaults:
                self.out_path.set(str(REPORTS_DIR / "WA_bt30_report.xlsx"))
        elif mode == "bt44":
            self.bt44_frame.pack(fill="x", padx=10, pady=6)
            if cur_basename in defaults:
                self.out_path.set(str(REPORTS_DIR / "WA_bt44_report.xlsx"))
        elif mode == "bill_payment":
            self.billpay_frame.pack(fill="x", padx=10, pady=6)
            if cur_basename in defaults:
                self.out_path.set(str(REPORTS_DIR / "WA_bill_payment_report.xlsx"))
        elif mode == "payment_receipts":
            self.payrcpt_frame.pack(fill="x", padx=10, pady=6)
            if cur_basename in defaults:
                self.out_path.set(str(REPORTS_DIR / "WA_payment_receipts_report.xlsx"))
        elif mode == "appointment":
            self.appt_frame.pack(fill="x", padx=10, pady=6)
            if cur_basename in defaults:
                self.out_path.set(str(REPORTS_DIR / "WA_appointment_report.xlsx"))
        elif mode == "bt30_ctn":
            # โหมดนี้แชร์ตัวเลือก e-Tracking (request_types + date filter) → แสดงร่วมกัน
            for f in self._etracking_frames:
                f.pack(fill="x", padx=10, pady=6)
            self.bt30ctn_frame.pack(fill="x", padx=10, pady=6)
            if cur_basename in defaults:
                self.out_path.set(str(REPORTS_DIR / "WA_bt30_ctn_report.xlsx"))
        elif mode == "namelist_alien":
            self.namelist_frame.pack(fill="x", padx=10, pady=6)
            if cur_basename in defaults:
                self.out_path.set(str(REPORTS_DIR / "WA_namelist_alien_report.xlsx"))
        elif mode == "booking_availability":
            self.booking_frame.pack(fill="x", padx=10, pady=6)
            if cur_basename in defaults:
                self.out_path.set(str(REPORTS_DIR / "WA_booking_availability.xlsx"))
        elif mode == "permit_report":
            # แชร์ตัวเลือก e-Tracking (request_types + date + status checkboxes)
            for f in self._etracking_frames:
                f.pack(fill="x", padx=10, pady=6)
            if cur_basename in defaults:
                self.out_path.set(str(REPORTS_DIR / "WA_permit_report.xlsx"))
        else:
            # e-Tracking: ตัวเลือก → รูปแบบรายงาน (Template) → ตัวกรองสถานะ → ดึงเฉพาะเลขคำขอ
            self._etracking_frames[0].pack(fill="x", padx=10, pady=6)
            self.etk_template_frame.pack(fill="x", padx=10, pady=6)
            for f in self._etracking_frames[1:]:
                f.pack(fill="x", padx=10, pady=6)
            self.etk_req_no_frame.pack(fill="x", padx=10, pady=6)
            if cur_basename in defaults:
                self.out_path.set(str(REPORTS_DIR / "WA_report.xlsx"))
        self._reorder_after_login()

    def _reorder_after_login(self) -> None:
        """จัดให้ปุ่ม/progress/log อยู่ล่างสุดหลังเปลี่ยน mode"""
        for child in (
            getattr(self, "_btns_frame", None),
            getattr(self, "_prog_frame", None),
            getattr(self, "_log_frame", None),
        ):
            if child is not None:
                child.pack_forget()
                if child is self._log_frame:
                    child.pack(fill="both", expand=True, padx=10, pady=6)
                else:
                    child.pack(fill="x", padx=10, pady=6)

    def _render_request_listbox(self, filt: str = "") -> None:
        """เติมรายการใน request_listbox ตามคำค้น + re-apply selection ที่เก็บไว้ (ข้าม filter)"""
        filt = (filt or "").strip().lower()
        self.request_listbox.delete(0, "end")
        self._req_visible_idx = []
        for i, (code, lbl) in enumerate(zip(self._req_codes, self._req_labels)):
            if filt and filt not in lbl.lower() and filt not in code.lower():
                continue
            self._req_visible_idx.append(i)
            self.request_listbox.insert("end", lbl)
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
        self._req_selected_codes -= visible_codes  # ตัวที่ visible เดิมลบก่อน
        self._req_selected_codes |= codes_now       # แล้วรวม selection ปัจจุบัน

    def _req_apply_search(self) -> None:
        """เรียกเมื่อพิมพ์ในช่องค้นหา — เก็บ selection เดิมก่อน แล้ว render ใหม่ตามคำค้น"""
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
            self.request_summary.set("[เลือก 0 รายการ] → จะดึงทุกรายการคำขอ")
        elif n == 1:
            self.request_summary.set(f"[เลือก 1 รายการ] code: {codes[0]}")
        else:
            preview = ", ".join(codes[:3])
            more = f" (+{n - 3} อื่น)" if n > 3 else ""
            self.request_summary.set(f"[เลือก {n} รายการ] {preview}{more}")

    def _req_select_all(self) -> None:
        """เลือกทุกรายการ (รวมที่ถูกซ่อนจากคำค้นด้วย)"""
        try:
            self._req_selected_codes = set(self._req_codes)
            self._render_request_listbox(self.request_search.get())
            self._on_request_changed()
        except Exception:
            pass

    def _req_clear(self) -> None:
        """ล้างการเลือกทั้งหมด (รวมที่ถูกซ่อนจากคำค้นด้วย)"""
        try:
            self._req_selected_codes.clear()
            self._render_request_listbox(self.request_search.get())
            self._on_request_changed()
        except Exception:
            pass

    def _on_request_changed(self, _evt=None) -> None:
        codes = self._get_selected_request_codes()
        # เก็บ code แรก (หรือว่าง) ไว้ที่ self.request_type สำหรับ backward compat
        first = codes[0] if codes else ""
        self.request_type.set(first)
        # อัปเดต profile defaults ตาม code แรกที่เลือก
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

    def _choose_register_input(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.register_input.set(path)

    def _choose_into(self, var: tk.StringVar) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            var.set(path)

    def _choose_out(self) -> None:
        REPORTS_DIR.mkdir(parents=True, exist_ok=True)
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialdir=str(REPORTS_DIR),
            initialfile="WA_report.xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if path:
            self.out_path.set(path)

    # ---- Branches (booking mode) ----
    _BRANCHES_CACHE = ROOT / "_branches_cache.json"

    def _booking_persist_selection(self) -> None:
        """สังเคราะห์ selection ปัจจุบัน → self._booking_selected_codes (เก็บข้าม filter)"""
        sel = self.booking_branches_lb.curselection()
        codes_now = set()
        for row_idx in sel:
            if row_idx < len(self._booking_visible_idx):
                master_idx = self._booking_visible_idx[row_idx]
                if 0 <= master_idx < len(self._booking_all_branches):
                    codes_now.add(self._booking_all_branches[master_idx]["code"])
        # อัปเดต: ตัวที่ visible = ให้สิทธิ์ตาม state ปัจจุบัน
        visible_codes = {
            self._booking_all_branches[i]["code"]
            for i in self._booking_visible_idx
            if 0 <= i < len(self._booking_all_branches)
        }
        self._booking_selected_codes -= visible_codes  # ตัวที่ visible เดิมให้ลบก่อน
        self._booking_selected_codes |= codes_now       # แล้วรวม selection ปัจจุบัน

    def _booking_render_listbox(self, filt: str = "") -> None:
        """Render listbox ตาม self._booking_all_branches + filter string"""
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
        # ก่อน re-render ให้เก็บ selection ที่ user ทำก่อน
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
        # เก็บ selection ปัจจุบันก่อนโหลดใหม่
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
                    # เซฟ cache
                    try:
                        self._BRANCHES_CACHE.write_text(
                            json.dumps(branches_clean, ensure_ascii=False, indent=2),
                            encoding="utf-8",
                        )
                    except Exception:
                        pass
                    # อัปเดต UI ใน main thread
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
        """รวม cfg (username/password/type/method) จากช่องข้อมูลเข้าสู่ระบบด้านบน"""
        u = (self.username.get() or "").strip()
        p = self.password.get() or ""
        if not u or not p:
            messagebox.showwarning(
                "ข้อมูลไม่ครบ",
                "กรุณากรอก Username/Password ในช่อง 'ข้อมูลเข้าสู่ระบบ' ก่อน แล้วค่อยกดโหลดสาขา",
            )
            return None
        return {
            "username": u, "password": p,
            "user_type": self.user_type.get(),
            "method": self.method.get(),
            "headless": bool(self.headless.get()),
        }

    def _open_folder(self) -> None:
        p = Path(self.out_path.get()).parent
        if p.exists():
            import os
            os.startfile(str(p))  # type: ignore[attr-defined]

    def _log(self, msg: str) -> None:
        self._log_q.put(msg)

    def _drain_log(self) -> None:
        try:
            while True:
                msg = self._log_q.get_nowait()
                self.log_text.config(state="normal")
                self.log_text.insert("end", msg + "\n")
                self.log_text.see("end")
                self.log_text.config(state="disabled")
        except queue.Empty:
            pass
        self.after(120, self._drain_log)

    def _set_progress(self, done: int, total: int) -> None:
        def _do():
            self.progress["maximum"] = max(total, 1)
            self.progress["value"] = done
            self.progress_lbl.config(text=f"{done} / {total}")
        self.after(0, _do)

    # ---------- actions ----------
    def _relogin_every_value(self) -> int:
        """แปลงตัวเลือก 'Login ใหม่ทุก' (dropdown หรือพิมพ์เอง) → จำนวนรายการ (0 = ปิด)"""
        raw = (self.relogin_every.get() or "").strip().replace(",", "")
        try:
            n = int(raw)
        except ValueError:
            return 0  # "ปิด (ไม่ login ใหม่)" หรือค่าที่แปลงไม่ได้
        return n if n > 0 else 0  # ติดลบ/ศูนย์ = ปิด

    def _on_start(self) -> None:
        mode = self.source_mode.get()
        etk_multi = (mode in ("etracking", "permit_report") and self.etk_multi.get())
        # รวมทุกบัญชีเป็นไฟล์เดียว — เฉพาะโหมด e-Tracking แบบหลายบัญชี (permit_report รวมเป็นไฟล์เดียวอยู่แล้ว)
        etk_combine = (mode == "etracking" and etk_multi and self.etk_combine.get())
        if mode not in ("register", "receipts", "results", "inform", "bt30", "bt44", "bill_payment", "payment_receipts", "appointment", "bt30_ctn") and not etk_multi:
            if not self.username.get().strip() or not self.password.get():
                messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณากรอก Username และ Password")
                return
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")
        self._cancel.clear()
        self.btn_start.config(state="disabled")
        self.btn_cancel.config(state="normal")
        self._pause_evt.set()
        self._set_progress(0, 0)

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
        out = Path(self.out_path.get())
        limit = int(self.limit.get() or 0)
        sub_tabs = [n for n, v in self.alien_subtab_vars.items() if v.get()]
        if mode == "aliens" and not sub_tabs:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณาเลือกอย่างน้อย 1 ตารางในโหมดข้อมูลคนต่างด้าว")
            self.btn_start.config(state="normal")
            self.btn_cancel.config(state="disabled")
            return

        etk_req_excel_str = self.etk_req_excel_input.get().strip() if mode == "etracking" else ""
        etk_req_excel = Path(etk_req_excel_str) if etk_req_excel_str else None
        etk_req_login = Path(self.etk_req_login_input.get()) if mode == "etracking" else None
        if mode == "etracking" and etk_req_excel is not None:
            if not etk_req_excel.exists():
                messagebox.showwarning("ไม่พบไฟล์ Ref_number", f"ไม่พบไฟล์: {etk_req_excel}")
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return
            if not etk_req_login or not etk_req_login.exists():
                messagebox.showwarning("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ UsernameLogin: {etk_req_login}")
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return

        register_input = Path(self.register_input.get()) if mode == "register" else None
        register_row_range = self.register_row_range.get().strip() or None
        if mode == "register":
            if not register_input or not register_input.exists():
                messagebox.showwarning("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ Excel: {register_input}")
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return

        receipt_request = Path(self.receipt_request_input.get()) if mode == "receipts" else None
        receipt_login = Path(self.receipt_login_input.get()) if mode == "receipts" else None
        receipt_row_range = self.receipt_row_range.get().strip() or None
        receipt_name_suffix = self.receipt_name_suffix.get().strip() if mode == "receipts" else ""
        receipt_make_folder = bool(self.receipt_make_folder.get()) if mode == "receipts" else False
        receipt_doc_types: list[str] | None = None
        if mode == "receipts":
            if not receipt_request or not receipt_request.exists():
                messagebox.showwarning("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ RequestData: {receipt_request}")
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return
            if not receipt_login or not receipt_login.exists():
                messagebox.showwarning("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ UsernameLogin: {receipt_login}")
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return
            receipt_doc_types = [
                key for key, var in (
                    ("receipt", self.doc_receipt),
                    ("bt44", self.doc_bt44),
                    ("bt22", self.doc_bt22),
                    ("bt53", self.doc_bt53),
                    ("bt55", self.doc_bt55),
                    ("bt52", self.doc_bt52),
                    ("bt56", self.doc_bt56),
                ) if var.get()
            ]
            if not receipt_doc_types:
                messagebox.showwarning(
                    "ข้อมูลไม่ครบ", "กรุณาเลือกประเภทเอกสารอย่างน้อย 1 อย่าง (ใบเสร็จ / บต.44 / บต.22 / บต.53 / บต.56)",
                )
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return

        result_login = Path(self.result_login_input.get()) if mode == "results" else None
        result_request_type = self.result_request_type.get().strip()
        result_ref_str = self.result_ref_input.get().strip() if mode == "results" else ""
        result_ref = Path(result_ref_str) if result_ref_str else None
        result_doc_keys: list[str] | None = None
        if mode == "results":
            if not result_login or not result_login.exists():
                messagebox.showwarning("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ UsernameLogin: {result_login}")
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return
            if result_ref is not None and not result_ref.exists():
                messagebox.showwarning(
                    "ไม่พบไฟล์ Ref_number",
                    f"ระบุไฟล์ Ref_number แต่หาไม่พบ:\n{result_ref}\n\n"
                    "ถ้าต้องการใช้โหมด e-Tracking ให้กดปุ่ม 'ล้าง' ที่ช่อง Ref_number",
                )
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return
            result_doc_keys = [
                key for key, var in (
                    ("result_notice", self.doc_result_notice),
                    ("request_receipt", self.doc_request_receipt),
                    ("bt50", self.doc_result_bt50),
                ) if var.get()
            ]
            if not result_doc_keys:
                messagebox.showwarning(
                    "ข้อมูลไม่ครบ", "กรุณาเลือกเอกสารอย่างน้อย 1 อย่าง (ใบแจ้งผล / ใบรับคำขอ)",
                )
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return
            cfg["request_type"] = result_request_type

        inform_excel = Path(self.inform_excel_input.get()) if mode == "inform" else None
        inform_login = Path(self.inform_login_input.get()) if mode == "inform" else None
        inform_row_range = self.inform_row_range.get().strip() or None
        inform_commit = bool(self.inform_commit.get()) if mode == "inform" else False
        if mode == "inform":
            if not inform_excel or not inform_excel.exists():
                messagebox.showwarning("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ FormRequestEmployment: {inform_excel}")
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return
            if not inform_login or not inform_login.exists():
                messagebox.showwarning("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ UsernameLogin: {inform_login}")
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return
            if inform_commit:
                if not messagebox.askyesno(
                    "ยืนยันส่งคำขอจริง",
                    "โหมด 'COMMIT' จะบันทึกข้อมูลจริงลงระบบ (ข้อ 7+8) — ยืนยันดำเนินการ?",
                ):
                    self.btn_start.config(state="normal")
                    self.btn_cancel.config(state="disabled")
                    return

        bt30_excel = Path(self.bt30_excel_input.get()) if mode == "bt30" else None
        bt30_login = Path(self.bt30_login_input.get()) if mode == "bt30" else None
        bt30_row_range = self.bt30_row_range.get().strip() or None
        bt30_do_step2 = bool(self.bt30_do_step2.get()) if mode == "bt30" else False
        bt30_submit = bool(self.bt30_submit.get()) if mode == "bt30" else False
        if mode == "bt30":
            if not bt30_excel or not bt30_excel.exists():
                messagebox.showwarning("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ from_bt30: {bt30_excel}")
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return
            if not bt30_login or not bt30_login.exists():
                messagebox.showwarning("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ UsernameLogin: {bt30_login}")
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return
            # ตรวจขนาดไฟล์แนบ 'before run' — ถ้ามีไฟล์เกิน 4 MB ให้ฟ้องและไม่เริ่มรัน
            # (เฉพาะเมื่อทำขั้นตอนที่ 2 ซึ่งมีการแนบเอกสาร — ขั้นตอน 1 อย่างเดียวไม่แนบไฟล์)
            if bt30_do_step2:
                try:
                    _recs = _read_bt30_excel(bt30_excel)
                    _sel = [_recs[i - 1] for i in _parse_row_range(bt30_row_range, len(_recs))]
                    _oversize = _bt30_preflight_doc_sizes(_sel, log=lambda *a: None)
                except Exception as e:
                    messagebox.showerror(
                        "อ่านไฟล์ไม่สำเร็จ",
                        f"ตรวจขนาดไฟล์แนบก่อนรันไม่สำเร็จ:\n{e}",
                    )
                    self.btn_start.config(state="normal")
                    self.btn_cancel.config(state="disabled")
                    return
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
                        f"กรุณาย่อ/บีบอัดไฟล์ให้≤ {_BT30_MAX_DOC_MB:.0f} MB แล้วกดรันใหม่",
                    )
                    self._log(
                        f"⛔ ยกเลิกการรัน — พบไฟล์แนบเกิน {_BT30_MAX_DOC_MB:.0f} MB "
                        f"จำนวน {len(_oversize)} ไฟล์ (ต้องแก้ไขก่อน)"
                    )
                    self.btn_start.config(state="normal")
                    self.btn_cancel.config(state="disabled")
                    return
            if bt30_submit and not bt30_do_step2:
                messagebox.showwarning(
                    "ตั้งค่าไม่ถูกต้อง",
                    "ต้องเปิด 'ทำขั้นตอนที่ 2 ต่อ' ก่อน จึงจะส่งคำขอจริง (Step 7-8) ได้",
                )
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return
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
                    self.btn_start.config(state="normal")
                    self.btn_cancel.config(state="disabled")
                    return

        etk_login = Path(self.etk_login_input.get()) if etk_multi else None
        if etk_multi:
            if not etk_login or not etk_login.exists():
                messagebox.showwarning("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ UsernameLogin: {etk_login}")
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return

        bt44_excel = Path(self.bt44_excel_input.get()) if mode == "bt44" else None
        bt44_login = Path(self.bt44_login_input.get()) if mode == "bt44" else None
        bt44_row_range = self.bt44_row_range.get().strip() or None
        bt44_dry_run = bool(self.bt44_dry_run.get()) if mode == "bt44" else False
        bt44_dry_stop = self.bt44_dry_stop.get() if mode == "bt44" else "step2"
        bt44_check_docs = bool(self.bt44_check_docs.get()) if mode == "bt44" else True
        if mode == "bt44":
            if not bt44_excel or not bt44_excel.exists():
                messagebox.showwarning("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ from_bt44: {bt44_excel}")
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return
            if not bt44_login or not bt44_login.exists():
                messagebox.showwarning("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ UsernameLogin: {bt44_login}")
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return

        billpay_login = Path(self.billpay_login_input.get()) if mode == "bill_payment" else None
        billpay_row_range = self.billpay_row_range.get().strip() or None
        billpay_request_types: list[str] | None = None
        if mode == "bill_payment":
            if not billpay_login or not billpay_login.exists():
                messagebox.showwarning("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ UsernameLogin: {billpay_login}")
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return
            sel_idx = self.billpay_reqtypes_lb.curselection()
            billpay_request_types = [self._billpay_reqtype_codes[i] for i in sel_idx] or None

        payrcpt_request = Path(self.payrcpt_request_input.get()) if mode == "payment_receipts" else None
        payrcpt_login = Path(self.payrcpt_login_input.get()) if mode == "payment_receipts" else None
        payrcpt_row_range = self.payrcpt_row_range.get().strip() or None
        if mode == "payment_receipts":
            if not payrcpt_request or not payrcpt_request.exists():
                messagebox.showwarning("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ Request_Receipt: {payrcpt_request}")
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return
            if not payrcpt_login or not payrcpt_login.exists():
                messagebox.showwarning("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ UsernameLogin: {payrcpt_login}")
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return

        appt_login = Path(self.appt_login_input.get()) if mode == "appointment" else None
        appt_row_range = self.appt_row_range.get().strip() or None
        if mode == "appointment":
            if not appt_login or not appt_login.exists():
                messagebox.showwarning("ข้อมูลไม่ครบ", f"ไม่พบไฟล์ UsernameLogin: {appt_login}")
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return

        bt30ctn_login = Path(self.bt30ctn_login_input.get()) if mode == "bt30_ctn" else None
        bt30ctn_row_range = self.bt30ctn_row_range.get().strip() or None
        bt30ctn_subfolder = bool(self.bt30ctn_make_subfolder.get()) if mode == "bt30_ctn" else False
        bt30ctn_do_ctn = bool(self.bt30ctn_do_ctn.get()) if mode == "bt30_ctn" else True
        bt30ctn_do_appt = bool(self.bt30ctn_do_appointment.get()) if mode == "bt30_ctn" else False
        bt30ctn_do_bt25 = bool(self.bt30ctn_do_bt25.get()) if mode == "bt30_ctn" else False
        bt30ctn_do_receipt = bool(self.bt30ctn_do_receipt.get()) if mode == "bt30_ctn" else False
        if mode == "bt30_ctn":
            if not (bt30ctn_do_ctn or bt30ctn_do_appt or bt30ctn_do_bt25 or bt30ctn_do_receipt):
                messagebox.showwarning(
                    "ข้อมูลไม่ครบ",
                    "กรุณาเลือกอย่างน้อย 1 เอกสาร (บต.30 / บต.25 / ใบนัดหมาย / ใบเสร็จ)",
                )
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return
            # บต.25 มีเฉพาะในรายการคำขอ MT_59 → ถ้ายังไม่เลือก เสนอเพิ่มให้อัตโนมัติ
            if bt30ctn_do_bt25 and "MT_59" not in cfg.get("request_types", []):
                _mt59_label = (
                    "การยื่นคำขอรับใบอนุญาตทำงานของคนต่างด้าวซึ่งได้รับอนุญาต"
                    "ให้อยู่ในราชอาณาจักรเป็นการชั่วคราว"
                )
                _ok = messagebox.askyesno(
                    "บต.25 ต้องเลือกรายการคำขอ MT_59",
                    "การดาวน์โหลด 'แบบ บต.25' ต้องกรอง 'รายการคำขอ' เป็น:\n"
                    f"  • {_mt59_label} (MT_59)\n\n"
                    "แต่ตอนนี้ยังไม่ได้เลือกไว้\n\n"
                    "กด 'Yes' เพื่อเพิ่ม MT_59 ให้อัตโนมัติแล้วรันต่อ\n"
                    "กด 'No' เพื่อยกเลิก แล้วกลับไปเลือกเอง",
                )
                if not _ok:
                    self.btn_start.config(state="normal")
                    self.btn_cancel.config(state="disabled")
                    return
                self._req_selected_codes.add("MT_59")
                self._render_request_listbox(self.request_search.get())
                self._update_request_summary()
                cfg["request_types"] = self._get_selected_request_codes()
                self._log("[i] เพิ่มรายการคำขอ MT_59 อัตโนมัติ (จำเป็นสำหรับ บต.25)")
            # ยืดหยุ่น: ถ้าไม่มี UsernameLogin.xlsx ต้องมี Username/Password ด้านบน
            _has_file = bool(bt30ctn_login and bt30ctn_login.exists())
            _has_creds = bool(self.username.get().strip() and self.password.get())
            if not _has_file and not _has_creds:
                messagebox.showwarning(
                    "ข้อมูลไม่ครบ",
                    "โหมดนี้ต้องมีอย่างใดอย่างหนึ่ง:\n"
                    "  1) ไฟล์ UsernameLogin.xlsx ที่มีข้อมูลบัญชี\n"
                    "  2) กรอก Username/Password ในช่อง 'ข้อมูลเข้าสู่ระบบ' ด้านบน",
                )
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return
            # ถ้าไม่มีไฟล์ → ส่ง None ให้ backend fallback ไปใช้ cfg
            if not _has_file:
                bt30ctn_login = None

        # ---- โหมด namelist_alien: อ่าน options ----
        namelist_form_type = (
            self.namelist_form_type.get().strip() if mode == "namelist_alien" else ""
        )
        namelist_limit = int(self.namelist_limit.get() or 0) if mode == "namelist_alien" else 0
        if mode == "namelist_alien":
            if not namelist_form_type:
                messagebox.showwarning(
                    "ข้อมูลไม่ครบ", "กรุณาระบุ Form Type (เช่น MT_63_2_3103_RENEWAL)",
                )
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return

        # ---- โหมด booking_availability: อ่าน options ----
        booking_login: Path | None = None
        booking_months_ahead = 3
        booking_branch_filter: list[str] | None = None
        if mode == "booking_availability":
            _bl = (self.booking_login_input.get() or "").strip()
            if _bl:
                p = Path(_bl)
                if p.exists():
                    booking_login = p
                else:
                    messagebox.showwarning(
                        "ข้อมูลไม่ครบ", f"ไม่พบไฟล์ UsernameLogin: {p}",
                    )
                    self.btn_start.config(state="normal")
                    self.btn_cancel.config(state="disabled")
                    return
            try:
                booking_months_ahead = max(1, int(self.booking_months_ahead.get() or 3))
            except Exception:
                booking_months_ahead = 3
            # ดึง selection จาก Listbox (รวม selection ที่ซ่อนอยู่จาก filter)
            self._booking_persist_selection()
            if self._booking_selected_codes:
                booking_branch_filter = sorted(self._booking_selected_codes)
            else:
                booking_branch_filter = None

        self._worker = threading.Thread(
            target=self._run,
            args=(mode, cfg, out, limit, sub_tabs, register_input, register_row_range,
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
                  booking_login, booking_months_ahead, booking_branch_filter),
            daemon=True,
        )
        self.btn_pause.config(state="normal")
        self.btn_resume.config(state="disabled")
        self._worker.start()

    def _on_cancel(self) -> None:
        if messagebox.askyesno("ยืนยัน", "ต้องการยกเลิกการดึงข้อมูลใช่หรือไม่?"):
            self._cancel.set()
            self._pause_evt.set()  # ปลุก worker ถ้ากำลังพักค้างอยู่ ให้ออกจาก loop
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
        """ทำงานต่อจากรายการเดิมทันที (ไม่ต้อง login ใหม่)"""
        self._pause_evt.set()
        self.btn_resume.config(state="disabled")
        self.btn_pause.config(state="normal")
        self._log("[▶] ทำงานต่อจากเดิม")

    def _wait_if_paused_or_cancelled(self) -> bool:
        """callback ส่งให้ run_* แทน is_cancelled:
        - ถ้าพักอยู่ → บล็อกค้างไว้จนกว่าจะกด 'ต่อ' หรือ 'ยกเลิก'
        - คืน True เมื่อผู้ใช้ยกเลิก (ให้ loop ออก)
        """
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

    def _done_ok(self, count: int, path: Path) -> None:
        self.btn_start.config(state="normal")
        self.btn_cancel.config(state="disabled")
        self.btn_pause.config(state="disabled")
        self.btn_resume.config(state="disabled")
        self._pause_evt.set()
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
        self.btn_start.config(state="normal")
        self.btn_cancel.config(state="disabled")
        self.btn_pause.config(state="disabled")
        self.btn_resume.config(state="disabled")
        self._pause_evt.set()
        messagebox.showerror("เกิดข้อผิดพลาด", msg)

    # ---- Booking Monitor (real-time) ----
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