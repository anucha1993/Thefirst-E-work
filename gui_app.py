"""GUI สำหรับดึงรายงาน 'รอยื่นเอกสารเพิ่มเติม' จาก e-WorkPermit
รัน: python gui_app.py
"""
from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from scrape_wa import (
    run_scrape,
    run_scrape_multi,
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
        rt_box = ttk.Frame(opt)
        rt_box.grid(row=1, column=0, columnspan=3, sticky="we", pady=(0, 4))
        rt_box.columnconfigure(0, weight=1)
        vsb = ttk.Scrollbar(rt_box, orient="vertical")
        hsb = ttk.Scrollbar(rt_box, orient="horizontal")
        self.request_listbox = tk.Listbox(
            rt_box, selectmode=tk.EXTENDED, height=8, exportselection=False,
            yscrollcommand=vsb.set, xscrollcommand=hsb.set,
            activestyle="dotbox",
        )
        vsb.config(command=self.request_listbox.yview)
        hsb.config(command=self.request_listbox.xview)
        for lbl in self._req_labels:
            self.request_listbox.insert("end", lbl)
        # pre-select default
        try:
            _idx = self._req_codes.index(_default_code)
            self.request_listbox.selection_set(_idx)
            self.request_listbox.see(_idx)
        except ValueError:
            pass
        self.request_listbox.grid(row=0, column=0, sticky="we")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="we")
        self.request_listbox.bind("<<ListboxSelect>>", self._on_request_changed)

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

        ttk.Label(opt, text="Login ใหม่ทุก (รายการ):").grid(row=7, column=0, sticky="w", pady=4)
        self.relogin_every = tk.StringVar(value="500")
        ttk.Combobox(
            opt, textvariable=self.relogin_every, width=10, state="readonly",
            values=["ปิด (ไม่ login ใหม่)", "500", "1000", "2000"],
        ).grid(row=7, column=1, sticky="w", padx=8)
        ttk.Label(opt, text="กัน session timeout ในงานยาว", foreground="gray").grid(
            row=7, column=2, sticky="w",
        )

        ttk.Label(opt, text="จำกัดจำนวน (0 = ทั้งหมด):").grid(row=8, column=0, sticky="w", pady=4)
        self.limit = tk.IntVar(value=0)
        ttk.Spinbox(opt, from_=0, to=10000, textvariable=self.limit, width=10).grid(
            row=8, column=1, sticky="w", padx=8,
        )

        ttk.Label(opt, text="ไฟล์ที่บันทึก:").grid(row=9, column=0, sticky="w", pady=4)
        self.out_path = tk.StringVar(value=str(REPORTS_DIR / "WA_report.xlsx"))
        ttk.Entry(opt, textvariable=self.out_path).grid(row=9, column=1, sticky="we", padx=8)
        ttk.Button(opt, text="เลือก...", command=self._choose_out).grid(row=9, column=2, padx=4)
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
                 "  • ใบนัดหมาย — tab 'การนัดหมาย' → iframe queue → PDF → ตั้งชื่อ {PASSPORT}_APPOINTMENT.pdf (สถานะ AP/APSS)",
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
        ttk.Checkbutton(
            _doc_box, text="ใบตอบรับ บต.30 (CTN) — สถานะ WP2",
            variable=self.bt30ctn_do_ctn,
        ).pack(side="left", padx=(0, 20))
        ttk.Checkbutton(
            _doc_box, text="ใบนัดหมาย (APPOINTMENT) — สถานะ AP/APSS",
            variable=self.bt30ctn_do_appointment,
        ).pack(side="left")

        self.bt30ctn_make_subfolder = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            self.bt30ctn_frame,
            text="สร้างโฟลเดอร์แยกตาม PASSPORT (เช่น bt30_ctn/MH788309/MH788309_BT30_CTN.pdf)",
            variable=self.bt30ctn_make_subfolder,
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(6, 2))

        ttk.Label(
            self.bt30ctn_frame,
            text="เซฟ PDF → reports/bt30_ctn/  |  รายงาน Excel มี hyperlink ไปที่ไฟล์ PDF (CTN + APPOINTMENT)",
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
        self.bt30ctn_frame.pack_forget()
        self.namelist_frame.pack_forget()
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
        else:
            for f in self._etracking_frames:
                f.pack(fill="x", padx=10, pady=6)
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

    def _get_selected_request_codes(self) -> list[str]:
        """คืน list ของ code ที่ผู้ใช้เลือกใน Listbox ตามลำดับ"""
        try:
            idxs = self.request_listbox.curselection()
        except Exception:
            return []
        return [self._req_codes[i] for i in idxs if 0 <= i < len(self._req_codes)]

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
        try:
            self.request_listbox.selection_set(0, "end")
            self._on_request_changed()
        except Exception:
            pass

    def _req_clear(self) -> None:
        try:
            self.request_listbox.selection_clear(0, "end")
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
        for w in (self.etk_login_entry, self.etk_login_btn):
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
        """แปลงตัวเลือก 'Login ใหม่ทุก' จาก dropdown → จำนวนรายการ (0 = ปิด)"""
        raw = (self.relogin_every.get() or "").strip()
        try:
            return int(raw)
        except ValueError:
            return 0  # "ปิด (ไม่ login ใหม่)" หรือค่าที่แปลงไม่ได้

    def _on_start(self) -> None:
        mode = self.source_mode.get()
        etk_multi = (mode == "etracking" and self.etk_multi.get())
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
        }
        out = Path(self.out_path.get())
        limit = int(self.limit.get() or 0)
        sub_tabs = [n for n, v in self.alien_subtab_vars.items() if v.get()]
        if mode == "aliens" and not sub_tabs:
            messagebox.showwarning("ข้อมูลไม่ครบ", "กรุณาเลือกอย่างน้อย 1 ตารางในโหมดข้อมูลคนต่างด้าว")
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
        if mode == "bt30_ctn":
            if not (bt30ctn_do_ctn or bt30ctn_do_appt):
                messagebox.showwarning(
                    "ข้อมูลไม่ครบ",
                    "กรุณาเลือกอย่างน้อย 1 เอกสาร (ใบตอบรับ หรือ ใบนัดหมาย)",
                )
                self.btn_start.config(state="normal")
                self.btn_cancel.config(state="disabled")
                return
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

        self._worker = threading.Thread(
            target=self._run,
            args=(mode, cfg, out, limit, sub_tabs, register_input, register_row_range,
                  receipt_request, receipt_login, receipt_row_range, receipt_doc_types,
                  receipt_name_suffix, receipt_make_folder,
                  result_login, result_doc_keys, result_ref, etk_multi, etk_login,
                  inform_excel, inform_login, inform_row_range, inform_commit,
                  bt30_excel, bt30_login, bt30_row_range, bt30_do_step2, bt30_submit,
                  bt44_excel, bt44_login, bt44_row_range, bt44_dry_run, bt44_dry_stop, bt44_check_docs,
                  billpay_login, billpay_row_range, billpay_request_types,
                  payrcpt_request, payrcpt_login, payrcpt_row_range,
                  appt_login, appt_row_range,
                  bt30ctn_login, bt30ctn_row_range, bt30ctn_subfolder,
                  bt30ctn_do_ctn, bt30ctn_do_appt,
                  namelist_form_type, namelist_limit),
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
        namelist_form_type: str = "", namelist_limit: int = 0,
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
            else:
                if etk_multi and etk_login is not None:
                    count, path = run_scrape_multi(
                        cfg, etk_login, out, limit=limit,
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


if __name__ == "__main__":
    App().mainloop()
