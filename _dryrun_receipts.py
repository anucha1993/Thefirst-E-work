# -*- coding: utf-8 -*-
"""Dry-run โหมดดาวน์โหลดใบเสร็จ: อ่านไฟล์จริง + จำลองชื่อไฟล์ที่จะได้
ไม่เปิดเบราว์เซอร์ ไม่ล็อกอิน ไม่โชว์รหัสผ่าน — แค่ยืนยันว่าชื่อไฟล์ไม่ชนกัน (แก้ SKIP_EXISTS)
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from scrape_wa import (  # noqa: E402
    _read_request_data,
    _read_login_accounts,
    _receipt_safe_name,
    DOC_TYPES,
    DOC_TYPES_DEFAULT,
)

req_xlsx = ROOT / "RequestData.xlsx"
login_xlsx = ROOT / "UsernameLogin.xlsx"

records = _read_request_data(req_xlsx)
accounts = _read_login_accounts(login_xlsx)

print("=" * 78)
print(f"RequestData.xlsx : {len(records)} แถว (ที่มีเลขคำขอ)")
print(f"UsernameLogin.xlsx: {len(accounts)} บัญชี")
print("-" * 78)
print("บัญชีล็อกอิน (ไม่โชว์รหัสผ่าน):")
for key, a in accounts.items():
    print(f"  - {a['username']}  | type={a['type']}  | method={a.get('method') or '(default)'}")
print("=" * 78)

# ---- จำลอง ambiguous_passports เหมือนใน run_receipts ----
_pp_to_reqs: dict[str, set[str]] = {}
for _rec in records:
    _pp = _receipt_safe_name(_rec.get("passport", "") or _rec.get("seq", "") or "")
    _pp_to_reqs.setdefault(_pp, set()).add(_rec.get("req_no", ""))
ambiguous = {pp for pp, rq in _pp_to_reqs.items() if len(rq) > 1}
print(f"PASSPORT ที่ซ้ำข้ามหลายเลขคำขอ (ambiguous): {len(ambiguous)} ค่า")
for pp in ambiguous:
    print(f"  * '{pp}'  → ใช้กับ {len(_pp_to_reqs[pp])} เลขคำขอ")
print("=" * 78)

# ---- จัดกลุ่มตาม username + เช็คว่ามีบัญชีไหม ----
groups: dict[str, list[dict]] = {}
for rec in records:
    groups.setdefault(rec["username"], []).append(rec)

print(f"จัดกลุ่มตาม Username: {len(groups)} กลุ่ม")
for u, recs in groups.items():
    acct = accounts.get((u or "").strip().lower())
    tag = "OK พบบัญชี" if acct else "!! ไม่พบบัญชีใน UsernameLogin → จะถูก NO_ACCOUNT"
    print(f"  - '{u}'  ({len(recs)} รายการ)  [{tag}]")
print("=" * 78)

# ---- จำลองชื่อไฟล์ต่อ record (doc types default) ----
doc_types = DOC_TYPES_DEFAULT
print(f"จำลองชื่อไฟล์ (doc_types = {doc_types}), name_suffix='' :")
print("-" * 78)
all_files: list[str] = []
for rec in records:
    seq = rec.get("seq", "")
    req_no = rec.get("req_no", "")
    passport = rec.get("passport", "") or seq
    passport_safe = _receipt_safe_name(passport)
    req_safe = _receipt_safe_name(req_no)
    name_key = f"{passport_safe}_{req_safe}" if (req_safe and passport_safe in ambiguous) else passport_safe
    sample = []
    for dt in doc_types:
        cfg = DOC_TYPES[dt]
        if cfg.get("multi"):
            sample.append(f"{name_key}_{cfg['suffix']}<ราคา>.pdf")
        else:
            fn = f"{name_key}_{cfg['suffix']}.pdf"
            sample.append(fn)
            all_files.append(fn)
    print(f"  #{seq:>3} req={req_no:<16} pass='{passport}'")
    print(f"       name_key='{name_key}'")
    print(f"       → {', '.join(sample)}")
print("=" * 78)

# ---- ตรวจว่าชื่อไฟล์ (แบบ single doc) ไม่ชนกัน ----
dupes = {f for f in all_files if all_files.count(f) > 1}
if dupes:
    print(f"!! ยังมีชื่อไฟล์ชนกัน {len(dupes)} ชื่อ:")
    for f in sorted(dupes):
        print(f"   - {f}")
else:
    print(f"OK ชื่อไฟล์ทั้งหมด {len(all_files)} ไฟล์ (BT44/BT22) ไม่ชนกันแล้ว → จะไม่ถูก SKIP_EXISTS")
print("=" * 78)
