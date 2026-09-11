"""ตรวจว่า username ในไฟล์รายการคำขอ ตรงกับบัญชีใน UsernameLogin.xlsx ไหม
(วินิจฉัยอาการ 'ดาวน์โหลด user เดียวแล้ว skip ที่เหลือ' — สาเหตุ NO_ACCOUNT)
พิมพ์เฉพาะ username เท่านั้น — ไม่แตะ/ไม่โชว์ password
รัน: python _check_receipt_users.py
"""
from __future__ import annotations
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from scrape_wa import (  # noqa: E402
    _read_request_data, _read_login_accounts,
    _read_request_receipt_excel, _read_ref_numbers,
)

ROOT = Path(__file__).parent


def summarize(label: str, records: list[dict], accounts: dict) -> None:
    c = Counter((r.get("username") or "").strip() for r in records)
    print(f"\n=== {label} ===")
    print(f"  รวม {len(records)} แถว | username ไม่ซ้ำ {len(c)} ค่า")
    matched = 0
    groups_ok = 0
    for u, n in sorted(c.items()):
        key = u.strip().lower()
        ok = key in accounts
        if ok:
            matched += n
            groups_ok += 1
        tag = "OK  " if ok else "MISS"
        note = "(มีใน UsernameLogin)" if ok else "(ไม่พบใน UsernameLogin → กลุ่มนี้จะถูกข้าม)"
        print(f"    [{tag}] {u!r:42} x{n:<4} {note}")
    print(f"  -> กลุ่มที่ login ได้: {groups_ok}/{len(c)} | แถวที่จะทำได้: {matched}/{len(records)}")


accounts = _read_login_accounts(ROOT / "UsernameLogin.xlsx")
print(f"UsernameLogin.xlsx: {len(accounts)} บัญชี")
for k in sorted(accounts.keys()):
    print(f"    - {k}")

for label, reader, fn in [
    ("RequestData.xlsx  (โหมด 'ดาวน์โหลดใบเสร็จ' = receipts)", _read_request_data, "RequestData.xlsx"),
    ("Request_Receipt.xlsx  (โหมด 'ใบเสร็จรับเงินทั้งชุด')", _read_request_receipt_excel, "Request_Receipt.xlsx"),
    ("Ref_number.xlsx", _read_ref_numbers, "Ref_number.xlsx"),
]:
    try:
        recs = reader(ROOT / fn)
        summarize(label, recs, accounts)
    except Exception as e:
        print(f"\n=== {label} ===\n  ERROR: {str(e).splitlines()[0]}")
