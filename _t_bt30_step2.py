"""ทดสอบ บต.30 ขั้นตอนที่ 2 แบบ end-to-end (เรียก run_bt30 do_step2=True แถวที่ 1)
ทำครบ: ค้นหา/บันทึกคนต่างด้าว → 2.1 รับรอง+ถัดไป → 2.2 ที่อยู่ → 2.3+2.4 ข้อมูลเพิ่มเติม → 2.5 ถัดไป
หยุดที่ขอบเขต 2.5 — ไม่ส่งคำขอ/ไม่ชำระเงิน

รัน: .venv\\Scripts\\python.exe _t_bt30_step2.py
"""
from __future__ import annotations

from scrape_wa import run_bt30, _read_login_accounts, ROOT, REPORTS_DIR


def main() -> None:
    accounts = _read_login_accounts(ROOT / "UsernameLogin.xlsx")
    acct = next(iter(accounts.values()))
    cfg = {
        "username": acct["username"],
        "password": acct["password"],
        "user_type": acct["type"],
        "method": acct.get("method") or "E-Workpermit",
        "headless": False,
    }
    out = REPORTS_DIR / "WA_bt30_step2_test.xlsx"
    count, path = run_bt30(
        cfg, ROOT / "from_bt30.xlsx", ROOT / "UsernameLogin.xlsx", out,
        row_range="1", do_step2=True, log=print,
    )
    print(f"[done] success={count} report={path}")


if __name__ == "__main__":
    main()
