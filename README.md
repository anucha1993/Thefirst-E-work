# คู่มือการใช้งาน — โปรแกรมดึงรายงาน e-WorkPermit (รอยื่นเอกสารเพิ่มเติม)

โปรแกรมนี้จะ Login เข้าระบบ e-WorkPermit → กรองคำขอที่มีสถานะ **"รอยื่นเอกสารเพิ่มเติม" (WA)** → เปิดดูทีละรายการเพื่อเก็บ **หมายเหตุ** และ **ข้อมูลคนต่างด้าว** → บันทึกเป็นไฟล์ **Excel (.xlsx)**

---

## 📋 สิ่งที่ต้องเตรียม (ครั้งแรกครั้งเดียว)

### 🧩 โปรแกรม/ไลบรารีที่ต้องติดตั้ง

| รายการ | เวอร์ชัน | ใช้ทำอะไร |
|---|---|---|
| **Python** | 3.10 ขึ้นไป | ตัวรันโปรแกรม |
| **playwright** | ≥ 1.47.0 | ควบคุมเบราว์เซอร์ดึงข้อมูลจากเว็บ |
| **Chromium** (ของ Playwright) | ล่าสุด | เบราว์เซอร์ที่ใช้เปิดเว็บ |
| **python-dotenv** | ≥ 1.0.1 | อ่านค่า Login จากไฟล์ `.env` |
| **openpyxl** | ≥ 3.1.5 | สร้างไฟล์ Excel (.xlsx) |
| **pypdf** | ≥ 6.0.0 | อ่านข้อมูลจากไฟล์ PDF (ใบเสร็จ/ใบแจ้งชำระเงิน) |
| **PyMySQL** | ≥ 1.1.0 | ส่งบันทึกการใช้งานเข้า MySQL |

> ไลบรารี 4 ตัวล่างอยู่ใน [requirements.txt](requirements.txt) แล้ว ติดตั้งครั้งเดียวจบ

### ⚡ ติดตั้งรวดเดียว (Copy วางได้เลย) 

**PowerShell** (คั่นด้วย `;`) — คลิกขวาโฟลเดอร์ → **Open in Terminal**:
```powershell
python -m pip install --upgrade pip; python -m pip install -r requirements.txt; python -m playwright install chromium
```

**cmd.exe** (คั่นด้วย `&&`) — Command Prompt:
```cmd
python -m pip install --upgrade pip && python -m pip install -r requirements.txt && python -m playwright install chromium
```

> ⚠️ **อย่าใช้ `;` ใน cmd** — cmd จะมองเป็นคำสั่งเดียวยาว ๆ แล้ว fail  
> ⚠️ **ใช้ `python -m pip` ทุกครั้ง** — Python 3.14 / Microsoft Store ไม่ได้ทำ `pip.exe` ไว้บน PATH

ถ้ายังไม่มีไฟล์ `requirements.txt` ให้ทดแทน `-r requirements.txt` ด้วย:
```
"playwright>=1.47.0" "python-dotenv>=1.0.1" "openpyxl>=3.1.5" "pypdf>=6.0.0"
```


### 1. ติดตั้ง Python (แนะนำ 3.12)
- ดาวน์โหลด **Python 3.12** จาก https://www.python.org/downloads/release/python-3128/ (Windows installer 64-bit)
- **ตอนติดตั้ง ติ๊ก ✅ "Add python.exe to PATH"**
- **⚠️ อย่าใช้ Python 3.14** — Playwright ยังไม่มี wheel สำหรับ Python 3.14 (เพิ่งมีตอนออกหมาดๆ) → `pip install playwright` จะ fail
- ตรวจสอบว่าติดตั้งสำเร็จ — เปิด Command Prompt / PowerShell แล้วพิมพ์:
  ```cmd
  python --version
  python -m pip --version
  ```
  ต้องขึ้นเลขเวอร์ชัน (เช่น `Python 3.12.x`) — หากเดินเครื่องมีหลายเวอร์ชัน ให้เรียกเจาะด้วย `py -3.12`:
  ```cmd
  py -3.12 --version
  py -3.12 -m pip install -r requirements.txt
  ```

### 2. เปิด PowerShell ที่โฟลเดอร์โปรเจกต์
- เปิด File Explorer ไปที่ `D:\Programing\E-work`
- กด Shift + คลิกขวาในโฟลเดอร์ → **"Open PowerShell window here"**

### 3. ติดตั้งไลบรารีทั้งหมด

**วิธี A — ติดตั้งลง Python ระบบ (ง่ายสุด)**
```cmd
python -m pip install -r requirements.txt
python -m playwright install chromium
```

**วิธี B — ใช้ Virtual Environment (.venv) แยกสภาพแวดล้อม (แนะนำ)**
```cmd
python -m venv .venv
.\.venv\Scripts\activate
python -m pip install -r requirements.txt
python -m playwright install chromium
```
> ถ้าใช้ **PowerShell** แล้ว `activate` ติด error เรื่อง Execution Policy ให้รันครั้งเดียว:
> ```powershell
> Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
> ```

**ถ้าไม่มีไฟล์ `requirements.txt`** — ติดตั้งทีละตัวได้:
```cmd
python -m pip install "playwright>=1.47.0" "python-dotenv>=1.0.1" "openpyxl>=3.1.5" "pypdf>=6.0.0"
python -m playwright install chromium
```

> รอประมาณ 1–3 นาที (ดาวน์โหลด Chromium ~150 MB)

### 4. ตรวจสอบว่าติดตั้งครบ (เลือกทำ)
```cmd
python -m pip show playwright python-dotenv openpyxl pypdf
```
ถ้าขึ้นข้อมูลของทั้ง 4 ตัว = พร้อมใช้งาน

---


## ▶️ การใช้งานปกติ (สำหรับ User ทั่วไป)

### วิธีที่ 1 — ใช้ GUI (แนะนำ) ⭐
**ดับเบิลคลิกไฟล์ `run_gui.bat`**

หน้าต่างแบ่งเป็น 3 ส่วน: **เมนูซ้าย (เลือกโหมด)** → **การ์ดตัวเลือกตรงกลาง** → **แถบปุ่ม + บันทึกการทำงานด้านล่าง**

| ลำดับ | ทำอะไร |
|---|---|
| 1 | เลือกโหมดจาก **เมนูซ้าย** (จัดกลุ่ม 4 หมวด: รายงาน / ดาวน์โหลดเอกสาร / ยื่นคำขอ / คิวถ่ายบัตร) — ด้านบนจะบอกชื่อโหมด คำอธิบาย และป้ายสี: 🟢 อ่านอย่างเดียว · 🟡 บันทึกข้อมูลจริง · 🔴 อาจส่งคำขอจริง/มีค่าธรรมเนียม |
| 2 | ถ้ามีการ์ด **บัญชีเข้าสู่ระบบ** → กรอก Username / Password / ประเภทผู้ใช้ / ระบบ (โหมดที่ใช้บัญชีจาก Excel จะไม่แสดงการ์ดนี้) |
| 3 | กรอกตัวเลือกของโหมด (ไฟล์ Excel, แถวที่จะทำ, เอกสารที่ต้องการ ฯลฯ) — เอาเมาส์ชี้ที่ช่องไฟล์จะมีคำอธิบายคอลัมน์ที่ต้องมี |
| 4 | การ์ด **ไฟล์ผลลัพธ์** — เปลี่ยนที่บันทึกได้ (ค่าเริ่มต้นอยู่ในโฟลเดอร์ `reports/` และต่อท้ายวันที่-เวลาให้อัตโนมัติ) |
| 5 | (เลือกได้) ติ๊ก **ซ่อนหน้าต่างเบราว์เซอร์ขณะทำงาน** มุมขวาบน — ใช้ได้ทุกโหมด |
| 6 | กด **▶ เริ่มทำงาน** (หรือกด `F5`) — ดูความคืบหน้าจาก progress bar และกล่อง "บันทึกการทำงาน" |
| 7 | ระหว่างทำงานกด **⏸ พัก / ▶ ต่อ / ✕ ยกเลิก** ได้ — ยกเลิกจะหยุดหลังเสร็จรายการปัจจุบันและบันทึกที่ได้แล้วลง Excel |
| 8 | เสร็จแล้วระบบจะถามว่า "เปิดไฟล์เลยไหม?" → กด **Yes** หรือกด **📁 เปิดโฟลเดอร์ผลลัพธ์** |

> 💡 โปรแกรมจำค่าที่กรอกล่าสุด (โหมด, Username, path ไฟล์, ขนาดหน้าต่าง) ไว้ใน `_gui_settings.json` — **ไม่เก็บรหัสผ่าน**

### วิธีที่ 2 — ใช้ Command Line (สำหรับนักพัฒนา)

ตั้งค่า [.env](.env) ก่อน:
```env
EWP_USERNAME=your_email@example.com
EWP_PASSWORD=your_password
EWP_USER_TYPE=ผู้กระทำการแทน
EWP_LOGIN_METHOD=E-Workpermit
EWP_HEADLESS=false
```

แล้วรัน:
```powershell
python scrape_wa.py                       # ดึงทั้งหมด → WA_report.xlsx
python scrape_wa.py --limit 5             # ทดสอบ 5 รายการแรก
python scrape_wa.py --out MyReport.xlsx   # ระบุชื่อไฟล์
```

---

## 📊 ไฟล์ Excel ที่ได้

จะมี 36 คอลัมน์ ครอบคลุม:

**คอลัมน์หลัก (10 คอลัมน์)**
- ลำดับ, เลขที่คำขอ, ผู้ยื่น, วันที่ยื่นคำขอ, อัปเดตล่าสุด
- สถานะ, หัวข้อแจ้งเตือน, **หมายเหตุ** (← จุดที่ต้องดู), รายการ, URL คำขอ

**ข้อมูลคนต่างด้าว (26 คอลัมน์)**
- ประเภทผู้ใช้งาน, เลขทะเบียนนิติบุคคล, ชื่อสถานประกอบการ (ไทย/Eng)
- ประเภทนิติบุคคล, วันที่จดทะเบียน, ทุนจดทะเบียน, ที่อยู่
- ใบอนุญาตเลขที่, ออก/หมดอายุ, อีเมล, โทรศัพท์
- ข้อมูลผู้กระทำการแทน, สัญชาติคนต่างด้าว, ข้อมูลบริษัทต้นทาง
- ข้อมูลนายจ้าง ฯลฯ

> Header สีชมพู, ตรึงแถวบนสุดและคอลัมน์ A ไว้ — เลื่อนดูสะดวก

---

## ❓ คำถามที่พบบ่อย

**Q: ใช้เวลานานแค่ไหน?**  
A: ประมาณ 5–10 วินาที/รายการ → 118 รายการใช้ ~15–20 นาที

**Q: ระหว่างทำงานเปิดงานอื่นได้ไหม?**  
A: ได้ แต่อย่าปิดหน้าต่างเบราว์เซอร์ที่โปรแกรมเปิด (ถ้าไม่ได้ติ๊ก headless)

**Q: รันแล้ว Login ไม่ผ่าน**  
A: ตรวจ Username/Password และเลือก **ประเภทผู้ใช้** กับ **ระบบ** ให้ตรงกับที่ใช้ปกติ

**Q: ขึ้น `'pip' is not recognized as an internal or external command`**  
A: Python 3.14 / Microsoft Store install ไม่ได้ทำ `pip.exe` ให้อยู่บน PATH → ให้ใช้ `python -m pip ...` ทุกครั้งแทน (เช่น `python -m pip install -r requirements.txt`)

**Q: ขึ้น `no such option: -m` ตอนรันคำสั่งติดตั้ง one-liner**  
A: cmd.exe ไม่รู้จัก `;` → มองเป็นคำสั่งเดียวยาว → ให้เปลี่ยนเป็น `&&` หรือรันทีละคำสั่ง (ดูหัวข้อ "❡️ ติดตั้งรวดเดียว")

**Q: `pip install playwright` fail ขึ้น `Could not find a version that satisfies the requirement`**  
A: คุณน่าจะใช้ Python 3.14 ซึ่ง Playwright ยังไม่รองรับ — ติดตั้ง **Python 3.12** เพิ่ม (ใช้คู่กันกับ 3.14 ได้) แล้วรันด้วย `py -3.12 -m pip install -r requirements.txt`

**Q: ขึ้น `ModuleNotFoundError: No module named 'dotenv'` (หรือ module อื่น)**  
A: เกิดจากรัน Python คนละตัวกับที่ลง library ไว้ แก้ได้ 2 ทาง:
  1. **ใช้ venv ของโปรเจกต์** (ถ้ามีโฟลเดอร์ `.venv`):
     ```powershell
     .\.venv\Scripts\Activate.ps1
     python scrape_wa.py
     ```
     ถ้าขึ้น error เรื่อง Execution Policy → รันครั้งเดียว:
     `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`
  2. **ติดตั้ง library ลง Python ระบบ** — รันคำสั่งในขั้นตอนติดตั้งซ้ำ

**Q: ขึ้น `ModuleNotFoundError` แบบอื่น**  
A: ยังไม่ได้ติดตั้ง library — รันคำสั่งในขั้นตอนที่ 3 ใหม่

**Q: ขึ้น `playwright._impl._errors.Error: Executable doesn't exist`**  
A: ยังไม่ได้ติดตั้ง browser — รัน `python -m playwright install chromium`

**Q: อยากเปลี่ยน Filter เป็นสถานะอื่น (ไม่ใช่ WA)?**  
A: ใน [scrape_wa.py](scrape_wa.py) ฟังก์ชัน `apply_wa_filter()` แก้ `'WA'` เป็นรหัสสถานะอื่น เช่น `'AP'`, `'RJ'`

---

## 📡 บันทึกการใช้งาน → MySQL (คิดค่าใช้จ่ายตามปริมาณ)

ทุกครั้งที่งานจบ (สำเร็จ / ยกเลิก / error) โปรแกรมจะบันทึก 1 แถว: ใครรัน (ช่อง **ผู้ใช้งาน** มุมขวาบน + Windows user + ชื่อเครื่อง), โหมดอะไร, บัญชี e-WorkPermit ที่ใช้, เวลาเริ่ม-จบ, จำนวนรายการ, ผลลัพธ์

การทำงาน: เขียนลง `usage/usage_pending.jsonl` ในเครื่องก่อนเสมอ → ส่งเข้า MySQL ใน thread แยก → ส่งแล้วย้ายไป `usage/usage_sent_YYYYMM.jsonl` ถ้าเน็ต/ฐานข้อมูลล่ม จะเก็บไว้ส่งตอนเปิดโปรแกรมครั้งถัดไป (ไม่กระทบการทำงานหลัก)

### ติดตั้งฝั่งเซิร์ฟเวอร์ (ทำครั้งเดียว)
สร้าง database และ user ผ่านหน้าควบคุมของโฮสต์ (cPanel/DirectAdmin) หรือถ้าโฮสต์ MySQL เองใช้ `webapp/sql/grants_optional.sql` แล้วรันใน database นั้น (phpMyAdmin → SQL หรือ):
```bash
mysql -h HOST -u USER -p DATABASE < webapp/sql/usage_schema.sql     # ตาราง usage_runs, usage_prices + view สรุป
mysql -h HOST -u USER -p DATABASE < webapp/sql/billing_schema.sql   # ตารางลูกค้า/การชำระเงิน (Lemon Squeezy) + ข้อความติดต่อ
```
แนะนำให้มี user 2 ตัว: ตัวที่ใส่ใน `.env` ของเครื่องผู้ใช้ควรมีสิทธิ์ **INSERT เฉพาะ usage_runs** (ผู้ใช้แก้/ลบ/อ่านข้อมูลไม่ได้) ส่วนเว็บแอปใช้อีก user ที่อ่านได้ทุกตาราง

### ติดตั้งฝั่งเครื่องผู้ใช้ (ทุกเครื่อง)
เพิ่มใน [.env](.env) (ดูตัวอย่างใน [.env.example](.env.example)):
```env
USAGE_ENABLED=true
USAGE_DB_HOST=<IP หรือชื่อเซิร์ฟเวอร์ MySQL>
USAGE_DB_PORT=3306
USAGE_DB_NAME=ewp_usage
USAGE_DB_USER=ewp_app
USAGE_DB_PASSWORD=<รหัสของ ewp_app>
```
แล้วเปิดโปรแกรม กดปุ่ม **สถานะการบันทึกใช้งาน** มุมขวาบน เพื่อทดสอบเชื่อมต่อและดูรายการค้างส่ง หรือรันจาก command line: `python usage_tracker.py`

> ⚠️ MySQL ต้องเปิดรับการเชื่อมต่อจากเครื่องผู้ใช้ (`bind-address`, firewall port 3306) — ถ้าเครื่องอยู่นอกออฟฟิศต้องผ่าน VPN หรือจำกัด IP ที่ `CREATE USER`

### เว็บไซต์ + หลังบ้าน + รับชำระเงิน
โฟลเดอร์ [webapp/](webapp/) (Node.js) มีหน้าเว็บสาธารณะสำหรับขาย (ราคา/นโยบาย/ติดต่อ ตามข้อกำหนดของ Lemon Squeezy), หลังบ้าน `/admin` ดูยอดต่อคน/ต่อโหมด ลูกค้า ส่งยอดไปเรียกเก็บผ่าน Lemon Squeezy, และ webhook รับคำสั่งซื้อ — วิธีติดตั้งดู [webapp/README.md](webapp/README.md)

### คิดเงิน
`SELECT * FROM v_usage_monthly_by_operator WHERE ym = '2026-10';` → ยอดต่อคนต่อเดือนตามราคาใน `usage_prices` (งานที่ยกเลิก/error คิดตามจำนวนที่ทำไปแล้วจริงจาก `progress_done`)

---

## 🌐 repo ของเว็บ (datacorpapi.com)

โฟลเดอร์ `webapp/` ถูกแยกขึ้น GitHub เป็น repo ของตัวเองที่ `anucha1993/datacorpapi.com` (มีเฉพาะเว็บ + `sql/`) เพื่อให้ Plesk ดึงเฉพาะส่วนเว็บ แก้โค้ดเว็บใน repo นี้ตามปกติ แล้วส่งขึ้น repo เว็บด้วย
```bash
git subtree push --prefix=webapp datacorp main
```
(remote `datacorp` = https://github.com/anucha1993/datacorpapi.com.git)

## 📁 ไฟล์ในโปรเจกต์

| ไฟล์ | หน้าที่ |
|---|---|
| [run_gui.bat](run_gui.bat) | **เปิด GUI** (ดับเบิลคลิก) |
| [gui_app.py](gui_app.py) | โค้ดหน้าต่าง Tkinter |
| [scrape_wa.py](scrape_wa.py) | โค้ดดึงข้อมูลหลัก (ใช้ได้ทั้ง CLI/GUI) |
| [login_ewp.py](login_ewp.py) | สคริปต์ทดสอบ Login อย่างเดียว |
| [usage_tracker.py](usage_tracker.py) | บันทึกการใช้งาน → MySQL (คิว + ส่ง) |
| [webapp/sql/usage_schema.sql](webapp/sql/usage_schema.sql) | สร้างตาราง/view/user สำหรับบันทึกการใช้งาน |
| [webapp/](webapp/) | เว็บไซต์ขาย + หลังบ้านดูยอด + รับชำระเงินผ่าน Lemon Squeezy (Node.js + MySQL) |
| [webapp/sql/billing_schema.sql](webapp/sql/billing_schema.sql) | ตารางลูกค้า/คำสั่งซื้อ/subscription จาก Lemon Squeezy |
| [requirements.txt](requirements.txt) | รายการ library ที่ต้องลง |
| [.env](.env) | ค่า Login สำหรับโหมด CLI |
| `WA_report.xlsx` | ไฟล์ Excel ที่ได้หลังรัน |

---

## 🔒 ความปลอดภัย

- **อย่า** push ไฟล์ [.env](.env) ขึ้น Git (มีรหัสผ่าน)
- ถ้าเคยพิมพ์รหัสผ่านในที่สาธารณะ → **เปลี่ยนรหัสทันที**
- GUI ไม่บันทึกรหัสผ่านลงไฟล์ — เก็บไว้เฉพาะใน RAM ระหว่างรัน
