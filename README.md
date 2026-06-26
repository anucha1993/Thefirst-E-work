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

> ไลบรารี 4 ตัวล่างอยู่ใน [requirements.txt](requirements.txt) แล้ว ติดตั้งครั้งเดียวจบ

### ⚡ ติดตั้งรวดเดียว (Copy วางได้เลย) 

วางทั้งบล็อกนี้ลง **PowerShell** ที่โฟลเดอร์โปรเจกต์ — ติดตั้งครบทุกอย่างในคำสั่งเดียว:
```powershell
python -m pip install --upgrade pip; pip install "playwright>=1.47.0" "python-dotenv>=1.0.1" "openpyxl>=3.1.5" "pypdf>=6.0.0"; python -m playwright install chromium
```

หรือถ้ามีไฟล์ `requirements.txt` อยู่แล้ว:
```powershell
python -m pip install --upgrade pip; pip install -r requirements.txt; python -m playwright install chromium
```


### 1. ติดตั้ง Python 3.10 ขึ้นไป
- ดาวน์โหลดจาก https://www.python.org/downloads/
- **ตอนติดตั้ง ติ๊ก ✅ "Add Python to PATH"**
- ตรวจสอบว่าติดตั้งสำเร็จ — เปิด Command Prompt / PowerShell แล้วพิมพ์:
  ```cmd
  python --version
  pip --version
  ```
  ต้องขึ้นเลขเวอร์ชัน (เช่น `Python 3.12.x`)

### 2. เปิด PowerShell ที่โฟลเดอร์โปรเจกต์
- เปิด File Explorer ไปที่ `D:\Programing\E-work`
- กด Shift + คลิกขวาในโฟลเดอร์ → **"Open PowerShell window here"**

### 3. ติดตั้งไลบรารีทั้งหมด (คำสั่ง cmd)

**วิธี A — ติดตั้งลง Python ระบบ (ง่ายสุด)**
```cmd
pip install -r requirements.txt
python -m playwright install chromium
```

**วิธี B — ใช้ Virtual Environment (.venv) แยกสภาพแวดล้อม (แนะนำ)**
```cmd
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
python -m playwright install chromium
```
> ถ้าใช้ **PowerShell** แล้ว `activate` ติด error เรื่อง Execution Policy ให้รันครั้งเดียว:
> ```powershell
> Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
> ```

**ถ้าไม่มีไฟล์ `requirements.txt`** — ติดตั้งทีละตัวได้:
```cmd
pip install "playwright>=1.47.0" "python-dotenv>=1.0.1" "openpyxl>=3.1.5"
python -m playwright install chromium
```

> รอประมาณ 1–3 นาที (ดาวน์โหลด Chromium ~150 MB)

### 4. ตรวจสอบว่าติดตั้งครบ (เลือกทำ)
```cmd
pip show playwright python-dotenv openpyxl
```
ถ้าขึ้นข้อมูลของทั้ง 3 ตัว = พร้อมใช้งาน

---


## ▶️ การใช้งานปกติ (สำหรับ User ทั่วไป)

### วิธีที่ 1 — ใช้ GUI (แนะนำ) ⭐
**ดับเบิลคลิกไฟล์ `run_gui.bat`**

หน้าต่างจะเปิดขึ้น ทำตามขั้นตอน:

| ลำดับ | ทำอะไร |
|---|---|
| 1 | กรอก **Username (อีเมล)** |
| 2 | กรอก **Password** (กดเช็คบ็อกซ์ "แสดงรหัสผ่าน" ถ้าอยากเห็น) |
| 3 | เลือก **ประเภทผู้ใช้** — ค่าเริ่มต้น `ผู้กระทำการแทน` |
| 4 | เลือก **ระบบ** — ค่าเริ่มต้น `E-Workpermit` |
| 5 | (เลือกได้) ติ๊ก "ซ่อนหน้าต่างเบราว์เซอร์" ถ้าไม่อยากให้เด้งขึ้นมา |
| 6 | (เลือกได้) ใส่ "จำกัดจำนวน" เช่น `5` เพื่อทดสอบ — ใส่ `0` = ทั้งหมด |
| 7 | กดปุ่ม **เลือก...** เพื่อระบุที่บันทึกไฟล์ (ค่าเริ่มต้น `WA_report.xlsx`) |
| 8 | กดปุ่ม **⬇ ดาวน์โหลด Report** |
| 9 | รอ — ดูความคืบหน้าจาก Progress bar และกล่อง Log |
| 10 | เสร็จแล้วระบบจะถามว่า "เปิดไฟล์เลยไหม?" → กด **Yes** |

> 💡 **กดยกเลิก** ระหว่างทำงานได้ ระบบจะหยุดหลังเสร็จรายการปัจจุบัน และบันทึกที่ดึงได้แล้วลง Excel

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

## 📁 ไฟล์ในโปรเจกต์

| ไฟล์ | หน้าที่ |
|---|---|
| [run_gui.bat](run_gui.bat) | **เปิด GUI** (ดับเบิลคลิก) |
| [gui_app.py](gui_app.py) | โค้ดหน้าต่าง Tkinter |
| [scrape_wa.py](scrape_wa.py) | โค้ดดึงข้อมูลหลัก (ใช้ได้ทั้ง CLI/GUI) |
| [login_ewp.py](login_ewp.py) | สคริปต์ทดสอบ Login อย่างเดียว |
| [requirements.txt](requirements.txt) | รายการ library ที่ต้องลง |
| [.env](.env) | ค่า Login สำหรับโหมด CLI |
| `WA_report.xlsx` | ไฟล์ Excel ที่ได้หลังรัน |

---

## 🔒 ความปลอดภัย

- **อย่า** push ไฟล์ [.env](.env) ขึ้น Git (มีรหัสผ่าน)
- ถ้าเคยพิมพ์รหัสผ่านในที่สาธารณะ → **เปลี่ยนรหัสทันที**
- GUI ไม่บันทึกรหัสผ่านลงไฟล์ — เก็บไว้เฉพาะใน RAM ระหว่างรัน
