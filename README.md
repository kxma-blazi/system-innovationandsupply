# 📡 LineBot-SYS2 — Telecom Ticket Close Job Automation (n8n Workflow) & Google App Script

ระบบอัตโนมัติสำหรับรับรายงาน **ปิดงาน (Close Job)** จากช่างซ่อมบำรุงผ่าน LINE OA  
รองรับทั้ง **ข้อความ (Text)** และ **รูปภาพ (Image)** → บันทึกลง Google Sheets → ส่งสรุปกลับ LINE

---

## 🔧 Tech Stack

| Component     | Tool                          |
|---------------|-------------------------------|
| Automation    | [n8n](https://n8n.io)         |
| Trigger       | LINE Messaging API (Webhook)  |
| Storage       | Google Sheets                 |
| Image Storage | Google Drive                  |
| Language      | JavaScript (Code Nodes)       |

---

## 📌 ภาพรวม Workflow

```
LINE Message (Text หรือ Image)
        ↓
IF (Check Message Type)1
   ↙ TRUE (Text)          ↘ FALSE (Image)
Parse close job1          HTTP Request
        ↓                 (ดึงรูปจาก LINE API)
Save Work Detail1                ↓
        ↓                 Upload file → Google Drive
Prepare Sheet Data1              ↓
        ↓                 Attach Ticket Image1
Format Message1                  ↓
        ↓                    (end ⚠️ ยังไม่ merge)
Append row in sheet1
        ↓
Line Messaging2
(ส่งสรุปกลับ LINE)
```

> ⚠️ เส้นรูปภาพ (Image Path) ยังไม่ merge กลับเข้าเส้นหลัก — รูปจะลง Drive แต่ยังไม่ลง Sheet อัตโนมัติ (TODO)

---

## ⚙️ Node และหน้าที่

### 1. `Line Messaging Trigger1`
- รับ Event `message` จาก LINE OA ผ่าน Webhook (Production 24hrs)
- ดึงข้อมูล: `message.type`, `message.id`, `message.text`, `source.userId`, `source.groupId`

---

### 2. `IF (Check Message Type)1`
ตรวจสอบ `$json.message.type`:

| ผลลัพธ์ | เส้นทาง |
|---|---|
| `text` (TRUE) | → Parse close job1 |
| `image` (FALSE) | → HTTP Request |

---

### 3. `Parse close job1` *(Text Path)*

JavaScript Regex แกะข้อมูลจากข้อความของช่าง รองรับรูปแบบ:

```
TT000000000000
Severity: SA2
Fault Date: 00/00/0000
Technician: กอไก่
Province: ขอไข่
Go Time: 16:45
Start Time: 17:37
Done Time: 20:00

Cause:
สายชำรุดเนื่องจากฟ้าผ่า

Fix:
เปลี่ยนสายใหม่และทดสอบสัญญาณ

Timeline:
00:00 รับงานเดินทาง
00:00 ถึงหน้างาน
00:00 เสร็จงาน

00.0000,000.0000
```

**Field ที่ดึงได้:**

| Field | วิธีดึง |
|---|---|
| `ticket_no` | Regex `TT\d+` |
| `severity` | Regex `severity: XXX` |
| `fault_date` | Regex `fault date: ...` |
| `expected_date` | Regex `expected date: ...` |
| `go_time` | Regex `go time: HH:mm` |
| `start_time` | Regex `start time: HH:mm` |
| `done_time` | Regex `done time: HH:mm` |
| `technician` | Regex `technician: ...` |
| `province` | Regex `province: ...` |
| `cause` | Section หลัง `Cause:` |
| `fix` | Section หลัง `Fix:` |
| `timeline` | Section หลัง `Timeline:` หรือบรรทัดที่ขึ้นต้นด้วย `HH:mm` |
| `location` | Regex `lat,lng` / `@lat,lng` / `?q=lat,lng` |

---

### 4. `Save Work Detail1`
จัดระเบียบข้อมูลที่ Parse ได้เป็น Object มาตรฐาน:

```json
{
  "ticket_no", "severity", "fault_date", "expected_date",
  "technician", "province",
  "go_time", "start_time", "done_time",
  "cause", "fix", "timeline",
  "location", "lat", "lng", "map_link",
  "image_url", "userId"
}
```

---

### 5. `Prepare Sheet Data1`
Remap ชื่อ Field ให้ตรงกับหัวคอลัมน์ใน Google Sheets:

| ชื่อใน Workflow | ชื่อใน Sheet |
|---|---|
| `ticket_no` | `tt_no` |
| `fault_date` | `open_date` |
| `cause` | `root_cause` |
| `fix` | `repair_method` |

---

### 6. `Format Message1`
สร้างข้อความสรุปภาษาไทยพร้อมส่งกลับ LINE:

```
🔧 รายงานปิดงาน
━━━━━━━━━━━━━━━━━━━━━━━━━━━
📌 Ticket No. : TT00000000000
📅 วันที่แจ้งเหตุ : 00/00/0000
⏰ กำหนดเสร็จ : 00/00/0000
👷 ช่าง : สมชาย
📍 จังหวัด : ชลบุรี
🗺️ พิกัด : 00.0000,000.0000
━━━━━━━━━━━━━━━━━━━━━━━━━━━
⏱️ เวลาดำเนินงาน
🚗 รับงาน/ออกเดินทาง : 00:00
🏁 ถึงหน้างาน : 00:00
✅ เสร็จงาน : 00:00
━━━━━━━━━━━━━━━━━━━━━━━━━━━
🔍 สาเหตุ ...
🛠️ วิธีแก้ไข ...
📋 Timeline ...
📊 สถานะ : ✅ ปิดงานเรียบร้อย
🗺️ แผนที่ : https://maps.google.com/?q=00.0000,000.0000
```

---

### 7. `Append row in sheet1`
บันทึกข้อมูลลง Google Sheets (Append Row)

- **Spreadsheet:** `testdataLAB(Node)` (`1SpkGa57uRLhoSbc0edGla0QX0gD586tjmrUFJ_xIFgU`)
- **Sheet ID:** `2097907540`
- **Region default:** `UPC05`
- **Columns ที่ Map:** `tt no.`, `severity`, `root_cause`, `Repair Method`, `Admin Name`, `CM report`, `userId`, `ticket_open_date`, `Region`

---

### 8. `Line Messaging2`
ส่งข้อความสรุปกลับไปยัง Group / Room / User ต้นทางใน LINE

---

### 9. `HTTP Request` *(Image Path)*
- ดึงไฟล์ Binary ของรูปภาพจาก LINE Content API
- URL: `https://api-data.line.me/v2/bot/message/{messageId}/content`
- Header: `Authorization: Bearer {Channel Access Token}`

---

### 10. `Upload file`
- อัปโหลดรูปภาพไปเก็บใน Google Drive
- ชื่อไฟล์: `{timestamp}.jpg`
- Folder: `1mLZqQPfs5U9V_Zb2ITU7bV6R2jKbl7re`

---

### 11. `Attach Ticket Image1`
- รับลิงก์รูปภาพจาก Google Drive (`webViewLink` / `webContentLink`)
- Merge กับข้อมูล Ticket จาก `Parse close job1`
- ดึง `userId` จาก Trigger

> ⚠️ **Known Issue:** Node นี้ยังไม่ต่อเข้าเส้นหลัก → รูปภาพยังไม่ลง Sheet โดยอัตโนมัติ

---

## 🗂️ Google Sheets Columns (คอลัมน์หลักที่บันทึก)

| Column | ข้อมูล |
|---|---|
| `Region` | รหัส Region (default: UPC05) |
| `tt no.` | หมายเลข Ticket |
| `severity` | --- |
| `ticket_open_date` | วันที่แจ้งเหตุ |
| `root_cause` | สาเหตุ |
| `Repair Method` | วิธีแก้ไข |
| `Admin Name` | ชื่อช่าง |
| `CM report` | ลิงก์รูปภาพ |
| `userId` | LINE User ID |

---

## 🚀 วิธี Import Workflow

1. เปิด n8n → **Workflows** → **Import from file**
2. เลือกไฟล์ `LineBot-SYS2.json`
3. ตั้งค่า Credentials:
   - **LINE Messaging API** → ใส่ Channel Access Token ของ Bot
   - **Google Sheets OAuth2** → เชื่อมบัญชี Google
   - **Google Drive OAuth2** → เชื่อมบัญชี Google
4. เปิด Webhook เป็น **Production**
5. Toggle เปิดใช้งาน Workflow

---

## ⚠️ Known Issues / TODO

- [ ] **Image Path ยังไม่ลง Sheet** — `Attach Ticket Image1` ไม่ได้ต่อกลับเข้าเส้นหลัก
- [ ] **Node ค้างเมื่อส่งรูป** — `Attach Ticket Image1` อาจหยุดทำงานเมื่อรูปเข้าเร็วเกิน
- [ ] **กันข้อมูลซ้ำ** — ยังไม่ check TT No. ก่อน Append → อาจบันทึกซ้ำได้
- [ ] **Region Code อัตโนมัติ** — ปัจจุบัน hardcode เป็น `UPC05`

---

> **Timezone:** Asia/Bangkok (UTC+7)  
> **Bot Credential:** SYS-Report1  
> **Webhook:** Production (Active 24hrs)
