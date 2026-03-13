# 📦 Google Sheets Telegram Stock Bot

ระบบ **Stock Management Automation** ที่เชื่อมต่อระหว่าง **Google Sheets** และ **Telegram Bot**  
เพื่อใช้จัดการคลังสินค้า ตรวจสอบสต็อก และแจ้งเตือนการเปลี่ยนแปลงแบบ **Real-time**

ระบบนี้ช่วยให้ทีมสามารถ

- ตรวจสอบสต็อกผ่าน Telegram
- เบิก / เติมสินค้า
- ค้นหาสินค้า
- ตรวจสอบสินค้าคงเหลือต่ำ
- รับแจ้งเตือนเมื่อมีการแก้ไขข้อมูลใน Spreadsheet

โดยไม่ต้องเปิด Spreadsheet ตลอดเวลา

---

# ⚙️ Tech Stack

| Component | Tool |
|---|---|
| Automation | Google Apps Script |
| Database | Google Sheets |
| Messaging | Telegram Bot API |
| Language | JavaScript |

---

# 📊 System Overview


ระบบจะใช้ **Telegram เป็น Interface หลัก**  
ส่วน **Google Sheets เป็นฐานข้อมูล Stock**

---

# 🗂️ Google Sheets Structure

ตัวอย่างโครงสร้าง Sheet `STOCK`

| Column | Description |
|---|---|
| A | Item ID |
| B | Material Code |
| C | Description |
| D | Serial Number |
| E | Total Stock |
| F | DP-CBR |
| G | DP-CCS |
| H | DP-SKO |
| I | DP-RYG |
| J | DP-TRT |

---

# 🤖 Telegram Commands

## ตรวจสอบสินค้า
ใช้คำสั่ง /stock CODE

## ค้นหาสินค้า
ใช้คำสั่ง /KEYWORD

## ดูสินค้าทั้งหมด
ใช้คำสั่ง /allstock

## ตรวจสอบสินค้าคงเหลือต่ำ
/lowstock
