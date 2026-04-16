// ============================================================
// CONFIGURATION - QAC & EHS Dashboard
// ============================================================

// ✅ แก้ไข SHEET_ID ให้ตรงกับ Google Sheet ของคุณ
// วิธีหา: เปิด Google Sheet → URL จะเป็น
// https://docs.google.com/spreadsheets/d/[SHEET_ID]/edit
var SHEET_ID = "1Jm40sTe0ijy-MLVcOPvFaCttSnZMLPcy86spPB849nQ";

// ชื่อชีทต้องตรงกับใน Google Sheet (รวมช่องว่างด้วย)
var SHEET_NAMES = {
  PPE: "PPE Inspec",
  TRAINING: "Training",
  AUDIT_PM: "Sum Audit Tools PM",
  AUDIT_NODE: "Sum Audit Tools CM Node",
  AUDIT_OFC: " Sum Audit Tools CM OFC", // มีช่องว่างนำหน้า
};
