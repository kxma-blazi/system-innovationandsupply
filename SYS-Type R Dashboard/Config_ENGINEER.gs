// ============================================================
// CONFIGURATION — แก้ที่นี่ที่เดียว
// ============================================================

var SHEET_ID = "1UPnPK1msv33sBs-ev-JFQGXb-1p--VqO7JEKjhoQFEg";

var SHEET_NAMES = {
  REPORT: "Engineer Application Report",
  EMPLOYEE: "Sheet1"
};

// กำหนดเกณฑ์วันค้าง
var THRESHOLDS = {
  WARNING_DAYS: 7,   // เหลือง: 7–13 วัน
  DANGER_DAYS:  14   // แดง:    14 วัน+
};

// Cache TTL (วินาที)
var CACHE_TTL = 300; // 5 นาที