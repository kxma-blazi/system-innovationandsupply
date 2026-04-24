// ============================================================
// BACKEND — Engineer Return Dashboard
// ============================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index_ENGINEER")
    .setTitle("ENGINEER Return Dashboard")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// ดึงข้อมูล Dashboard (มี Cache 5 นาที)
// ============================================================
function getDashboardData() {
  try {
    var cache  = CacheService.getScriptCache();
    var cached = cache.get("dashboardData");
    if (cached) {
      return JSON.parse(cached);
    }

    var result = _buildDashboardData();

    // เก็บ Cache เฉพาะเมื่อสำเร็จ
    if (result.status === "success") {
      cache.put("dashboardData", JSON.stringify(result), CACHE_TTL);
    }
    return result;

  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

// ล้าง Cache (เรียกจาก Frontend ปุ่ม Refresh)
function clearCache() {
  try {
    CacheService.getScriptCache().remove("dashboardData");
    return getDashboardData();
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

// ============================================================
// PRIVATE — สร้างข้อมูล Dashboard จาก Sheet
// ============================================================
function _buildDashboardData() {
  var ss         = SpreadsheetApp.openById(SHEET_ID);
  var allSheets  = ss.getSheets();
  var summary    = { totalAll: 0, totalPending: 0, nearDeadline: 0, overdue: 0 };
  var allData    = [];
  var SKIP_NAMES = ["Setting", "Dashboard", SHEET_NAMES.EMPLOYEE];

  allSheets.forEach(function(sheet) {
    var sName = sheet.getName();
    if (SKIP_NAMES.indexOf(sName) > -1) return;

    var data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return;

    // หาแถว Header
    var headerRowIndex = -1;
    for (var i = 0; i < Math.min(data.length, 5); i++) {
      if (data[i].join("").toUpperCase().indexOf("ORDER NO.") > -1) {
        headerRowIndex = i;
        break;
      }
    }
    if (headerRowIndex === -1) return;

    var headers = data[headerRowIndex].map(function(h) { return h.trim().toUpperCase(); });

    // Map columns
    var col = {
      order:     _findCol(headers, ["ORDER NO."]),
      status:    _findCol(headers, ["ORDER STATUS"]),
      days:      _findCol(headers, ["BORROW DAYS"]),
      applicant: _findCol(headers, ["APPLICANT"]),
      item:      _findCol(headers, ["ITEM NAME"]),
      itemType:  _findCol(headers, ["ITEM TYPE"]),
      domain:    _findCol(headers, ["DOMAIN"]),
      province:  _findCol(headers, ["PROVINCE"]),
      site:      _findCol(headers, ["SITE NAME"]),
      vendor:    _findCol(headers, ["VENDOR"]),
      create:    _findCol(headers, ["CREATE TIME"]),
      podDate:   _findCol(headers, ["1ST POD DATE"])
    };

    data.slice(headerRowIndex + 1).forEach(function(row) {
      var orderNo = _cell(row, col.order);
      if (!orderNo) return;

      summary.totalAll++;

      var status     = _cell(row, col.status);
      var daysRaw    = parseFloat(_cell(row, col.days)) || 0;
      var statusLow  = status.toLowerCase();
      var isPending  = statusLow.indexOf("waiting") > -1 ||
                       (statusLow.indexOf("pod")    === -1 &&
                        statusLow.indexOf("finish") === -1 &&
                        statusLow.indexOf("cancel") === -1);

      var alert = "normal";
      if (isPending) {
        summary.totalPending++;
        if      (daysRaw >= THRESHOLDS.DANGER_DAYS)  { alert = "danger";  summary.overdue++; }
        else if (daysRaw >= THRESHOLDS.WARNING_DAYS) { alert = "warning"; summary.nearDeadline++; }
      }

      allData.push({
        orderNo:    orderNo,
        applicant:  _cell(row, col.applicant) || "-",
        itemName:   _cell(row, col.item)      || "-",
        itemType:   _cell(row, col.itemType)  || "-",
        domain:     _cell(row, col.domain)    || "-",
        province:   _cell(row, col.province)  || "-",
        siteName:   _cell(row, col.site)      || "-",
        vendor:     _cell(row, col.vendor)    || "-",
        createTime: _cell(row, col.create)    || "-",
        podDate:    _cell(row, col.podDate)   || "-",
        borrowDays: daysRaw.toFixed(1),
        status:     status,
        alert:      alert,
        isPending:  isPending,
        sheetName:  sName
      });
    });
  });

  // เรียงจากวันมากสุดก่อน
  allData.sort(function(a, b) { return parseFloat(b.borrowDays) - parseFloat(a.borrowDays); });

  return {
    status:       "success",
    summary:      summary,
    recentOrders: allData,
    updatedAt:    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss")
  };
}

// ============================================================
// HELPERS
// ============================================================
function _findCol(headers, names) {
  for (var i = 0; i < names.length; i++) {
    var idx = headers.indexOf(names[i]);
    if (idx > -1) return idx;
  }
  return -1;
}

function _cell(row, idx) {
  if (idx === -1 || idx >= row.length) return "";
  return (row[idx] || "").toString().trim();
}

// ทดสอบสิทธิ์ (เรียกใน Editor เท่านั้น)
function testAccess() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  Logger.log("OK: " + ss.getName());
  var res = _buildDashboardData();
  Logger.log("Rows: " + res.recentOrders.length + " | Pending: " + res.summary.totalPending);
}