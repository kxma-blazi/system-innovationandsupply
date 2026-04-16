// ============================================================
// QAC & EHS DASHBOARD — SERVER SIDE (Code_QAC.gs)
// ============================================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("Index_QAC")
    .setTitle("QAC & EHS Dashboard")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// MAIN: ดึงข้อมูลทั้งหมดในครั้งเดียว (cached 5 นาที)
// ============================================================
function getAllDashboardData() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get("qac_all");
    if (cached) return JSON.parse(cached);

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const result = {
      ppe: getPPEData(ss),
      training: getTrainingData(ss),
      auditPM: getAuditPMData(ss),
      auditNode: getAuditCMNodeData(ss),
      auditOFC: getAuditCMOFCData(ss),
      updatedAt: Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm"),
    };

    cache.put("qac_all", JSON.stringify(result), 300);
    return result;
  } catch (e) {
    Logger.log("getAllDashboardData Error: " + e);
    return { error: e.toString() };
  }
}

function clearDashboardCache() {
  CacheService.getScriptCache().remove("qac_all");
}

// ============================================================
// 1. PPE INSPEC
// ============================================================
function getPPEData(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.PPE);
  if (!sheet) return { error: "ไม่พบชีต: " + SHEET_NAMES.PPE };

  // 1. ดึงแถวสุดท้ายจริง ๆ ของชีต
  const lastRow = sheet.getLastRow();

  // ป้องกันกรณีชีตว่าง หรือมีแต่หัวข้อ (ข้อมูลจริงเริ่มแถว 4)
  if (lastRow < 4) {
    return {
      summary: {
        complete: 0,
        incomplete: 0,
        total: 0,
        waitingAudit: 0,
        completePercent: 0,
      },
      byProvince: {},
      rounds: { r1: 0, r2: 0, r3: 0, total: 0 },
      tableRows: [],
    };
  }

  // 2. ดึงข้อมูลทั้งหมดจากแถว 1 ถึง lastRow (32 คอลัมน์)
  const data = sheet.getRange(1, 1, lastRow, 32).getValues();

  // ─── Summary จากแถวที่ 1 และ 2 ──────────────
  const summary = {
    complete: Number(data[0][4]) || 0,
    incomplete: Number(data[0][6]) || 0,
    total: Number(data[0][9]) || 0,
    waitingAudit: Number(data[1][4]) || 0,
  };
  summary.completePercent =
    summary.total > 0
      ? Math.round((summary.complete / summary.total) * 1000) / 10
      : 0;

  // ─── นับรายจังหวัด + ผลตรวจแต่ละรอบ (เริ่มแถวที่ 4 คือ index 3) ────────
  const byProvince = {};
  let r1Pass = 0,
    r2Pass = 0,
    r3Pass = 0;
  const tableRows = [];

  for (let i = 3; i < data.length; i++) {
    const row = data[i];

    // ตรวจสอบว่าคอลัมน์ A (ลำดับ) ต้องมีค่า และไม่ใช่ค่าว่างหรือ "nan"
    const rowId = String(row[0] || "").trim();
    if (!rowId || rowId === "" || rowId === "nan") continue;

    // จังหวัด (Column B = index 1)
    const prov = String(row[1] || "").trim();
    if (prov) {
      byProvince[prov] = (byProvince[prov] || 0) + 1;
    }

    // ผลการตรวจ (R1=Col Z, R2=Col AB, R3=Col AF)
    const r1 = String(row[25] || "")
      .trim()
      .toUpperCase();
    const r2 = String(row[27] || "")
      .trim()
      .toUpperCase();
    const r3 = String(row[31] || "")
      .trim()
      .toUpperCase();

    if (r1 === "PASS") r1Pass++;
    if (r2 === "PASS") r2Pass++;
    if (r3 === "PASS") r3Pass++;

    // ฟังก์ชันช่วยจัดรูปแบบวันที่
    const fmtDate = (v) => {
      if (!v || String(v).trim() === "" || String(v).trim() === "-") return "-";
      try {
        // ถ้าเป็น Object Date อยู่แล้วให้ Format เลย ถ้าไม่ใช่ให้ลองแปลงก่อน
        const d = v instanceof Date ? v : new Date(v);
        return Utilities.formatDate(d, "GMT+7", "dd/MM/yy");
      } catch (e) {
        return String(v).trim();
      }
    };

    const resultBadge = (val) => {
      const s = String(val || "")
        .trim()
        .toUpperCase();
      if (s === "PASS") return "PASS";
      if (s === "-" || s === "") return "-";
      return s.substring(0, 15); // ตัดข้อความ Remark ให้ไม่ยาวเกินไป
    };

    // เก็บข้อมูลลง Array เพื่อส่งไปหน้าบ้าน
    tableRows.push({
      no: rowId,
      province: prov,
      team: String(row[2] || "").trim(),
      name: String(row[8] || "").trim(),
      workType: String(row[4] || "").trim(),
      mateline: String(row[5] || "").trim(),
      position: String(row[6] || "").trim(),
      r1Date: fmtDate(row[24]),
      r1Result: resultBadge(row[25]),
      r2Date: fmtDate(row[26]),
      r2Result: resultBadge(row[27]),
      r3Date: fmtDate(row[30]),
      r3Result: resultBadge(row[31]),
    });
  }

  return {
    summary,
    byProvince,
    rounds: {
      r1: r1Pass,
      r2: r2Pass,
      r3: r3Pass,
      total: summary.total,
    },
    // ปรับเพิ่มจำนวนแถวที่จะแสดงหน้าเว็บตามต้องการ (เช่น 500 หรือเอา slice ออกเลย)
    tableRows: tableRows.slice(0, 500),
  };
}

// ============================================================
// 2. TRAINING
// ============================================================
function getTrainingData(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.TRAINING);
  if (!sheet) return { error: "ไม่พบชีต: " + SHEET_NAMES.TRAINING };

  const data = sheet.getRange(1, 1, 5, 38).getValues();

  // data[1][1]=173 total, [1][2]=52 Node, [1][3]=101 OFC, [1][4]=20 PM
  const summary = {
    total: Number(data[1][1]) || 0,
    node: Number(data[1][2]) || 0,
    ofc: Number(data[1][3]) || 0,
    pm: Number(data[1][4]) || 0,
  };

  // G course: data[1][27]=164 trained, data[1][28]=9 waiting
  // EC course: data[1][31]=39 trained, data[1][32]=134 waiting
  const gTrained = Number(data[1][27]) || 0;
  const gWaiting = Number(data[1][28]) || 0;
  const ecTrained = Number(data[1][31]) || 0;
  const ecWaiting = Number(data[1][32]) || 0;

  const gTotal = gTrained + gWaiting || 1;
  const ecTotal = ecTrained + ecWaiting || 1;

  return {
    summary,
    gCourse: {
      trained: gTrained,
      waiting: gWaiting,
      total: gTrained + gWaiting,
      percent: Math.round((gTrained / gTotal) * 1000) / 10,
    },
    ecCourse: {
      trained: ecTrained,
      waiting: ecWaiting,
      total: ecTrained + ecWaiting,
      percent: Math.round((ecTrained / ecTotal) * 1000) / 10,
    },
  };
}

// ============================================================
// 3. AUDIT TOOLS PM
// ============================================================
function getAuditPMData(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.AUDIT_PM);
  if (!sheet) return { error: "ไม่พบชีต: " + SHEET_NAMES.AUDIT_PM };

  const data = sheet.getRange(1, 1, 6, 23).getValues();

  // Summary: data[1][5]=PM NodeB=7, [1][6]=PM Bignode=2, [1][7]=Clear def=1
  const nodeB = Number(data[1][5]) || 0;
  const bignode = Number(data[1][6]) || 0;
  const clearDf = Number(data[1][7]) || 0;

  // Teams: data[3][col] = score (0-1), data[4][col] = name, every 2 cols from col 2
  const teams = [];
  for (let col = 2; col <= 20; col += 2) {
    const name = String(data[4][col] || "").trim();
    if (!name) continue;
    const raw = data[3][col];
    const pct = typeof raw === "number" ? Math.round(raw * 100) : null;
    teams.push({
      name,
      percent: pct,
      status: pct === null ? "Pending" : pct >= 100 ? "Done" : "Not Done",
      remark: String(data[3][col + 1] || "").trim(),
    });
  }

  const done = teams.filter((t) => t.status === "Done").length;
  const notDone = teams.filter((t) => t.status === "Not Done").length;

  return {
    summary: {
      totalTeams: teams.length,
      nodeB,
      bignode,
      clearDf,
      done,
      notDone,
      totalItems: Number(data[4][0]) || 0,
    },
    teams,
  };
}

// ============================================================
// 4. AUDIT TOOLS CM NODE
// ============================================================
function getAuditCMNodeData(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.AUDIT_NODE);
  if (!sheet) return { error: "ไม่พบชีต: " + SHEET_NAMES.AUDIT_NODE };

  const data = sheet.getRange(1, 1, 6, 63).getValues();

  // Summary: data[1][5]=WAH=22, [1][6]=OG=8, [1][7]=inspected=19, [1][8]=pending=11
  const summary = {
    totalTeams: 30,
    wah: Number(data[1][5]) || 0,
    og: Number(data[1][6]) || 0,
    inspected: Number(data[1][7]) || 0,
    pending: Number(data[1][8]) || 0,
    totalItems: Number(data[4][0]) || 0,
  };

  const teams = [];
  for (let col = 2; col <= 60; col += 2) {
    const name = String(data[4][col] || "").trim();
    if (!name) continue;
    const raw = data[3][col];
    const pct = typeof raw === "number" ? Math.round(raw * 100) : null;
    teams.push({
      name,
      percent: pct,
      status: pct === null ? "Pending" : pct >= 100 ? "Done" : "Not Done",
    });
  }

  return { summary, teams };
}

// ============================================================
// 5. AUDIT TOOLS CM OFC
// ============================================================
function getAuditCMOFCData(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.AUDIT_OFC);
  if (!sheet) return { error: "ไม่พบชีต: " + SHEET_NAMES.AUDIT_OFC };

  const data = sheet.getRange(1, 1, 6, 99).getValues();

  // Summary: data[1][5]=2MP=43, [1][6]=3MP=5, [1][7]=inspected=22, [1][8]=pending=26
  const summary = {
    totalTeams: 48,
    twoMP: Number(data[1][5]) || 0,
    threeMP: Number(data[1][6]) || 0,
    inspected: Number(data[1][7]) || 0,
    pending: Number(data[1][8]) || 0,
    totalItems: Number(data[4][0]) || 0,
  };

  const teams = [];
  for (let col = 2; col <= 96; col += 2) {
    const name = String(data[4][col] || "").trim();
    if (!name) continue;
    const raw = data[3][col];
    const pct = typeof raw === "number" ? Math.round(raw * 100) : null;
    teams.push({
      name,
      percent: pct,
      status: pct === null ? "Pending" : pct >= 100 ? "Done" : "Not Done",
    });
  }

  return { summary, teams };
}
