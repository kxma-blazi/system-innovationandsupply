// ============================================================
// SYS-QAC & EHS DASHBOARD — SERVER SIDE (Code_QAC.gs)
// ============================================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("Index_QAC")
    .setTitle("QAC & EHS Dashboard")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onEdit(e) {
  const prop = PropertiesService.getScriptProperties();

  // ✅ อัป version ใหม่เมื่อมีการแก้จริง
  prop.setProperty("DATA_VERSION", new Date().getTime());

  clearDashboardCache();
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
    const prop = PropertiesService.getScriptProperties();
    let version = prop.getProperty("DATA_VERSION");

    if (!version) {
      version = new Date().getTime();
      prop.setProperty("DATA_VERSION", version);
    }

    const result = {
      version: Number(version),
      ppe: getPPEData(ss),
      training: getTrainingData(ss),
      pm:   getAuditUniversalData(ss, SHEET_NAMES.AUDIT_PM.trim()),
      node: getAuditUniversalData(ss, SHEET_NAMES.AUDIT_NODE.trim()),
      ofc: getAuditUniversalData(ss, SHEET_NAMES.AUDIT_OFC.trim()),
      updatedAt: Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm"),
    };

    cache.put("qac_all", JSON.stringify(result), 300);
    return result;
  } catch (e) {
    return { error: e.toString() };
  }
}


// ============================================================
// 🚜 ฟังก์ชันครอบจักรวาล: กำหนดตำแหน่งดึงข้อมูลแยกรายหน้า
// ============================================================
function getAuditUniversalData(ss, sheetName) {
  const sheet = ss.getSheets().find(s => s.getName().trim() === sheetName.trim());
  if (!sheet) return { summary: {}, teams: [] };

  // --- 1. ตั้งค่า Default (ถ้าหน้าไหนไม่ได้ระบุข้างล่าง จะใช้ค่าชุดนี้) ---
  let config = {
    dataRow: 3,      // แถวตัวเลขสรุป
    pctRow: 4,       // แถว % สรุป
    colTotal: "I",   // จำนวนทีม
    colWait: "J",    // รอแก้ Defect
    colDone: "K",    // Complete 100%
    colAudit: "L",   // รอตรวจ Audit
    teamHeaderRow: 7 // แถวชื่อทีม (คำว่า Team)
  };

  // --- 2. ส่วนสำคัญ: กำหนดตำแหน่งแยกตามชื่อหน้า (พี่แก้ตรงนี้ได้เลย) ---
  
  if (sheetName.trim() === SHEET_NAMES.AUDIT_PM.trim()) {
    config.colTotal = "I";
    config.colWait  = "J";
    config.colDone  = "K";
    config.colAudit = "L";
    config.teamHeaderRow = 7;
  }
  else if (sheetName === SHEET_NAMES.AUDIT_NODE.trim()) {
    // หน้า CM Node: พี่บอกว่าจำนวนทีมอยู่ที่ H3
    config.colTotal = "H";     // จำนวนทีม (H3)
    config.colWait = "I";      // รอแก้ (I3)
    config.colDone = "J";      // Done (J3)
    config.colAudit = "K";     // รอตรวจ (K3)
    config.teamHeaderRow = 7;  // ชื่อทีมอยู่แถว 7
  } 
  else if (sheetName === SHEET_NAMES.AUDIT_OFC.trim()) {
    // กำหนดตำแหน่งตามที่ระบุมาเป๊ะๆ
    config = {
      dataRow: 3,        // แถวที่ 3 (สำหรับค่าตัวเลข H3, I3, J3, K3)
      pctRow: 4,         // แถวที่ 4 (สำหรับค่า % สรุป)
      colTotal: "H",     // จำนวนทีมทั้งหมด -> H3
      colWait: "I",      // รอแก้ Defect -> I3
      colDone: "J",      // Complete 100% -> J3
      colAudit: "K",     // รอตรวจเครื่องมือ -> K3
      teamHeaderRow: 7   // แถวรายชื่อทีม
    };
  }

  // --- 3. ส่วนประมวลผลดึงข้อมูล ---
  const colToIndex = (col) => {
    let base = 0, ch = col.toUpperCase();
    for (let i = 0; i < ch.length; i++) base = base * 26 + ch.charCodeAt(i) - 64;
    return base - 1;
  };

  const data = sheet.getRange(1, 1, 15, 100).getValues();
  const r  = config.dataRow - 1;
  const pr = config.pctRow - 1;

  const summary = {
    totalTeams:     Number(data[r][colToIndex(config.colTotal)]) || 0,
    waitingDef:     Number(data[r][colToIndex(config.colWait)])  || 0,
    complete100:    Number(data[r][colToIndex(config.colDone)])  || 0,
    waitingAudit:   Number(data[r][colToIndex(config.colAudit)]) || 0,
    pctWaitingDef:  parseAuditPercent(data[pr][colToIndex(config.colWait)]),
    pctComplete100: parseAuditPercent(data[pr][colToIndex(config.colDone)]),
    pctWaiting:     parseAuditPercent(data[pr][colToIndex(config.colAudit)])
  };

  const teams = [];
  const tr  = config.teamHeaderRow - 1;
  const tpr = tr - 1;

  if (sheetName.trim() === SHEET_NAMES.AUDIT_OFC.trim()) {
    // ✅ OFC: ชื่อทีม C7,E7,G7... % อยู่ C6,E6,G6...
    for (let col = 2; col < 100; col += 2) {
      const name = String(data[tr][col] || "").trim();
      if (!name || name === "Team" || name === "Item") continue;
      const pct = parseAuditPercent(data[tpr][col]);
      teams.push({ name, done: pct, status: pct >= 100 ? "done" : "pending" });
    }
  } else {
    // PM / NODE: หาคอลัมน์ที่มีคำว่า "Team"
    let startCol = -1;
    for (let c = 0; c < 30; c++) {
      if (String(data[tr][c]).trim() === "Team") { startCol = c; break; }
    }
    if (startCol !== -1) {
      for (let col = startCol; col < 100; col += 2) {
        const name = String(data[tr][col] || "").trim();
        if (!name || name === "Team" || name === "Item") continue;
        const pct = parseAuditPercent(data[tpr][col]);
        teams.push({ name, done: pct, status: pct >= 100 ? "done" : "pending" });
      }
    }
  }

  return { summary, teams };
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
  //const summary = {
  //  complete: Number(data[0][4]) || 0,
  //  incomplete: Number(data[0][6]) || 0,
  //  total: Number(data[0][9]) || 0,
  //  waitingAudit: Number(data[1][4]) || 0,
  //};

  // ─── Summary จากแถวที่ 1 ──────────────
const summary = {
  complete:     Number(data[0][2]) || 0,  // C1 = PPE ครบ ✅
  incomplete:   Number(data[0][4]) || 0,  // E1 = PPE ไม่ครบ ✅
  total:        Number(data[0][9]) || 0,  // J1 = Total ✅
  waitingAudit: Number(data[0][6]) || 0,  // G1 = รอ Audit ✅
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
    tableRows: tableRows,
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
// HELPER: แปลงค่าช่อง % ที่อาจเป็น number, "Done", หรือ "=xx/xx"
// ✅ แก้ไข: รองรับกรณีที่ Sheet เก็บค่าเป็น String แทน Number
// ============================================================
function parseAuditPercent(raw) {
  if (raw === null || raw === undefined || raw === "") return 0;
  
  // กรณีเป็นตัวเลขปกติ (0–1) เช่น 0.75 → 75%
  if (typeof raw === "number") {
    // ถ้าเลขมาเป็น 100 อยู่แล้วไม่ต้องคูณ (ป้องกันกรณีใส่เลข 100 มาตรงๆ)
    if (raw > 1) return Math.round(raw); 
    return Math.round(raw * 100);
  }

  if (typeof raw === "string") {
    const s = raw.trim().toUpperCase();
    if (s === "DONE" || s === "PASS") return 100;

    // กรณี "62/62"
    const match = s.match(/^=?(\d+)\/(\d+)$/);
    if (match) {
      const num = parseInt(match[1], 10);
      const den = parseInt(match[2], 10);
      return den > 0 ? Math.round((num / den) * 100) : 0;
    }
    
    // กรณี "90%"
    if (s.includes("%")) return parseInt(s.replace("%", ""), 10);
  }
  return 0;
}

// ============================================================
// 🚜 3. AUDIT PM (พิกัด: I3-L3, ทีมเริ่ม C5/C4)
// ============================================================
function getAuditPMData(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.AUDIT_PM);
  if (!sheet) return { summary: {}, teams: [] };

  const data = sheet.getRange(1, 1, 10, 100).getValues(); 

  const summary = {
    totalTeams:   Number(data[2][8]) || 0,   // I3
    waitingDef:   Number(data[2][9]) || 0,   // J3
    complete100:  Number(data[2][10]) || 0,  // K3
    waitingAudit: Number(data[2][11]) || 0,  // L3
    pctComplete100: parseAuditPercent(data[3][10])
  };

  const teams = [];
  // PM ทีมเริ่มคอลัมน์ C (Index 2), ชื่อแถว 5 (Index 4), % แถว 4 (Index 3)
  for (let col = 2; col < 100; col += 2) {
    const name = String(data[4][col] || "").trim(); 
    if (!name || name === "" || name === "Team") continue;
    const pct = parseAuditPercent(data[3][col]); 
    teams.push({ name, done: pct, status: pct >= 100 ? "done" : "pending" });
  }
  return { summary, teams };
}


// 📡 CM NODE (พิกัด: สรุป I3-L3, ทีมเริ่ม G7/G6)
// ============================================================
function getAuditCMNodeData(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.AUDIT_NODE);
  if (!sheet) return { summary: {}, teams: [] };

  const data = sheet.getRange(1, 1, 10, 100).getValues(); 

  const summary = {
    totalTeams:   Number(data[2][7]) || 0,   // H3 (Index 6) -> ต้องได้ 31
    waitingDef:   Number(data[2][8]) || 0,   // I3 (Index 7) -> ต้องได้ 24
    complete100:  Number(data[2][9]) || 0,  // J3 (Index 8) -> ต้องได้ 6
    waitingAudit: Number(data[2][10]) || 0,  // K3 (Index 9) -> ต้องได้ 1
    
    // เปอร์เซ็นต์บรรทัดที่ 4 (Index 3)
    pctWaitingDef:  parseAuditPercent(data[3][9]),  // J4
    pctComplete100: parseAuditPercent(data[3][10]), // K4
    pctWaiting:     parseAuditPercent(data[3][11])  // L4
  };

  const teams = [];
  // ทีมเริ่มคอลัมน์ G (Index 6), ชื่อแถว 7 (Index 6), % แถว 6 (Index 5)
  for (let col = 6; col < 100; col += 2) {
    const name = String(data[6][col] || "").trim(); 
    if (!name || name === "" || name === "Team") continue;
    const pct = parseAuditPercent(data[5][col]); 
    teams.push({ name, done: pct, status: pct >= 100 ? "done" : "pending" });
  }
  return { summary, teams };
}

function getVersion() {
  const prop = PropertiesService.getScriptProperties();
  const version = prop.getProperty("DATA_VERSION");

  return Number(version) || 0;
}


// ============================================================
// DEBUG — ลบทิ้งหลัง debug เสร็จ
// ============================================================
function debugNodeData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.AUDIT_NODE);
  const data = sheet.getRange(1, 1, 6, 63).getValues();

  Logger.log("Row3 H-K: " + data[2][7] + ", " + data[2][8] + ", " + data[2][9] + ", " + data[2][10]);
  Logger.log("Row4 H-J: " + data[3][7] + ", " + data[3][8] + ", " + data[3][9]);
  Logger.log("Row5 col2: " + data[4][2]);
}

function clearAndDebug() {
  CacheService.getScriptCache().remove("qac_all");
  debugNodeData();
}

function testNode() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const result = getAuditCMNodeData(ss);
  Logger.log(JSON.stringify(result.summary));
}

// Debug
function debugOFCData() {
  CacheService.getScriptCache().remove("qac_all");
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.AUDIT_OFC);
  const data = sheet.getRange(1, 1, 6, 20).getValues();

  Logger.log("Row3 (H-K): " + data[2][7] + " | " + data[2][8] + " | " + data[2][9] + " | " + data[2][10]);
  Logger.log("Row4 (H-K): " + data[3][7] + " | " + data[3][8] + " | " + data[3][9] + " | " + data[3][10]);
  Logger.log("Row5 col C(2): " + data[4][2]);
  Logger.log("Row5 col D(3): " + data[4][3]);
  Logger.log("Row5 col E(4): " + data[4][4]);
  Logger.log("Row5 col F(5): " + data[4][5]);
}

function clearAndReload() {
  CacheService.getScriptCache().remove("qac_all");
}

function debugOFCFull() {
  CacheService.getScriptCache().remove("qac_all");
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.AUDIT_OFC);
  
  if (!sheet) {
    Logger.log("❌ ไม่พบ Sheet ชื่อ: " + SHEET_NAMES.AUDIT_OFC);
    Logger.log("Sheet ที่มีทั้งหมด:");
    ss.getSheets().forEach(s => Logger.log("  → " + s.getName()));
    return;
  }
  
  Logger.log("✅ พบ Sheet: " + sheet.getName());
  Logger.log("lastRow=" + sheet.getLastRow() + ", lastCol=" + sheet.getLastColumn());
  
  const data = sheet.getRange(1, 1, 6, 20).getValues();
  
  for (let r = 0; r < 6; r++) {
    for (let c = 0; c < 20; c++) {
      if (data[r][c] !== "") {
        Logger.log("Row" + (r+1) + " Col" + (c+1) + " [" + String.fromCharCode(65+c) + "]: " + data[r][c]);
      }
    }
  }
}

// Debug
function debugOFCLive() {
  CacheService.getScriptCache().remove("qac_all");
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  // ตรวจสอบชื่อ Sheet ทั้งหมดก่อน
  const allSheets = ss.getSheets().map(s => s.getName());
  Logger.log("📋 All sheets: " + JSON.stringify(allSheets));
  Logger.log("🔍 Looking for: '" + SHEET_NAMES.AUDIT_OFC + "'");

  const sheet = ss.getSheetByName(SHEET_NAMES.AUDIT_OFC.trim());
  if (!sheet) {
    Logger.log("❌ ไม่พบ Sheet!");
    return;
  }

  const data = sheet.getRange(1, 1, 8, 15).getValues();

  // พิมพ์ทุก cell แถว 1-8 ที่มีค่า
  for (let r = 0; r < 8; r++) {
    for (let c = 0; c < 15; c++) {
      if (data[r][c] !== "") {
        Logger.log(`Row${r+1} Col${c+1} [${String.fromCharCode(65+c)}]: "${data[r][c]}" (${typeof data[r][c]})`);
      }
    }
  }
}

// DeBug2
function debugOFCLive2() {
  CacheService.getScriptCache().remove("qac_all");
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  // หา Sheet โดย trim ชื่อ
  const sheet = ss.getSheets().find(s => s.getName().trim() === "Sum Audit Tools CM OFC");
  if (!sheet) { Logger.log("❌ ยังหาไม่เจอ!"); return; }
  
  Logger.log("✅ พบ Sheet: '" + sheet.getName() + "'");
  
  const data = sheet.getRange(1, 1, 8, 15).getValues();
  
  // พิมพ์ทุก cell แถว 1-8 ที่มีค่า
  for (let r = 0; r < 8; r++) {
    for (let c = 0; c < 15; c++) {
      if (data[r][c] !== "") {
        Logger.log(`Row${r+1} [${String.fromCharCode(65+c)}${r+1}]: "${data[r][c]}"`);
      }
    }
  }
}

function testOFCDirect() {
  CacheService.getScriptCache().remove("qac_all");
  const ss = SpreadsheetApp.openById(SHEET_ID);
  Logger.log("AUDIT_OFC name: '" + SHEET_NAMES.AUDIT_OFC + "'");
  const result = getAuditUniversalData(ss, SHEET_NAMES.AUDIT_OFC.trim());
  Logger.log("summary: " + JSON.stringify(result.summary));
  Logger.log("teams: " + result.teams.length);
}
