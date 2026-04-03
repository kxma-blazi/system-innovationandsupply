// ============================================================
// 2. WEB DASHBOARD (หน้าเว็บแสดงผล)
// ============================================================
function doGet(e) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const stockSheet = ss.getSheetByName("STOCK");
  
  // 🎯 ใช้ getDisplayValues เพื่อดึงค่าที่เห็นในเซลล์ (กันพวก Object รูปภาพ)
  const data = stockSheet ? stockSheet.getDataRange().getDisplayValues() : [];

  let rows = [];
  let lowCount = 0;
  let outCount = 0;
  
  // ใช้ค่าจาก Config หรือกำหนดเอง (ต้องตรงกับที่ใช้เช็ค Alert)
  const LIMIT = 5; 

  for (let i = 1; i < data.length; i++) {
    // 🎯 ดึงจาก Column B (Index 1) เท่านั้น
    let matCode = data[i][1] ? String(data[i][1]).trim() : ""; 
    
    // 🛡️ ถ้า B ว่าง หรือเป็นรูปภาพ ให้ขึ้นว่า "ไม่มีรหัส"
    if (matCode === "" || matCode.toLowerCase().includes("cellimage")) {
      matCode = "ไม่มีรหัสสินค้า"; 
    }

    const qty = Number(data[i][5]) || 0; // คอลัมน์ F (Index 5)
    
    // นับจำนวนสินค้าใกล้หมด/หมด
    if (qty <= 0) { outCount++; } 
    else if (qty <= LIMIT) { lowCount++; }

    rows.push({
      code: matCode,             // ข้อมูลจากคอลัมน์ B
      desc: data[i][3] || "-",   // ข้อมูลจากคอลัมน์ D
      qty:  qty,                 // ข้อมูลจากคอลัมน์ F
      cbr:  data[i][6] || 0,     // G
      ccs:  data[i][7] || 0,     // H
      sko:  data[i][8] || 0,     // I
      ryg:  data[i][9] || 0,     // J
      trt:  data[i][10] || 0     // K
    });
  }

  const output = HtmlService.createTemplateFromFile('Index');
  output.rows = rows;
  output.lowCount = lowCount;
  output.outCount = outCount;

  return output.evaluate()
    .setTitle("SYS Stock Dashboard")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ฟังก์ชันสำหรับค้นหาและดึงข้อมูลไปแสดงบนตารางหน้าเว็บ
function getStockRows(query) {
  const ss = SpreadsheetApp.openById(SHEET_ID); // ตรวจสอบว่ามีตัวแปร SHEET_ID ประกาศไว้แล้ว
  const sheet = ss.getSheetByName("STOCK");
  
  // ใช้ getDisplayValues เพื่อกันปัญหาติดรูปภาพในเซลล์
  const data = sheet.getDataRange().getDisplayValues(); 
  const rows = [];
  const q = query ? query.toLowerCase() : "";

  for (let i = 1; i < data.length; i++) {
    let code = data[i][1] ? String(data[i][1]).trim() : "";
    let desc = data[i][3] ? String(data[i][3]).trim() : "";
    
    if (code === "" || code.toLowerCase().includes("cellimage")) {
      continue; // ข้ามแถวที่ไม่มีรหัส
    }

    // ถ้าไม่มีคำค้นหา หรือ คำค้นหาตรงกับรหัส/ชื่อสินค้า ให้เก็บลง Array
    if (q === "" || code.toLowerCase().includes(q) || desc.toLowerCase().includes(q)) {
      rows.push({
        code: code,
        desc: desc,
        qty: Number(data[i][5]) || 0,
        cbr: data[i][6] || "-",
        ccs: data[i][7] || "-",
        sko: data[i][8] || "-",
        ryg: data[i][9] || "-",
        trt: data[i][10] || "-"
      });
    }
  }
  return rows;
}

// 🚩 อย่าลืมฟังก์ชันสร้างแถว ตรวจสอบว่าใช้ r.code นะครับ
function buildHtmlRows(rows) {
  return rows.map(r => {
    const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${encodeURIComponent(r.code)}`;
    const statusClass = r.qty <= 0 ? "status-red" : r.qty <= 5 ? "status-yellow" : "status-green";
    
    return `<tr class="${statusClass}">
      <td data-label="ID"><b>${r.code}</b></td> 
      <td data-label="QR" style="text-align:center;"><img src="${qrUrl}" width="40"></td>
      <td data-label="Description">${r.desc}</td>
      <td data-label="QTY">${r.qty}</td>
      <td>${r.cbr}</td><td>${r.ccs}</td><td>${r.sko}</td><td>${r.ryg}</td><td>${r.trt}</td>
      <td data-label="Status"><span class="badge">${r.qty <= 0 ? 'หมด' : 'ปกติ'}</span></td>
    </tr>`;
  }).join("");
}                                                                                                                                                                                        

/**
 * ดึงข้อมูลสินค้าที่ใกล้หมด (ฉบับปรับปรุงให้เสถียรขึ้น)
 */
function getStockAlerts() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const stockSheet = ss.getSheetByName("STOCK");
    const data = stockSheet.getDataRange().getDisplayValues();
    const alerts = [];
    const LIMIT = 5;

    for (let i = 1; i < data.length; i++) {
      let matCode = data[i][1] ? String(data[i][1]).trim() : "";
      if (matCode === "" || matCode.toLowerCase().includes("cellimage")) continue;

      const qty = Number(data[i][5]) || 0;
      if (qty <= 0) {
        alerts.push({ code: matCode, name: data[i][3], qty: qty, level: "critical", icon: "🔴" });
      } else if (qty <= LIMIT) {
        alerts.push({ code: matCode, name: data[i][3], qty: qty, level: "warning", icon: "🟡" });
      }
    }
    return alerts.sort((a, b) => a.qty - b.qty);
  } catch (e) {
    return [];
  }
}
/**
 * ดึงสถิติแต่ละแผนก
 */
function getDepartmentStats() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const stockSheet = ss.getSheetByName("STOCK");
    const rawData = stockSheet.getRange(1, 1, stockSheet.getLastRow(), 11).getValues();
    
    const DEPTS = {
      "CBR": 6,
      "CCS": 7,
      "SKO": 8,
      "RYG": 9,
      "TRT": 10
    };
    
    const stats = {};
    
    for (const [deptName, colIdx] of Object.entries(DEPTS)) {
      let totalItems = 0;
      let totalQty = 0;
      let itemsWithStock = 0;
      
      for (let i = 1; i < rawData.length; i++) {
        const code = rawData[i][1];
        const deptQty = Number(rawData[i][colIdx]) || 0;
        
        if (code && !String(code).includes("CellImage") && deptQty > 0) {
          totalItems++;
          totalQty += deptQty;
          itemsWithStock++;
        }
      }
      
      stats[deptName] = {
        items: itemsWithStock,
        quantity: totalQty,
        efficiency: itemsWithStock > 0 ? Math.round((itemsWithStock / rawData.length) * 100) : 0
      };
    }
    
    return stats;
  } catch (e) {
    Logger.log("Error getDepartmentStats: " + e);
    return {};
  }
}

/**
 * ดึง Top Low Stock Items
 */
function getTopLowStockItems(limit = 10) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const stockSheet = ss.getSheetByName("STOCK");
    const rawData = stockSheet.getRange(1, 1, stockSheet.getLastRow(), 6).getValues();
    
    const items = [];
    
    for (let i = 1; i < rawData.length; i++) {
      let matCode = "";
      let codeValue = rawData[i][1];
      
      if (codeValue && String(codeValue).includes("CellImage")) {
        codeValue = rawData[i][2];
      }
      
      if (codeValue) {
        matCode = typeof codeValue === 'number' ? codeValue.toFixed(0) : String(codeValue).trim();
      }

      if (matCode && !matCode.includes("CellImage")) {
        items.push({
          code: matCode,
          name: rawData[i][3],
          qty: Number(rawData[i][5])
        });
      }
    }
    
    return items.sort((a, b) => a.qty - b.qty).slice(0, limit);
  } catch (e) {
    Logger.log("Error getTopLowStockItems: " + e);
    return [];
  }
}

/**
 * ดึงสถิติการเบิก (ย้อนหลัง 7 วัน)
 */
function getWithdrawalStats() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const logSheet = ss.getSheetByName("Logs");
    if (!logSheet) return { daily: {}, total: 0, avgPerDay: 0 };
    
    const data = logSheet.getDataRange().getValues();
    const sevenDaysAgo = new Date();
    sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
    
    const daily = {};
    let total = 0;
    
    for (let i = 1; i < data.length; i++) {
      const date = new Date(data[i][0]);
      const action = data[i][2];
      const qty = Number(data[i][4]) || 0;
      
      if (action === "Withdraw" && date > sevenDaysAgo) {
        const dateStr = Utilities.formatDate(date, "GMT+7", "dd/MM");
        daily[dateStr] = (daily[dateStr] || 0) + qty;
        total += qty;
      }
    }
    
    return {
      daily: daily,
      total: total,
      avgPerDay: Math.round(total / 7)
    };
  } catch (e) {
    Logger.log("Error getWithdrawalStats: " + e);
    return { daily: {}, total: 0, avgPerDay: 0 };
  }
}

// ============================================================
// TEST: ตรวจสอบ Material Code Extraction
// ============================================================
function testMaterialCodeExtraction() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const stockSheet = ss.getSheetByName("STOCK");
  const rawData = stockSheet.getRange(1, 1, 20, 12).getValues();

  Logger.log("═══════════════════════════════════════");
  Logger.log("🔍 Material Code Extraction Test");
  Logger.log("═══════════════════════════════════════");

  for (let i = 1; i < rawData.length; i++) {
    const colB = rawData[i][1];
    const colC = rawData[i][2];
    
    let extracted = "";
    if (colB && !String(colB).includes("CellImage")) {
      extracted = typeof colB === 'number' ? colB.toFixed(0) : String(colB).trim();
    } else if (colC && !String(colC).includes("CellImage")) {
      extracted = typeof colC === 'number' ? colC.toFixed(0) : String(colC).trim();
    }

    Logger.log(`Row ${i + 1}: ✅ ${extracted}`);
  }
  
  Logger.log("═══════════════════════════════════════");
}
function getInventoryTurnover() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const logSheet = ss.getSheetByName("Logs");
  const stockSheet = ss.getSheetByName("STOCK");
  
  const logData = logSheet.getDataRange().getValues();
  const stockData = stockSheet.getDataRange().getValues();
  
  // นับการเบิกล่าสุด 30 วัน
  const thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
  
  const recentWithdraws = logData.filter(r => 
    new Date(r[0]) > thirtyDaysAgo && r[2] === 'Withdraw'
  );
  
  return {
    totalWithdraws: recentWithdraws.length,
    avgPerDay: (recentWithdraws.length / 30).toFixed(1),
    topItems: getTopWithdrawItems(recentWithdraws)
  };
}

// ปรับปรุงฟังก์ชันดึงข้อมูล SERIAL_DATA
function getDeliveryRecords(limit = 100) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("SERIAL_DATA");
    if (!sheet) return { records: [] };
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { records: [] };
    
    return {
      records: data.slice(1)
        .filter(r => r[2] !== "") // กรองเฉพาะแถวที่มีซีเรียลใน Col C (นำจ่ายแล้ว)
        .reverse()
        .slice(0, limit)
        .map(r => ({
          code: r[0],         // A: Code
          serial: r[2],       // C: Serial นำจ่าย
          dateDelivered: r[3] ? Utilities.formatDate(new Date(r[3]), "GMT+7", "dd/MM/yy HH:mm") : "-", // D: วันที่
          note: r[4] || "-",  // E: แผนก
          deliveredBy: r[5] || "-" // F: ผู้เบิก
        }))
    };
  } catch (e) { 
    // ตรวจสอบเวลาพังได้
    Logger.log("Error getDeliveryRecords: " + e.toString()); 
    return { records: [] }; 
  }
}

// ============================================================
// 3. TELEGRAM LOGIC (ระบบสั่งการผ่านบอท)
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.callback_query) return; 
    if (!data.message || !data.message.text) return;
    
    const chatId = data.message.chat.id;
    const text = data.message.text.trim();
    const user = data.message.from.first_name || "User";
    const args = text.split(/\s+/);
    const cmd = args[0].toLowerCase();

    // 1. เมนูช่วยเหลือ
    if (cmd === "/start" || cmd === "/help" || cmd === "/menu") {
      const welcomeMsg = "📦 *ระบบจัดการสต็อก SYS (Auto-Detect)*\n\n" +
                         "🔹 *วิธีเบิกสินค้า:*\n" +
                         "• ใส่ซีเรียลหรือรหัส: \`/dp-xxx [SN หรือ ID]\`\n" +
                         "• ระบุจำนวนเพิ่ม: \`/dp-xxx [ID] [จำนวน]\` \n\n" +
                         "*(ระบบจะค้นหาซีเรียลก่อน ถ้าไม่เจอจะเบิกเป็นจำนวนให้เอง)*";
      sendTelegram(chatId, welcomeMsg);
    }
    
    // 2. คำสั่งเบิกสินค้า (เช็คจาก DEPT_MAP ใน Config.gs)
    else if (DEPT_MAP[cmd]) {
      const input = args[1]; 
      if (!input) {
        sendTelegram(chatId, `⚠️ กรุณาระบุรหัสสินค้าหรือซีเรียล\nเช่น: \`${cmd} SN999\` หรือ \`${cmd} Q001\``);
        return;
      }

      const colIdx = DEPT_MAP[cmd];
      const deptName = cmd.replace("/dp-","").toUpperCase();

      // --- ขั้นตอนที่ 1: ลองเบิกแบบซีเรียลก่อน ---
      const serialResult = withdrawBySerial(input, colIdx, deptName, user);

      // ถ้าเจอซีเรียล (เบิกสำเร็จ หรือ ถูกเบิกไปแล้ว)
      if (serialResult !== "NOT_FOUND") {
        sendTelegram(chatId, serialResult);
        // ส่งแจ้งเตือนเข้ากลุ่ม (CHAT_ID จาก Config)
        if (typeof CHAT_ID !== 'undefined' && !serialResult.includes("❌")) {
          sendTelegram(CHAT_ID, `📣 *บันทึก (SN):* ${serialResult}\n(โดย: ${user})`);
        }
        return;
      }

      // --- ขั้นตอนที่ 2: ถ้าไม่ใช่ซีเรียล ให้เบิกแบบจำนวน (QTY) ---
      let qtyMatch = (args[2] || "1").match(/\d+/); 
      let qty = qtyMatch ? parseInt(qtyMatch[0], 10) : 1;

      const result = withdraw(input, qty, colIdx, deptName, user);
      sendTelegram(chatId, result);
      
      if (typeof CHAT_ID !== 'undefined' && !result.includes("❌")) {
        sendTelegram(CHAT_ID, `📣 *บันทึก (QTY):* ${result}\n(โดย: ${user})`);
      }
    } 
    
    // 3. คำสั่งเช็คสต็อก
    else if (cmd === "/stock") {
      sendTelegram(chatId, getStockInfo(args[1]));
    } 
    
    // 4. คำสั่งเติมสต็อก
    else if (cmd === "/restock") {
       if (args.length < 3) {
         sendTelegram(chatId, "⚠️ รูปแบบผิด! ต้องเป็น: `/restock [รหัส] [จำนวน]`");
       } else {
         let qtyMatch = args[2].match(/\d+/); 
         let qty = qtyMatch ? parseInt(qtyMatch[0], 10) : NaN;
         sendTelegram(chatId, restock(args[1], qty, user));
       }
    } 
    
    // 5. ถ้าไม่ใช่คำสั่ง ให้ส่งไปหา Gemini AI
    else {
      sendTelegram(chatId, callGemini(text, chatId));
    }
  } catch (err) {
    Logger.log("doPost Error: " + err.toString());
  }
}

// ฟังก์ชันเบิกสินค้า (แกนหลัก)
function withdraw(code, qty, colIndex, deptName, user) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // รอคิว 15 วินาที
    
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("STOCK");
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      // ค้นหาจาก Column C (Index 2) ตามที่เฮียกำหนดไว้
      if (data[i][2].toString().toLowerCase() === code.toString().toLowerCase()) {
        const currentTotal = Number(data[i][5]); 
        if (currentTotal < qty) return `❌ ${code} ของไม่พอ (คงเหลือ ${currentTotal})`;
        
        const newTotal = currentTotal - qty;
        const newDeptTotal = (Number(data[i][colIndex]) || 0) + qty;
        
        sheet.getRange(i+1, 6).setValue(newTotal); // Col F
        sheet.getRange(i+1, colIndex+1).setValue(newDeptTotal);
        
        SpreadsheetApp.flush(); // บังคับเขียนข้อมูลทันที
        writeLog(user, "Withdraw", code, qty, deptName, newTotal);
        
        let resMsg = `✅ *เบิกสำเร็จ*\n📦 รายการ: ${data[i][3]}\n📍 แผนก: ${deptName}\n📉 จำนวน: -${qty}\n📊 เหลือรวม: ${newTotal}`;
        if (newTotal <= LOW_STOCK_LIMIT) resMsg += `\n\n⚠️ *ALERT: สินค้าใกล้หมดแล้ว!*`;
        return resMsg;
      }
    }
    return `❌ ไม่พบรหัสสินค้า "${code}" ในระบบ`;
  } finally {
    lock.releaseLock(); // ปลดล็อคเสมอ
  }
}

// เช็คข้อมูลในหน้า SERIAL_DATA
function withdrawBySerial(serial, colIndex, deptName, user) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const serialSheet = ss.getSheetByName("SERIAL_DATA");
  if (!serialSheet) return "❌ ไม่พบชีต 'SERIAL_DATA'";

  const sData = serialSheet.getDataRange().getValues();
  let foundRow = -1;
  let takerInfo = null;
  let matCode = "";
  serial = serial.toString().trim();

  for (let i = 1; i < sData.length; i++) {
    // 1. เช็คว่ามีคนเบิกไปหรือยัง (Col C)
    if (sData[i][2].toString().trim() === serial) {
      takerInfo = { matCode: sData[i][0], date: sData[i][3], name: sData[i][5] };
      break;
    }
    // 2. เช็คว่าซีเรียลยังว่างอยู่ไหม (Col B)
    if (sData[i][1].toString().trim() === serial) {
      foundRow = i + 1;
      matCode = sData[i][0];
      break;
    }
  }

  // กรณีเบิกซ้ำ
  if (takerInfo) {
    return `⚠️ *ซีเรียลนี้ถูกเบิกไปแล้ว!*\n👤 โดย: ${takerInfo.name}\n📅 เมื่อ: ${Utilities.formatDate(new Date(takerInfo.date), "GMT+7", "dd/MM/yy HH:mm")}\n📦 รหัสสินค้า: ${takerInfo.matCode}`;
  }

  // ถ้าไม่เจอซีเรียลเลย ส่งกลับไปให้ doPost ทำงานต่อแบบ ID
  if (foundRow === -1) return "NOT_FOUND";

  // ทำการเบิก: ย้ายจาก B ไป C
  const now = new Date();
  serialSheet.getRange(foundRow, 2).clearContent(); // ลบจากช่องว่าง (Col B)
  serialSheet.getRange(foundRow, 3).setValue(serial); // ใส่ช่องนำจ่าย (Col C)
  serialSheet.getRange(foundRow, 4).setValue(now); // วันที่ (Col D)
  serialSheet.getRange(foundRow, 5).setValue(deptName); // แผนก (Col E)
  serialSheet.getRange(foundRow, 6).setValue(user); // ชื่อคนเบิก (Col F)

  // ไปตัดสต็อกในหน้า STOCK (ปรับการค้นหาที่ Col C ตามโครงสร้างใหม่)
  const stockSheet = ss.getSheetByName("STOCK");
  const stockData = stockSheet.getDataRange().getValues();
  let newTotal = 0;

  for (let j = 1; j < stockData.length; j++) {
    if (stockData[j][2].toString().toLowerCase() === matCode.toString().toLowerCase()) {
      newTotal = (Number(stockData[j][5]) || 0) - 1; // คอลัมน์ F (Total)
      const newDeptQty = (Number(stockData[j][colIndex]) || 0) + 1; // คอลัมน์แผนก
      
      stockSheet.getRange(j + 1, 6).setValue(newTotal);
      stockSheet.getRange(j + 1, colIndex + 1).setValue(newDeptQty);
      break;
    }
  }

  writeLog(user, "Withdraw (SN)", matCode, 1, `SN: ${serial} (${deptName})`, newTotal);
  
  let resMsg = `✅ *เบิกสำเร็จ (Serial)*\n📦 รหัส: ${matCode}\n🏷 S/N: \`${serial}\`\n📊 เหลือรวม: ${newTotal}`;
  if (newTotal <= LOW_STOCK_LIMIT) resMsg += `\n\n⚠️ *ALERT: สินค้าใกล้หมดแล้ว!*`;
  return resMsg;
}

// ฟังก์ชันเติมสต็อก
function restock(code, qty, user) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("STOCK");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    // แก้จาก [0] เป็น [2] เพื่อเช็ค Material Code (Q001, Q002...)
    if (data[i][2].toString().toLowerCase() === code.toString().toLowerCase()) {
      const newQty = (Number(data[i][5]) || 0) + qty;
      
      // อัปเดตที่คอลัมน์ F (Index 6)
      sheet.getRange(i+1, 6).setValue(newQty);
      
      writeLog(user, "Restock", code, qty, "STOCK", newQty);
      return `✅ *เติมสต็อกสำเร็จ*\n📦 รายการ: ${data[i][3]}\n➕ เพิ่ม: ${qty}\n📊 ยอดรวมใหม่: ${newQty}`;
    }
  }
  return `❌ ไม่พบรหัสสินค้า "${code}" เพื่อเติมสต็อก`;
}

// ============================================================
// 4. UTILITY FUNCTIONS (ฟังก์ชันช่วยทำงาน)
// ============================================================

function writeLog(user, act, code, qty, note, rem) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName("Logs");
    
    if (!sheet) {
      sheet = ss.insertSheet("Logs");
      sheet.appendRow(["วันเวลา", "ผู้ใช้งาน", "กิจกรรม", "รหัสสินค้า", "จำนวน", "แผนก/หมายเหตุ", "ยอดรวมคงเหลือ"]);
    }

    sheet.appendRow([new Date(), user, act, code, qty, note, rem]);
    SpreadsheetApp.flush(); 
    
    Logger.log("✅ บันทึก Log สำเร็จ: " + code);
  } catch (e) {
    Logger.log("❌ Error writeLog: " + e.toString());
  }
}

function getStockInfo(id) {
  if (!id) return "⚠️ ระบุรหัสสินค้า เช่น `/stock Q001`";
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const data = ss.getSheetByName("STOCK").getDataRange().getValues();
  
  // แก้จาก x[0] เป็น x[2] เพื่อหาตาม Material Code
  const r = data.find(x => x[2] && x[2].toString().trim().toLowerCase() === id.toString().trim().toLowerCase());
  
  if (!r) return "❌ ไม่พบรหัสสินค้า: " + id;

  const name   = r[3] || "ไม่มีชื่อ";
  const total  = (r[5] === "" || isNaN(r[5])) ? 0 : r[5];
  const cbr    = (r[6] === "" || isNaN(r[6])) ? 0 : r[6];
  const ccs    = (r[7] === "" || isNaN(r[7])) ? 0 : r[7];
  const sko    = (r[8] === "" || isNaN(r[8])) ? 0 : r[8];
  const ryg    = (r[9] === "" || isNaN(r[9])) ? 0 : r[9];
  const trt    = (r[10] === "" || isNaN(r[10])) ? 0 : r[10]; 
  const status = r[11] || "-";

  return `📦 *${id.toUpperCase()} - ${name}*\n\n` +
         `📊 *คงเหลือรวม: ${total}*\n` +
         `🔔 สถานะ: ${status}\n` +
         `------------------\n` +
         `🏢 CBR: ${cbr} | CCS: ${ccs}\n` +
         `🏢 SKO: ${sko} | RYG: ${ryg}\n` +
         `🏢 TRT: ${trt}`; 
}

function sendTelegram(chatId, text) {
  const payload = { chat_id: chatId, text: text, parse_mode: "Markdown" };
  const options = { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };
  UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`, options);
}

// แจ้งเตือนเมื่อแก้ชีตโดยตรง
function onEdit(e) {
  try {
    const range = e.range;
    const sheetName = range.getSheet().getName();
    
    const skipSheets = ["Logs", "Summary", "Config", "AI_Memory"];
    if (skipSheets.includes(sheetName)) return;

    const oldVal = e.oldValue || "ว่าง";
    const newVal = e.value    || "ว่าง";
    const user   = Session.getActiveUser().getEmail() || "Unknown";

    const msg = `⚠️ *มีการแก้ไขข้อมูลโดยตรง!*\n` +
                `📍 ชีต: ${sheetName}\n` +
                `📌 เซลล์: ${range.getA1Notation()}\n` +
                `🔄 เดิม: ${oldVal}\n` +
                `✏️ ใหม่: ${newVal}\n` +
                `👤 โดย: ${user}`;
    sendTelegram(CHAT_ID, msg);
  } catch (err) {
    Logger.log("onEdit Error: " + err.toString());
  }
}


// ============================================================
// AI Gemini (Version: Real-time Stock Sync)
// ============================================================
function callGemini(msg, chatId) {
  const API_KEY = GEMINI_API_KEY;
  const MODEL = "gemini-1.5-flash";
  
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const stockSheet = ss.getSheetByName("STOCK");
    
    // 1. ดึงข้อมูลแค่คอลัมน์ที่จำเป็น (B:F) และเอาแค่ 100 แถวแรกพอ
    const dataRange = stockSheet.getRange("B2:F100").getValues();
    
    // 2. กรองเอาเฉพาะรายการที่มีชื่อสินค้า และแปลงเป็นข้อความสั้นๆ
    let stockSummary = dataRange
      .filter(r => r[0] && r[2]) // กรองแถวว่าง
      .map(r => `${r[2]} (รหัส: ${r[0]}) คงเหลือ: ${r[4]}`) // คอลัมน์ D คือชื่อ, F คือยอดรวม
      .join("\n");

    // 3. จัดการ Memory (ล้างประวัติที่เสีย)
    let memSheet = ss.getSheetByName("AI_Memory") || ss.insertSheet("AI_Memory");
    let history = [];
    const memData = memSheet.getDataRange().getValues();
    let userRow = -1;

    for (let i = 1; i < memData.length; i++) {
      if (String(memData[i][0]) === String(chatId)) {
        userRow = i + 1;
        try {
          history = JSON.parse(memData[i][1]);
        } catch (e) { history = []; }
        break;
      }
    }

    // 4. สร้าง Payload โดยคุมเข้มไม่ให้ยาวเกินไป
    const systemPrompt = `คุณคือบอทเช็คสต็อกของ SYS ตอบเป็นภาษาไทยสั้นๆ ข้อมูลปัจจุบัน:\n${stockSummary}`;
    
    const payload = {
      contents: [
        { role: "user", parts: [{ text: msg }] }
      ],
      systemInstruction: { parts: [{ text: systemPrompt }] },
      generationConfig: { maxOutputTokens: 200 }
    };

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1/models/${MODEL}:generateContent?key=${API_KEY}`, options);
    const resText = response.getContentText();
    const resCode = response.getResponseCode();

    if (resCode !== 200) {
      // ถ้าพังอีก ให้ล้าง Memory ของคนนี้ทิ้งทันที
      if (userRow !== -1) memSheet.getRange(userRow, 2).setValue("[]");
      console.error("Gemini Error: " + resText);
      return `🤖 AI มึนหัว (Error ${resCode}) รบกวนเฮียลองพิมพ์ใหม่อีกทีครับ`;
    }

    const json = JSON.parse(resText);
    const aiResponse = json.candidates[0].content.parts[0].text;

    // 5. บันทึกประวัติ (เก็บแค่ 3 คู่ล่าสุด)
    history.push({ role: "user", text: msg }, { role: "model", text: aiResponse });
    const limitedHistory = JSON.stringify(history.slice(-6));

    if (userRow !== -1) {
      memSheet.getRange(userRow, 2).setValue(limitedHistory);
    } else {
      memSheet.appendRow([String(chatId), limitedHistory, new Date()]);
    }

    return aiResponse;

  } catch (e) {
    return "🤖 ระบบขัดข้อง: " + e.toString();
  }
}

// ทดสอบ
function testAi() {
  // จำลองว่าเราคือ User พิมพ์หาบอท
  const testMessage = "สวัสดีครับบอท เช็คสต็อกให้หน่อย";
  const testChatId = "12345"; 
  
  const response = callGemini(testMessage, testChatId);
  Logger.log("🤖 ผลการทดสอบ: " + response);
}

function setWebhook() {
  const webAppUrl = "https://script.google.com/macros/s/AKfycbyuwxAzdc-GU1vWzuB8kTg737TBola4vBsMU380s9_fiHAoCYprjbvFxifIUDSnH0C7/exec";
  const response = UrlFetchApp.fetch("https://api.telegram.org/bot" + TELEGRAM_TOKEN + "/setWebhook?url=" + webAppUrl);
  Logger.log("🎯 ผลการเชื่อมต่อ: " + response.getContentText());
}


// เพิ่ม
// ============================================================
// SYNC MATERIAL CODE: External Spare_TypeC → STOCK Sheet
// ============================================================

// ID ของ Spare_TypeC Google Sheet (จากลิงก์ที่ส่งมา)
const SPARE_TYPE_C_SHEET_ID = "1dZXwYcPWc0OgWriXnA5ZxbAui8GmJDLRddrKTlZPuic";

/**
 * ซิงค์ Material Code จาก Spare_TypeC (Google Sheet ภายนอก) ไปยัง STOCK sheet
 * จะเอา Material Code (เช่น 2000039121) แทน Q001, Q002, ...
 */
function syncMaterialCodeFromSpareTypeC() {
  try {
    const stockSS = SpreadsheetApp.openById(SHEET_ID);
    const stockSheet = stockSS.getSheetByName("STOCK");
    const spareSS = SpreadsheetApp.openById(SPARE_TYPE_C_SHEET_ID);
    const spareTypeC = spareSS.getSheetByName("Spare Type C");
    
    // ดึงข้อมูลทั้งหมดมา
    const spareData = spareTypeC.getDataRange().getValues();
    
    // --- จุดสำคัญ: เริ่มวนลูปจากแถวที่รหัสสินค้าเริ่มปรากฏ ---
    // สมมติในไฟล์ Spare_TypeC รหัสเริ่มที่แถวที่ 5 (Index คือ 4)
    // และรหัสอยู่ใน Column C (Index คือ 2)
    const START_ROW_SPARE = 4; // แถวที่ 5
    const MATERIAL_COL_INDEX = 1; // Column C
    
    let syncCount = 0;
    
    for (let i = START_ROW_SPARE; i < spareData.length; i++) {
      const materialCode = spareData[i][MATERIAL_COL_INDEX];
      
      if (materialCode && materialCode.toString().trim() !== "") {
        // คำนวณแถวที่จะไปวางใน STOCK (เริ่มวางที่แถว 2 เป็นต้นไป)
        const targetRow = (i - START_ROW_SPARE) + 2; 
        
        // เขียนลง Column C ของหน้า STOCK
        stockSheet.getRange(targetRow, 3).setValue(materialCode.toString().trim());
        syncCount++;
      }
    }
    
    SpreadsheetApp.flush();
    Logger.log(`✅ ซิงค์สำเร็จทั้งหมด ${syncCount} รายการ`);
    
  } catch (err) {
    Logger.log("❌ Error: " + err.toString());
  }
}

/**
 * Preview ทำงาน ก่อนทำการเปลี่ยนจริง
 */
function previewMaterialCodeSync() {
  try {
    const spareSS = SpreadsheetApp.openById(SPARE_TYPE_C_SHEET_ID);
    const spareTypeC = spareSS.getSheetByName("Spare Type C");
    
    if (!spareTypeC) {
      Logger.log("❌ ไม่พบ Spare Type C sheet");
      return;
    }
    
    const spareData = spareTypeC.getDataRange().getValues();
    const stockSS = SpreadsheetApp.openById(SHEET_ID);
    const stockData = stockSS.getSheetByName("STOCK").getDataRange().getValues();
    
    // หาคอลัมน์ Material Code
    let materialCodeColIndex = -1;
    for (let col = 0; col < spareData[0].length; col++) {
      const header = spareData[0][col] ? spareData[0][col].toString().trim().toLowerCase() : "";
      if (header.includes("material code")) {
        materialCodeColIndex = col;
        break;
      }
    }
    
    if (materialCodeColIndex === -1) {
      Logger.log("❌ ไม่พบ Material Code column");
      return;
    }
    
    Logger.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    Logger.log("📋 Preview: Spare_TypeC → STOCK");
    Logger.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    
    for (let i = 1; i < Math.min(spareData.length, 15); i++) {
      const spareMatCode = spareData[i][materialCodeColIndex] ? spareData[i][materialCodeColIndex].toString().trim() : "";
      
      const stockMatCode = stockData[i] ? (stockData[i][2] ? stockData[i][2].toString().trim() : "") : "";
      
      if (spareMatCode) {
        Logger.log(`Row ${i + 1}:`);
        Logger.log(`  Spare_TypeC: ${spareMatCode}`);
        Logger.log(`  Current STOCK: ${stockMatCode}`);
        Logger.log(`  → จะเปลี่ยน: ${stockMatCode} → ${spareMatCode}`);
        Logger.log("");
      }
    }
    
    Logger.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    Logger.log("✅ ถ้าถูกต้อง ให้รัน: syncMaterialCodeFromSpareTypeC()");
    Logger.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    
  } catch (err) {
    Logger.log("❌ Error: " + err.toString());
  }
}

/**
 * Undo: กลับไปเป็น Q001, Q002, Q003 เดิม
 */
function undoMaterialCodeSync() {
  try {
    const stockSS = SpreadsheetApp.openById(SHEET_ID);
    const stockSheet = stockSS.getSheetByName("STOCK");
    
    if (!stockSheet) {
      Logger.log("❌ ไม่พบ STOCK sheet");
      return;
    }
    
    const data = stockSheet.getDataRange().getValues();
    
    Logger.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    Logger.log("🔙 กำลัง Undo...");
    Logger.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    
    for (let i = 1; i < data.length; i++) {
      const oldCode = `Q${String(i).padStart(3, '0')}`;
      stockSheet.getRange(i + 1, 3).setValue(oldCode); // Column C
      
      if (i <= 10) {
        Logger.log(`✅ Row ${i + 1}: ${oldCode}`);
      }
    }
    
    SpreadsheetApp.flush();
    
    Logger.log("...");
    Logger.log(`✅ Undo เสร็จ!`);
    Logger.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    
  } catch (err) {
    Logger.log("❌ Error: " + err.toString());
  }
}

/**
 * ทดสอบการเชื่อมต่อ
 */
function testConnection() {
  try {
    Logger.log("🧪 ทดสอบการเชื่อมต่อ...");
    
    // Test STOCK sheet
    const stockSS = SpreadsheetApp.openById(SHEET_ID);
    const stockSheet = stockSS.getSheetByName("STOCK");
    Logger.log(`✅ เชื่อมต่อ STOCK sheet สำเร็จ (${stockSheet.getLastRow()} rows)`);
    
    // Test Spare_TypeC sheet
    const spareSS = SpreadsheetApp.openById(SPARE_TYPE_C_SHEET_ID);
    const spareTypeC = spareSS.getSheetByName("Spare Type C");
    Logger.log(`✅ เชื่อมต่อ Spare Type C สำเร็จ (${spareTypeC.getLastRow()} rows)`);
    
    Logger.log("✅ ทดสอบสำเร็จ!");
    
  } catch (err) {
    Logger.log("❌ Error: " + err.toString());
  }
}

function sendMenu(chatId) {
  const payload = {
    chat_id: chatId,
    text: "📦 *กรุณาเลือกแผนกที่ต้องการเบิกสินค้า:*",
    parse_mode: "Markdown",
    reply_markup: {
      keyboard: [
        [{ text: "/dp-cbr" }, { text: "/dp-ccs" }],
        [{ text: "/dp-sko" }, { text: "/dp-ryg" }, { text: "/dp-trt" }],
        [{ text: "/stock" }, { text: "/help" }]
      ],
      resize_keyboard: true,
      one_time_keyboard: false
    }
  };
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };
  UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`, options);
}
