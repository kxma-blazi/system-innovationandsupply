// ============================================================
// 1. WEB DASHBOARD SERVER-SIDE
// ============================================================

function doGet(e) {
  const userEmail = Session.getActiveUser().getEmail() || "Guest";

  // บันทึก Log เมื่อมีคนเปิดหน้าเว็บ
  writeLog(userEmail, "เปิด Dashboard", "-", 0, "Access via Web", "-");

  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("SYS Stock Dashboard")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getStockRows(q) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("STOCK");
  if (!sheet) return [];

  const data = sheet.getDataRange().getDisplayValues();
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    let matCode = data[i][1] ? String(data[i][1]).trim() : "";

    if (
      !matCode ||
      matCode.toLowerCase().includes("cellimage") ||
      matCode === "ID"
    ) {
      continue;
    }

    let qty = !isNaN(data[i][5]) && data[i][5] !== "" ? Number(data[i][5]) : 0;

    // คำนวณยอดเบิกสะสมจริง โดยรวมจากคอลัมน์ G ถึง K (Index 6 ถึง 10)
    let usageCount = 0;
    for (let col = 6; col <= 10; col++) {
      let val = data[i][col];
      usageCount += !isNaN(val) && val !== "" ? Number(val) : 0;
    }

    rows.push({
      code: matCode,
      desc: data[i][3] || "-",
      qty: qty,
      usage: usageCount, // ส่งค่ารวมยอดใช้งานจริงไปที่หน้าเว็บ
      cbr: Number(data[i][6]) || 0,
      ccs: Number(data[i][7]) || 0,
      sko: Number(data[i][8]) || 0,
      ryg: Number(data[i][9]) || 0,
      trt: Number(data[i][10]) || 0,
    });
  }
  return rows;
}

// ============================================================
// 2. SYSTEM LOGGING (ฟังก์ชันบันทึก Log )
// ============================================================

function writeLog(user, act, code, qty, note, rem) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName("Logs");

    if (!sheet) {
      sheet = ss.insertSheet("Logs");
      sheet.appendRow([
        "วันเวลา",
        "ผู้ใช้งาน",
        "กิจกรรม",
        "รหัสสินค้า",
        "จำนวน",
        "แผนก/หมายเหตุ",
        "ยอดรวมคงเหลือ",
      ]);
      sheet
        .getRange("A1:G1")
        .setBackground("#252b38")
        .setFontColor("#ffffff")
        .setFontWeight("bold");
    }

    sheet.appendRow([new Date(), user, act, code, qty, note, rem]);
    SpreadsheetApp.flush();
    Logger.log("✅ บันทึก Log สำเร็จ: " + act);
  } catch (e) {
    Logger.log("❌ Error writeLog: " + e.toString());
  }
}

// ฟังก์ชันสำหรับค้นหาและดึงข้อมูลไปแสดงบนตารางหน้าเว็บ
function getStockAlerts() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const stockSheet = ss.getSheetByName("STOCK");
    const data = stockSheet.getDataRange().getDisplayValues();
    const alerts = [];
    const LIMIT = 5;

    for (let i = 1; i < data.length; i++) {
      // ดึงรหัสสินค้าและตัดช่องว่าง
      let matCode = data[i][1] ? String(data[i][1]).trim() : "";
      if (matCode === "" || matCode.toLowerCase().includes("cellimage"))
        continue;

      // ตรวจสอบจำนวนสินค้า ถ้าไม่ใช่ตัวเลขให้เป็น 0
      let rawQty = data[i][5];
      let qty = !isNaN(rawQty) && rawQty !== "" ? Number(rawQty) : 0;

      if (qty <= 0) {
        alerts.push({
          code: matCode,
          name: data[i][3],
          qty: qty,
          level: "critical",
          icon: "🔴",
        });
      } else if (qty <= LIMIT) {
        alerts.push({
          code: matCode,
          name: data[i][3],
          qty: qty,
          level: "warning",
          icon: "🟡",
        });
      }
    }

    // เรียงตามจำนวนจากน้อยไปมาก
    return alerts.sort((a, b) => a.qty - b.qty);
  } catch (e) {
    // ถ้ามี error คืนค่าเป็น array ว่าง
    return [];
  }
}

// 🚩 ฟังก์ชันสร้างแถว ตรวจสอบว่าใช้ r.code
function buildHtmlRows(rows) {
  return rows
    .map((r) => {
      const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${encodeURIComponent(r.code)}`;
      const statusClass =
        r.qty <= 0
          ? "status-red"
          : r.qty <= 5
            ? "status-yellow"
            : "status-green";

      return `<tr class="${statusClass}">
      <td data-label="ID"><b>${r.code}</b></td> 
      <td data-label="QR" style="text-align:center;"><img src="${qrUrl}" width="40"></td>
      <td data-label="Description">${r.desc}</td>
      <td data-label="QTY">${r.qty}</td>
      <td>${r.cbr}</td><td>${r.ccs}</td><td>${r.sko}</td><td>${r.ryg}</td><td>${r.trt}</td>
      <td data-label="Status"><span class="badge">${r.qty <= 0 ? "หมด" : "ปกติ"}</span></td>
    </tr>`;
    })
    .join("");
}

/**
 * ดึงสถิติแต่ละแผนก
 */
function getDepartmentStats() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const stockSheet = ss.getSheetByName("STOCK");
    const rawData = stockSheet
      .getRange(1, 1, stockSheet.getLastRow(), 11)
      .getValues();

    const DEPTS = {
      CBR: 6,
      CCS: 7,
      SKO: 8,
      RYG: 9,
      TRT: 10,
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
        efficiency:
          itemsWithStock > 0
            ? Math.round((itemsWithStock / rawData.length) * 100)
            : 0,
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
    const rawData = stockSheet
      .getRange(1, 1, stockSheet.getLastRow(), 6)
      .getValues();

    const items = [];

    for (let i = 1; i < rawData.length; i++) {
      let matCode = "";
      let codeValue = rawData[i][1];

      if (codeValue && String(codeValue).includes("CellImage")) {
        codeValue = rawData[i][2];
      }

      if (codeValue) {
        matCode =
          typeof codeValue === "number"
            ? codeValue.toFixed(0)
            : String(codeValue).trim();
      }

      if (matCode && !matCode.includes("CellImage")) {
        items.push({
          code: matCode,
          name: rawData[i][3],
          qty: Number(rawData[i][5]),
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
      avgPerDay: Math.round(total / 7),
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
      extracted =
        typeof colB === "number" ? colB.toFixed(0) : String(colB).trim();
    } else if (colC && !String(colC).includes("CellImage")) {
      extracted =
        typeof colC === "number" ? colC.toFixed(0) : String(colC).trim();
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

  const recentWithdraws = logData.filter(
    (r) => new Date(r[0]) > thirtyDaysAgo && r[2] === "Withdraw",
  );

  return {
    totalWithdraws: recentWithdraws.length,
    avgPerDay: (recentWithdraws.length / 30).toFixed(1),
    topItems: getTopWithdrawItems(recentWithdraws),
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
      records: data
        .slice(1)
        .filter((r) => r[2] !== "") // กรองเฉพาะแถวที่มีซีเรียลใน Col C (นำจ่ายแล้ว)
        .reverse()
        .slice(0, limit)
        .map((r) => ({
          code: r[0], // A: Code
          serial: r[2], // C: Serial นำจ่าย
          dateDelivered: r[3]
            ? Utilities.formatDate(new Date(r[3]), "GMT+7", "dd/MM/yy HH:mm")
            : "-", // D: วันที่
          note: r[4] || "-", // E: แผนก
          deliveredBy: r[5] || "-", // F: ผู้เบิก
        })),
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
      const welcomeMsg =
        "📦 *ระบบจัดการสต็อก SYS (Auto-Detect)*\n\n" +
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
        sendTelegram(
          chatId,
          `⚠️ กรุณาระบุรหัสสินค้าหรือซีเรียล\nเช่น: \`${cmd} SN999\` หรือ \`${cmd} Q001\``,
        );
        return;
      }

      const colIdx = DEPT_MAP[cmd];
      const deptName = cmd.replace("/dp-", "").toUpperCase();

      // --- ขั้นตอนที่ 1: ลองเบิกแบบซีเรียลก่อน ---
      const serialResult = withdrawBySerial(input, colIdx, deptName, user);

      // ถ้าเจอซีเรียล (เบิกสำเร็จ หรือ ถูกเบิกไปแล้ว)
      if (serialResult !== "NOT_FOUND") {
        sendTelegram(chatId, serialResult);
        // ส่งแจ้งเตือนเข้ากลุ่ม (CHAT_ID จาก Config)
        if (typeof CHAT_ID !== "undefined" && !serialResult.includes("❌")) {
          sendTelegram(
            CHAT_ID,
            `📣 *บันทึก (SN):* ${serialResult}\n(โดย: ${user})`,
          );
        }
        return;
      }

      // --- ขั้นตอนที่ 2: ถ้าไม่ใช่ซีเรียล ให้เบิกแบบจำนวน (QTY) ---
      let qtyMatch = (args[2] || "1").match(/\d+/);
      let qty = qtyMatch ? parseInt(qtyMatch[0], 10) : 1;

      const result = withdraw(input, qty, colIdx, deptName, user);
      sendTelegram(chatId, result);

      if (typeof CHAT_ID !== "undefined" && !result.includes("❌")) {
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
        sendTelegram(
          chatId,
          "⚠️ รูปแบบผิด! ต้องเป็น: `/restock [รหัส] [จำนวน]`",
        );
      } else {
        let qtyMatch = args[2].match(/\d+/);
        let qty = qtyMatch ? parseInt(qtyMatch[0], 10) : NaN;
        sendTelegram(chatId, restock(args[1], qty, user));
      }
    }

    // 5. ถ้าไม่ใช่คำสั่ง ให้ส่งไปหา Groq AI
    else {
      sendTelegram(chatId, callGroq(text, chatId));
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
      if (
        data[i][1].toString().toLowerCase().trim() ===
        code.toString().toLowerCase().trim()
      ) {
        const currentTotal = Number(data[i][5]);
        if (currentTotal < qty)
          return `❌ ${code} ของไม่พอ (คงเหลือ ${currentTotal})`;

        const newTotal = currentTotal - qty;
        const newDeptTotal = (Number(data[i][colIndex]) || 0) + qty;

        sheet.getRange(i + 1, 6).setValue(newTotal); // Col F
        sheet.getRange(i + 1, colIndex + 1).setValue(newDeptTotal);

        SpreadsheetApp.flush(); // บังคับเขียนข้อมูลทันที
        writeLog(user, "Withdraw", code, qty, deptName, newTotal);

        let resMsg = `✅ *เบิกสำเร็จ*\n📦 รายการ: ${data[i][3]}\n📍 แผนก: ${deptName}\n📉 จำนวน: -${qty}\n📊 เหลือรวม: ${newTotal}`;
        if (newTotal <= LOW_STOCK_LIMIT)
          resMsg += `\n\n⚠️ *ALERT: สินค้าใกล้หมดแล้ว!*`;
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

  // 🔍 ค้นหา Serial
  for (let i = 1; i < sData.length; i++) {
    const colB = (sData[i][1] || "").toString().trim(); // Serial ว่าง
    const colC = (sData[i][2] || "").toString().trim(); // Serial ที่ถูกเบิกแล้ว

    // 1. เช็คว่าถูกเบิกไปแล้ว
    if (colC === serial) {
      takerInfo = {
        matCode: sData[i][0],
        date: sData[i][3],
        name: sData[i][5],
      };
      break;
    }

    // 2. เช็คว่ายังว่างอยู่
    if (colB === serial) {
      foundRow = i + 1;
      matCode = sData[i][0];
      break;
    }
  }

  // ❌ เบิกซ้ำ
  if (takerInfo) {
    return `⚠️ *ซีเรียลนี้ถูกเบิกไปแล้ว!*
👤 โดย: ${takerInfo.name}
📅 เมื่อ: ${Utilities.formatDate(new Date(takerInfo.date), "GMT+7", "dd/MM/yy HH:mm")}
📦 รหัสสินค้า: ${takerInfo.matCode}`;
  }

  // ❌ ไม่เจอ
  if (foundRow === -1) return "NOT_FOUND";

  // ✅ ทำการเบิก SERIAL_DATA
  const now = new Date();
  serialSheet.getRange(foundRow, 2).clearContent(); // B
  serialSheet.getRange(foundRow, 3).setValue(serial); // C
  serialSheet.getRange(foundRow, 4).setValue(now); // D
  serialSheet.getRange(foundRow, 5).setValue(deptName); // E
  serialSheet.getRange(foundRow, 6).setValue(user); // F

  // ======================================================
  // 🔥 ตัด STOCK (จุดสำคัญสุด)
  // ======================================================
  const stockSheet = ss.getSheetByName("STOCK");
  const stockData = stockSheet.getDataRange().getValues();

  let newTotal = 0;
  let foundStock = false;

  for (let j = 1; j < stockData.length; j++) {
    const stockCode = (stockData[j][1] || "").toString().trim().toLowerCase(); // ✅ Column B
    const targetCode = matCode.toString().trim().toLowerCase();

    if (stockCode === targetCode) {
      const currentTotal = Number(stockData[j][5]) || 0;

      if (currentTotal <= 0) {
        return `❌ สินค้า ${matCode} หมดสต็อกแล้ว`;
      }

      newTotal = currentTotal - 1;
      const newDeptQty = (Number(stockData[j][colIndex]) || 0) + 1;

      stockSheet.getRange(j + 1, 6).setValue(newTotal); // F
      stockSheet.getRange(j + 1, colIndex + 1).setValue(newDeptQty);

      foundStock = true;
      break;
    }
  }

  // ❌ กันเคสหาไม่เจอใน STOCK
  if (!foundStock) {
    return `❌ ไม่พบรหัสสินค้า ${matCode} ใน STOCK`;
  }

  // ✅ Log
  writeLog(
    user,
    "Withdraw (SN)",
    matCode,
    1,
    `SN: ${serial} (${deptName})`,
    newTotal,
  );

  // ✅ Response
  let resMsg = `✅ *เบิกสำเร็จ (Serial)*
📦 รหัส: ${matCode}
🏷 S/N: \`${serial}\`
📊 เหลือรวม: ${newTotal}`;

  if (newTotal <= LOW_STOCK_LIMIT) {
    resMsg += `\n\n⚠️ *ALERT: สินค้าใกล้หมดแล้ว!*`;
  }

  return resMsg;
}

// ฟังก์ชันเติมสต็อก
function restock(code, qty, user) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("STOCK");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    // แก้จาก [0] เป็น [2] เพื่อเช็ค Material Code (Q001, Q002...)
    if (data[i][1].toString().toLowerCase() === code.toString().toLowerCase()) {
      const newQty = (Number(data[i][5]) || 0) + qty;

      // อัปเดตที่คอลัมน์ F (Index 6)
      sheet.getRange(i + 1, 6).setValue(newQty);

      writeLog(user, "Restock", code, qty, "STOCK", newQty);
      return `✅ *เติมสต็อกสำเร็จ*\n📦 รายการ: ${data[i][3]}\n➕ เพิ่ม: ${qty}\n📊 ยอดรวมใหม่: ${newQty}`;
    }
  }
  return `❌ ไม่พบรหัสสินค้า "${code}" เพื่อเติมสต็อก`;
}

// ============================================================
// 4. UTILITY FUNCTIONS (ฟังก์ชันช่วยทำงาน)
// ============================================================

function getStockInfo(id) {
  if (!id) return "⚠️ ระบุรหัสสินค้า เช่น `/stock Q001`";
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const data = ss.getSheetByName("STOCK").getDataRange().getValues();

  // แก้จาก x[0] เป็น x[2] เพื่อหาตาม Material Code
  const r = data.find(
    (x) =>
      x[2] &&
      x[2].toString().trim().toLowerCase() ===
        id.toString().trim().toLowerCase(),
  );

  if (!r) return "❌ ไม่พบรหัสสินค้า: " + id;

  const name = r[3] || "ไม่มีชื่อ";
  const total = r[5] === "" || isNaN(r[5]) ? 0 : r[5];
  const cbr = r[6] === "" || isNaN(r[6]) ? 0 : r[6];
  const ccs = r[7] === "" || isNaN(r[7]) ? 0 : r[7];
  const sko = r[8] === "" || isNaN(r[8]) ? 0 : r[8];
  const ryg = r[9] === "" || isNaN(r[9]) ? 0 : r[9];
  const trt = r[10] === "" || isNaN(r[10]) ? 0 : r[10];
  const status = r[11] || "-";

  return (
    `📦 *${id.toUpperCase()} - ${name}*\n\n` +
    `📊 *คงเหลือรวม: ${total}*\n` +
    `🔔 สถานะ: ${status}\n` +
    `------------------\n` +
    `🏢 CBR: ${cbr} | CCS: ${ccs}\n` +
    `🏢 SKO: ${sko} | RYG: ${ryg}\n` +
    `🏢 TRT: ${trt}`
  );
}

function sendTelegram(chatId, text) {
  const payload = { chat_id: chatId, text: text, parse_mode: "Markdown" };
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };
  UrlFetchApp.fetch(
    `https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`,
    options,
  );
}

// แจ้งเตือนเมื่อแก้ชีต เเละ เก็บ Logs
function onEdit(e) {
  try {
    const range = e.range;
    const sheetName = range.getSheet().getName();

    const skipSheets = ["Logs", "Summary", "Config", "AI_Memory"];
    if (skipSheets.includes(sheetName)) return;

    const oldVal = e.oldValue || "ว่าง";
    const newVal = e.value || "ว่าง";
    const user = Session.getActiveUser().getEmail() || "Unknown";

    // ✅ 1. บันทึกลง Logs Sheet
    const ss = range.getSheet().getParent();
    const logSheet = ss.getSheetByName("Logs") || ss.insertSheet("Logs");
    logSheet.appendRow([
      new Date(), // เวลาแก้ไข
      user, // ผู้แก้ไข
      sheetName, // ชื่อชีต
      range.getA1Notation(), // เซลล์
      oldVal, // ค่าเก่า
      newVal, // ค่าใหม่
    ]);

    const msg =
      `⚠️ *มีการแก้ไขข้อมูลโดยตรง!*\n` +
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

function updateStock(rowNum, newData) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("STOCK");
    const logSheet = ss.getSheetByName("Logs") || ss.insertSheet("Logs");

    const oldValues = sheet
      .getRange(rowNum, 1, 1, sheet.getLastColumn())
      .getValues()[0];

    // อัปเดตค่าใหม่
    const newValues = [
      newData.code || oldValues[1],
      oldValues[2], // สมมติคอลัมน์ C ไม่เปลี่ยน
      newData.desc || oldValues[3],
      newData.qty !== undefined ? newData.qty : oldValues[5],
      newData.cbr !== undefined ? newData.cbr : oldValues[6],
      newData.ccs !== undefined ? newData.ccs : oldValues[7],
      newData.sko !== undefined ? newData.sko : oldValues[8],
      newData.ryg !== undefined ? newData.ryg : oldValues[9],
      newData.trt !== undefined ? newData.trt : oldValues[10],
    ];

    sheet.getRange(rowNum, 1, 1, newValues.length).setValues([newValues]);

    // บันทึก Log
    logSheet.appendRow([
      new Date(),
      Session.getActiveUser().getEmail() || "Unknown",
      "STOCK",
      `Row ${rowNum}`,
      JSON.stringify(oldValues),
      JSON.stringify(newValues),
    ]);

    // ส่ง Telegram
    const msg = `✏️ *Stock Updated*\nRow: ${rowNum}\nOld: ${JSON.stringify(oldValues)}\nNew: ${JSON.stringify(newValues)}`;
    sendTelegram(CHAT_ID, msg);

    return true;
  } catch (err) {
    Logger.log("updateStock Error: " + err.toString());
    return false;
  }
}

function testGeminiKey() {
  const API_KEY = GEMINI_API_KEY; // ใส่ key ใหม่ใน Config.gs ก่อน
  const payload = {
    contents: [{ role: "user", parts: [{ text: "hello" }] }],
  };
  const res = UrlFetchApp.fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-lite:generateContent?key=${API_KEY}`,
    {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    },
  );
  Logger.log(
    res.getResponseCode() + " → " + res.getContentText().slice(0, 300),
  );
}

/*
// ============================================================
// AI Gemini (Version: Real-time Stock Sync)
// ============================================================
function callGemini(msg, chatId) {
  const API_KEY = GEMINI_API_KEY;
  const MODEL   = "gemini-2.0-flash-lite"; // Model

  try {
    const ss         = SpreadsheetApp.openById(SHEET_ID);
    const stockSheet = ss.getSheetByName("STOCK");

    // ── 1. ดึงข้อมูล Stock ───────────────────────────────────────────────────
    const dataRange = stockSheet.getRange("B2:F100").getValues();
    const stockSummary = dataRange
      .filter(r => r[0] && !String(r[0]).includes("CellImage") && r[2])
      .slice(0, 40)
      .map(r => {
        const code = typeof r[0] === "number" ? r[0].toFixed(0) : String(r[0]).trim();
        const name = String(r[2] || "").trim();
        const qty  = Number(r[4]) || 0;
        return `${code} | ${name} | คงเหลือ: ${qty}`;
      })
      .join("\n");

    // ── 2. โหลด Memory ───────────────────────────────────────────────────────  ← แทนที่ตรงนี้ทั้งบล็อก
    const memSheet = ss.getSheetByName("AI_Memory") || ss.insertSheet("AI_Memory");
    const memData  = memSheet.getDataRange().getValues();
    let history    = [];
    let userRow    = -1;

    for (let i = 1; i < memData.length; i++) {
      if (String(memData[i][0]) === String(chatId)) {
        userRow = i + 1;
        try {
          const parsed = JSON.parse(memData[i][1]);
          const isValid = Array.isArray(parsed) && parsed.every(h =>
            h.role &&
            (h.role === "user" || h.role === "model") &&
            h.text &&
            String(h.text).trim().length > 0
          );
          history = isValid ? parsed : [];
          if (!isValid) {
            Logger.log("⚠️ พบ history เสีย — reset แล้ว");
            memSheet.getRange(userRow, 2).setValue("[]");
          }
        } catch {
          history = [];
          memSheet.getRange(userRow, 2).setValue("[]");
        }
        break;
      }
    }

    Logger.log("📤 history ที่จะส่ง: " + JSON.stringify(history));
    Logger.log("📤 msg: " + msg);

    // ── 3. สร้าง Payload ─────────────────────────────────────────────────────
    const systemPrompt =
      `คุณคือบอทเช็คสต็อกของ SYS ตอบเป็นภาษาไทยสั้นๆ กระชับ ไม่เกิน 3 บรรทัด\n\nข้อมูลสต็อกปัจจุบัน:\n${stockSummary}`;

    const historyContents = history
      .filter(h => h.role && h.text && String(h.text).trim().length > 0)
      .map(h => ({
        role:  h.role === "model" ? "model" : "user",
        parts: [{ text: String(h.text).trim() }]
      }));

    const payload = {
      contents: [
        ...historyContents,
        { role: "user", parts: [{ text: msg }] }
      ],
      systemInstruction: {
        role:  "system",
        parts: [{ text: systemPrompt }]
      },
      generationConfig: {
        maxOutputTokens: 200,
        temperature:     0.7
      }
    };

    // ── 4. เรียก API ─────────────────────────────────────────────────────────
    const options = {
      method:             "post",
      contentType:        "application/json",
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent?key=${API_KEY}`,
      options
    );

    const resCode = response.getResponseCode();
    const resText = response.getContentText();

    // ── 5. จัดการ Error ──────────────────────────────────────────────────────
    if (resCode !== 200) {
      Logger.log(`❌ Gemini Error ${resCode}: ${resText}`);
      if (userRow !== -1) memSheet.getRange(userRow, 2).setValue("[]");
      let errMsg = "";
      try { errMsg = JSON.parse(resText)?.error?.message || resText; } catch { errMsg = resText; }
      Logger.log("📋 Error detail: " + errMsg);
      return `🤖 ระบบขัดข้อง (${resCode}) กรุณาลองใหม่ครับ`;
    }

    // ── 6. ดึงคำตอบ ──────────────────────────────────────────────────────────
    const json = JSON.parse(resText);
    const aiResponse = json?.candidates?.[0]?.content?.parts?.[0]?.text || "ไม่สามารถดึงคำตอบได้";

    // ── 7. บันทึก Memory ─────────────────────────────────────────────────────
    history.push(
      { role: "user",  text: msg },
      { role: "model", text: aiResponse }
    );
    const limitedHistory = JSON.stringify(history.slice(-6));

    if (userRow !== -1) {
      memSheet.getRange(userRow, 2).setValue(limitedHistory);
    } else {
      memSheet.appendRow([String(chatId), limitedHistory, new Date()]);
    }

    return aiResponse;

  } catch (e) {
    Logger.log("❌ callGemini Exception: " + e.toString());
    return "🤖 ระบบขัดข้อง: " + e.toString();
  }
}
*/

/*
// ทดสอบ
function testAi() {
  // จำลองว่าเราคือ User พิมพ์หาบอท
  const testMessage = "สวัสดีครับบอท เช็คสต็อกให้หน่อย";
  const testChatId = "12345";

  const response = callGemini(testMessage, testChatId);
  Logger.log("🤖 ผลการทดสอบ: " + response);
}
*/

function testAi() {
  const response = callGroq("สวัสดีครับ", "12345");
  Logger.log("🤖 ผลการทดสอบ: " + response);
}

function setWebhook() {
  const webAppUrl =
    "https://script.google.com/macros/s/AKfycbyuwxAzdc-GU1vWzuB8kTg737TBola4vBsMU380s9_fiHAoCYprjbvFxifIUDSnH0C7/exec";
  const response = UrlFetchApp.fetch(
    "https://api.telegram.org/bot" +
      TELEGRAM_TOKEN +
      "/setWebhook?url=" +
      webAppUrl,
  );
  Logger.log("🎯 ผลการเชื่อมต่อ: " + response.getContentText());
}

// ============================================================
// SYNC MATERIAL CODE: External Spare_TypeC → STOCK Sheet
// ============================================================

// ID ของ Spare_TypeC Google Sheet (จากลิงก์ที่ส่งมา)
const SPARE_TYPE_C_SHEET_ID = "1dZXwYcPWc0OgWriXnA5ZxbAui8GmJDLRddrKTlZPuic";

function syncMaterialCodeFromSpareTypeC() {
  try {
    // ── เปิด Sheet ทั้งสองฝั่ง ──────────────────────────────────────────────
    const stockSS = SpreadsheetApp.openById(SHEET_ID);
    const stockSheet = stockSS.getSheetByName("STOCK");

    const spareSS = SpreadsheetApp.openById(SPARE_TYPE_C_SHEET_ID);
    const spareSheet = spareSS.getSheetByName("Spare Type C");

    if (!stockSheet) throw new Error("ไม่พบ STOCK sheet");
    if (!spareSheet) throw new Error("ไม่พบ Spare Type C sheet");

    // ── Config ───────────────────────────────────────────────────────────────
    const SPARE_START_ROW = 4; // แถวแรกที่มีข้อมูลใน Spare_TypeC (0-based index = แถว 5 ใน Sheet)
    const SPARE_MAT_COL = 1; // Column B (0-based) = Material Code ใน Spare_TypeC
    const STOCK_MAT_COL = 2; // Column B (1-based ใช้กับ getRange) = Material Code ใน STOCK

    // ── ดึงข้อมูลจาก Spare_TypeC ─────────────────────────────────────────────
    const spareData = spareSheet.getDataRange().getValues();

    // กรองเฉพาะแถวที่มี Material Code จริงๆ พร้อมจำ targetRow ที่จะเขียนใน STOCK
    const updates = []; // [ { targetRow, materialCode }, ... ]

    for (let i = SPARE_START_ROW; i < spareData.length; i++) {
      const rawCode = spareData[i][SPARE_MAT_COL];
      if (!rawCode || rawCode.toString().trim() === "") continue;

      const materialCode =
        typeof rawCode === "number"
          ? rawCode.toFixed(0) // ตัดทศนิยมออก เช่น 2000039121.0 → "2000039121"
          : rawCode.toString().trim();

      // STOCK เริ่มวางข้อมูลที่แถว 2 (แถว 1 คือ Header)
      const targetRow = i - SPARE_START_ROW + 2;

      updates.push({ targetRow, materialCode });
    }

    if (updates.length === 0) {
      Logger.log("⚠️ ไม่พบข้อมูลใน Spare_TypeC ที่จะซิงค์");
      return;
    }

    // ── ป้องกันการเขียนเกินแถวที่มีอยู่ใน STOCK ─────────────────────────────
    const stockLastRow = stockSheet.getLastRow();
    const safeUpdates = updates.filter((u) => u.targetRow <= stockLastRow);
    const skipped = updates.length - safeUpdates.length;

    if (skipped > 0) {
      Logger.log(`⚠️ ข้าม ${skipped} รายการ (เกินจำนวนแถวใน STOCK)`);
    }

    // ── เขียนลง STOCK ทีเดียว (batch) ────────────────────────────────────────
    // เรียงตาม targetRow เพื่อเขียนเป็น batch range เดียว (ประหยัด API call)
    safeUpdates.sort((a, b) => a.targetRow - b.targetRow);

    const firstRow = safeUpdates[0].targetRow;
    const lastRow = safeUpdates[safeUpdates.length - 1].targetRow;
    const numRows = lastRow - firstRow + 1;

    // สร้าง array ขนาดเต็ม แล้วใส่ค่าในช่องที่ต้องการ (ช่องที่ข้ามไว้เป็นสตริงว่าง)
    const batchValues = Array.from({ length: numRows }, () => [""]);
    safeUpdates.forEach((u) => {
      batchValues[u.targetRow - firstRow][0] = u.materialCode;
    });

    // เขียนทีเดียว — Column B ของ STOCK (STOCK_MAT_COL = 2)
    stockSheet
      .getRange(firstRow, STOCK_MAT_COL, numRows, 1)
      .setValues(batchValues);

    SpreadsheetApp.flush();
    Logger.log(
      `✅ ซิงค์สำเร็จ ${safeUpdates.length} รายการ (แถว ${firstRow} ถึง ${lastRow})`,
    );
  } catch (err) {
    Logger.log("❌ syncMaterialCodeFromSpareTypeC Error: " + err.toString());
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
    const stockData = stockSS
      .getSheetByName("STOCK")
      .getDataRange()
      .getValues();

    // หาคอลัมน์ Material Code
    let materialCodeColIndex = -1;
    for (let col = 0; col < spareData[0].length; col++) {
      const header = spareData[0][col]
        ? spareData[0][col].toString().trim().toLowerCase()
        : "";
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
      const spareMatCode = spareData[i][materialCodeColIndex]
        ? spareData[i][materialCodeColIndex].toString().trim()
        : "";

      const stockMatCode = stockData[i]
        ? stockData[i][2]
          ? stockData[i][2].toString().trim()
          : ""
        : "";

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
      const oldCode = `Q${String(i).padStart(3, "0")}`;
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
    Logger.log(
      `✅ เชื่อมต่อ STOCK sheet สำเร็จ (${stockSheet.getLastRow()} rows)`,
    );

    // Test Spare_TypeC sheet
    const spareSS = SpreadsheetApp.openById(SPARE_TYPE_C_SHEET_ID);
    const spareTypeC = spareSS.getSheetByName("Spare Type C");
    Logger.log(
      `✅ เชื่อมต่อ Spare Type C สำเร็จ (${spareTypeC.getLastRow()} rows)`,
    );

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
        [{ text: "/stock" }, { text: "/help" }],
      ],
      resize_keyboard: true,
      one_time_keyboard: false,
    },
  };
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
  };
  UrlFetchApp.fetch(
    `https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`,
    options,
  );
}

// เก็บ Logs
function writeSimpleLog(action, details) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let logSheet = ss.getSheetByName("LOGS");

    // ถ้ายังไม่มี Sheet ชื่อ LOGS ให้สร้างขึ้นมาใหม่พร้อม Header
    if (!logSheet) {
      logSheet = ss.insertSheet("LOGS");
      logSheet.appendRow(["Timestamp", "User Email", "Action", "Details"]);
      logSheet
        .getRange("A1:D1")
        .setBackground("#252b38")
        .setFontColor("#ffffff")
        .setFontWeight("bold");
    }

    const userEmail = Session.getActiveUser().getEmail() || "Guest/Anonymous";
    const timestamp = new Date();

    logSheet.appendRow([timestamp, userEmail, action, details]);
  } catch (err) {
    Logger.log("Log Error: " + err.toString());
  }
}

function getTopWithdrawItems(rows) {
  const counter = {};
  rows.forEach((r) => (counter[r[3]] = (counter[r[3]] || 0) + Number(r[4])));
  return Object.entries(counter)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);
}

function clearAiMemory() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const memSheet = ss.getSheetByName("AI_Memory");
  if (!memSheet) {
    Logger.log("ไม่พบ AI_Memory sheet");
    return;
  }

  const data = memSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    memSheet.getRange(i + 1, 2).setValue("[]"); // ล้าง history ทุก user
  }
  SpreadsheetApp.flush();
  Logger.log("✅ ล้าง AI_Memory เรียบร้อย");
}

// เช็ค Key
function checkCurrentKey() {
  Logger.log("Key ที่ใช้อยู่: " + GEMINI_API_KEY.slice(0, 15) + "...");
  Logger.log("Model ที่ใช้อยู่: " + "gemini-2.5-flash-lite");
}

function callGroq(msg, chatId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const stockSheet = ss.getSheetByName("STOCK");

    const dataRange = stockSheet.getRange("B2:F100").getValues();
    const stockSummary = dataRange
      .filter((r) => r[0] && !String(r[0]).includes("CellImage") && r[2])
      .slice(0, 40)
      .map((r) => {
        const code =
          typeof r[0] === "number" ? r[0].toFixed(0) : String(r[0]).trim();
        const qty = Number(r[4]) || 0;
        return `${code} | ${String(r[2]).trim()} | คงเหลือ: ${qty}`;
      })
      .join("\n");

    // ✅ เพิ่มตรงนี้
    const userMsg = (msg || "").toString().trim();
    if (!userMsg) return "⚠️ กรุณาพิมพ์ข้อความครับ";

    const systemPrompt = `คุณคือบอทเช็คสต็อกของ SYS ตอบภาษาไทยสั้นๆ ไม่เกิน 3 บรรทัด\n\nสต็อกปัจจุบัน:\n${stockSummary}`;

    const payload = {
      model: "llama-3.3-70b-versatile",
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: userMsg }, // ✅ เปลี่ยน msg → userMsg
      ],
      max_tokens: 200,
      temperature: 0.7,
    };

    const response = UrlFetchApp.fetch(
      "https://api.groq.com/openai/v1/chat/completions",
      {
        method: "post",
        contentType: "application/json",
        headers: { Authorization: "Bearer " + GROQ_API_KEY },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      },
    );

    const resCode = response.getResponseCode();
    const resText = response.getContentText();

    if (resCode !== 200) {
      Logger.log("❌ Groq Error " + resCode + ": " + resText);
      return "🤖 ระบบขัดข้อง (" + resCode + ") กรุณาลองใหม่ครับ";
    }

    const json = JSON.parse(resText);
    return json?.choices?.[0]?.message?.content || "ไม่สามารถดึงคำตอบได้";
  } catch (e) {
    Logger.log("❌ callGroq Exception: " + e.toString());
    return "🤖 ระบบขัดข้อง: " + e.toString();
  }
}
