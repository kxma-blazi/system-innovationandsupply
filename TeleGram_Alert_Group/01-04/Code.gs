// ============================================================
// 2. WEB DASHBOARD (หน้าเว็บแสดงผล)
// ============================================================
function doGet(e) {
  const page = (e && e.parameter.page) || "stock";
  const search = (e && e.parameter.search) || "";

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const stockSheet = ss.getSheetByName("STOCK");
  const logSheet = ss.getSheetByName("Logs") || ss.insertSheet("Logs");

  const stockData = stockSheet ? stockSheet.getDataRange().getValues() : [];
  const logData = logSheet ? logSheet.getDataRange().getValues() : [];

  // 1. กรองข้อมูลสต็อก
  // แก้ r[0] เป็น r[2] เพื่อให้กรองเฉพาะแถวที่มีรหัสสินค้าจริงๆ
  let rows = stockData
    .slice(1)
    .filter((r) => r[2] && r[2].toString().trim() !== "");

  if (search) {
    const kw = search.toLowerCase();
    rows = rows.filter(
      (r) =>
        (r[2] && r[2].toString().toLowerCase().includes(kw)) || // ค้นหาจาก Material Code (Col C) **เพิ่มอันนี้**
        (r[3] && r[3].toString().toLowerCase().includes(kw)) || // ค้นหาจากชื่อสินค้า (Col D)
        (r[0] && r[0].toString().toLowerCase().includes(kw)), // ค้นหาจากลำดับ (Col A)
    );
  }

  // ใน doGet ส่วนสร้าง stockRows
  const stockRows = rows
    .map((r) => {
      const qty = Number(r[5]) || 0;
      const matCode = r[2] || "-";

      const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${encodeURIComponent(matCode)}`;
      const qrImg = `<img src="${qrUrl}" width="45" height="45" style="border-radius:4px; cursor:pointer; border:1px solid #ccc;" onclick="window.open('${qrUrl}')">`;

      const badge =
        qty <= 0
          ? `<span class="badge red">หมด</span>`
          : qty <= LOW_STOCK_LIMIT
            ? `<span class="badge yellow">ใกล้หมด</span>`
            : `<span class="badge green">ปกติ</span>`;

      return `<tr>
      <td><code>${matCode}</code></td>
      <td style="text-align:center;">${qrImg}</td>
      <td>${r[3] || "-"}</td>
      <td class="qty">${qty}</td>
      <td>${r[6] || 0}</td>
      <td>${r[7] || 0}</td>
      <td>${r[8] || 0}</td>
      <td>${r[9] || 0}</td>
      <td>${r[10] || 0}</td>
      <td>${badge}</td>
    </tr>`;
    })
    .join("");

  const allStockRows = stockData.slice(1).filter((r) => r[0]);
  const totalItems = allStockRows.length;
  const lowCount = allStockRows.filter((r) => {
    const q = Number(r[5]);
    return q > 0 && q <= LOW_STOCK_LIMIT;
  }).length;
  const outCount = allStockRows.filter((r) => Number(r[5]) <= 0).length;

  const logRows = logData
    .slice(1)
    .reverse()
    .slice(0, 50)
    .map((r) => {
      const dt = r[0]
        ? Utilities.formatDate(new Date(r[0]), "GMT+7", "dd/MM/yy HH:mm")
        : "-";
      const ac = r[2] === "Withdraw" ? "withdraw" : "restock";
      return `<tr><td>${dt}</td><td>${r[1] || "-"}</td><td><span class="action ${ac}">${r[2] || "-"}</span></td><td><code>${r[3] || "-"}</code></td><td>${r[4] || 0}</td><td>${r[5] || "-"}</td><td class="qty">${r[6] || 0}</td></tr>`;
    })
    .join("");

  // 5. แก้ไขการสร้าง deliveryRows ให้ดึงข้อมูลมาโชว์
  let deliveryRows = "";
  if (page === "delivery") {
    const deliveryData = getDeliveryRecords(100);
    deliveryRows =
      (deliveryData.records || [])
        .map((r) => {
          return `<tr>
        <td><code>${r.code}</code></td>
        <td><code>${r.serial}</code></td>
        <td>${r.dateDelivered}</td>
        <td><span class="action withdraw">${r.note}</span></td>
        <td>${r.deliveredBy}</td>
      </tr>`;
        })
        .join("") ||
      `<tr><td colspan="5" class="empty">ยังไม่มีการนำจ่ายซีเรียล</td></tr>`;
  }

  const tmpl = HtmlService.createTemplateFromFile("Index");
  tmpl.page = page;
  tmpl.search = search;
  tmpl.stockRows = stockRows;
  tmpl.logRows = logRows;
  tmpl.deliveryRows = deliveryRows; // ส่งตัวแปรนี้ไปที่หน้าเว็บ
  tmpl.totalItems = totalItems;
  tmpl.lowCount = lowCount;
  tmpl.outCount = outCount;
  tmpl.resultCount = rows.length;

  return tmpl
    .evaluate()
    .setTitle("SYS Stock Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
      if (
        data[i][2].toString().toLowerCase() === code.toString().toLowerCase()
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

  for (let i = 1; i < sData.length; i++) {
    // 1. เช็คว่ามีคนเบิกไปหรือยัง (Col C)
    if (sData[i][2].toString().trim() === serial) {
      takerInfo = {
        matCode: sData[i][0],
        date: sData[i][3],
        name: sData[i][5],
      };
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
    if (
      stockData[j][2].toString().toLowerCase() ===
      matCode.toString().toLowerCase()
    ) {
      newTotal = (Number(stockData[j][5]) || 0) - 1; // คอลัมน์ F (Total)
      const newDeptQty = (Number(stockData[j][colIndex]) || 0) + 1; // คอลัมน์แผนก

      stockSheet.getRange(j + 1, 6).setValue(newTotal);
      stockSheet.getRange(j + 1, colIndex + 1).setValue(newDeptQty);
      break;
    }
  }

  writeLog(
    user,
    "Withdraw (SN)",
    matCode,
    1,
    `SN: ${serial} (${deptName})`,
    newTotal,
  );

  let resMsg = `✅ *เบิกสำเร็จ (Serial)*\n📦 รหัส: ${matCode}\n🏷 S/N: \`${serial}\`\n📊 เหลือรวม: ${newTotal}`;
  if (newTotal <= LOW_STOCK_LIMIT)
    resMsg += `\n\n⚠️ *ALERT: สินค้าใกล้หมดแล้ว!*`;
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

// แจ้งเตือนเมื่อแก้ชีตโดยตรง
function onEdit(e) {
  try {
    const range = e.range;
    const sheetName = range.getSheet().getName();

    const skipSheets = ["Logs", "Summary", "Config", "AI_Memory"];
    if (skipSheets.includes(sheetName)) return;

    const oldVal = e.oldValue || "ว่าง";
    const newVal = e.value || "ว่าง";
    const user = Session.getActiveUser().getEmail() || "Unknown";

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

// ============================================================
// AI Gemini (Version: Real-time Stock Sync)
// ============================================================
function callGemini(msg, chatId) {
  const props = PropertiesService.getScriptProperties();
  const API_KEY = GEMINI_API_KEY;
  const MODEL = props.getProperty("GEMINI_MODEL") || "gemini-1.5-flash";
  const SYSTEM_BASE_PROMPT =
    props.getProperty("SYSTEM_PROMPT") || "คุณคือ SYS Bot ผู้ช่วยจัดการสต็อก";

  if (!API_KEY) return "❌ กรุณาตั้งค่า API_KEY ใน Script Properties";
  if (!msg || String(msg).trim() === "") return "🤖 รอรับคำสั่งครับเฮีย";

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);

    // --- ส่วนที่เพิ่มใหม่: ดึงข้อมูลจากชีต STOCK มาทำเป็นบริบทให้ AI ---
    const stockSheet = ss.getSheetByName("STOCK");
    const stockData = stockSheet.getDataRange().getValues();
    let stockSummary = "ข้อมูลสต็อกปัจจุบัน (อัปเดตล่าสุด):\n";

    // วนลูปดึงเฉพาะ คอลัมน์ C (รหัส), D (ชื่อ), F (คงเหลือรวม)
    for (let i = 1; i < stockData.length; i++) {
      const code = stockData[i][2]; // Index 2 = Col C
      const name = stockData[i][3]; // Index 3 = Col D
      const qty = stockData[i][5]; // Index 5 = Col F
      if (code) {
        stockSummary += `- ${code}: ${name} (คงเหลือ: ${qty})\n`;
      }
    }

    // รวมร่าง Prompt: คำสั่งพื้นฐาน + ข้อมูลสต็อกสดๆ
    const finalPrompt = `${SYSTEM_BASE_PROMPT}\n\n${stockSummary}\nหมายเหตุ: ถ้าผู้ใช้ถามเช็คของ ให้ตอบตามข้อมูลข้างบนนี้ได้เลย`;
    // -----------------------------------------------------------

    let memSheet =
      ss.getSheetByName("AI_Memory") || ss.insertSheet("AI_Memory");
    const memData = memSheet.getDataRange().getValues();
    let history = [];
    let userRow = -1;

    for (let i = 1; i < memData.length; i++) {
      if (
        memData[i][0] &&
        memData[i][0].toString() === (chatId || "").toString()
      ) {
        userRow = i + 1;
        try {
          const rawHistory = JSON.parse(memData[i][1] || "[]");
          history = Array.isArray(rawHistory)
            ? rawHistory.filter(
                (h) => h.role && h.text && String(h.text).trim() !== "",
              )
            : [];
        } catch (e) {
          history = [];
        }
        break;
      }
    }

    const contents = history.map((h) => ({
      role: h.role === "model" ? "model" : "user",
      parts: [{ text: String(h.text).trim() }],
    }));

    contents.push({ role: "user", parts: [{ text: String(msg).trim() }] });

    const payload = {
      contents: contents,
      systemInstruction: { parts: [{ text: finalPrompt }] },
    };

    const url = `https://generativelanguage.googleapis.com/v1/models/${MODEL}:generateContent?key=${API_KEY}`;
    const res = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    const resCode = res.getResponseCode();
    const resBody = res.getContentText();

    if (resCode !== 200) {
      if (resCode === 400 && userRow !== -1)
        memSheet.getRange(userRow, 2).setValue("[]");
      return `🤖 AI ไม่พร้อมใช้งาน (${resCode})`;
    }

    const json = JSON.parse(resBody);
    const aiResponse = json?.candidates?.[0]?.content?.parts?.[0]?.text;
    if (!aiResponse) return "🤖 ขอโทษครับเฮีย ผมงงนิดหน่อย";

    // บันทึกความจำ
    const newHistory = history;
    newHistory.push({ role: "user", text: String(msg).trim() });
    newHistory.push({ role: "model", text: aiResponse.trim() });
    const finalHistory = JSON.stringify(newHistory.slice(-10));

    if (userRow !== -1) {
      memSheet.getRange(userRow, 2).setValue(finalHistory);
      memSheet.getRange(userRow, 3).setValue(new Date());
    } else {
      memSheet.appendRow([String(chatId), finalHistory, new Date()]);
    }

    return aiResponse.trim();
  } catch (e) {
    Logger.log("🔥 Critical Error: " + e.toString());
    return "🤖 ระบบ AI ขัดข้องชั่วคราว";
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
    const MATERIAL_COL_INDEX = 2; // Column C

    let syncCount = 0;

    for (let i = START_ROW_SPARE; i < spareData.length; i++) {
      const materialCode = spareData[i][MATERIAL_COL_INDEX];

      if (materialCode && materialCode.toString().trim() !== "") {
        // คำนวณแถวที่จะไปวางใน STOCK (เริ่มวางที่แถว 2 เป็นต้นไป)
        const targetRow = i - START_ROW_SPARE + 2;

        // เขียนลง Column C ของหน้า STOCK
        stockSheet
          .getRange(targetRow, 3)
          .setValue(materialCode.toString().trim());
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
