// ============================================================
// 1. CONFIGURATION
// ============================================================
const TELEGRAM_TOKEN = "Eight327163778:AAFM4aKpxT29WTB4z_StzKEEcRGrSBDS2_s";
const CHAT_ID = "-One003731290917"; 
const LOW_STOCK_LIMIT = 1;
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbw5KR2d8ukXl0DVUAj8Lo11JQ8F00vL-ax7pkqRbDeSv1NTLNIRy094narUUyz_TbXH/exec";

// ============================================================
// 2. CORE SYSTEM (Support ทั้งพิมพ์ และ กดปุ่ม)
// ============================================================
function doPost(e) {
  try {

    if (!e || !e.postData) return;

    const data = JSON.parse(e.postData.contents);

    // =============================
    // 1️⃣ CALLBACK BUTTON
    // =============================
    if (data.callback_query) {

      const callbackData = data.callback_query.data;
      const callbackId = data.callback_query.id;

      if (!data.callback_query.message) return;

      const chatId = data.callback_query.message.chat.id;

      let responseText = "";

      if (callbackData === "/allstock") {
        sendStockPage(chatId, 1);
        answerCallback(callbackId);
        return;
      }

      else if (callbackData.startsWith("/allstock_")) {

        const page = Number(callbackData.split("_")[1]) || 1;

        sendStockPage(chatId, page);
        answerCallback(callbackId);
        return;
      }

      else if (callbackData === "/lowstock") {
        responseText = lowStock();
      }

      else if (callbackData === "/report") {
        responseText = report();
      }

      else if (callbackData === "/history") {
        responseText = history();
      }

      else if (callbackData === "/help") {
        responseText = helpMenu();
      }

      if (responseText) {
        sendTelegram(chatId, responseText);
      }

      answerCallback(callbackId);
      return;
    }

    // =============================
    // 2️⃣ TEXT COMMAND
    // =============================
    if (!data.message || !data.message.text) return;

    const chatId = data.message.chat.id;
    const text = data.message.text;

    const user =
      data.message.from.first_name ||
      data.message.from.username ||
      "Unknown";

    let cmd = text.split(" ")[0].toLowerCase();

    if (cmd.includes("@")) {
      cmd = cmd.split("@")[0];
    }

    const args = text.split(" ");

    switch (cmd) {

      case "/exportpdf":
        sendTelegram(chatId,"📄 กำลังสร้าง PDF...");
        exportStockPDF(chatId);
        break;

      case "/start":
      case "/menu":
        sendControlPanel(chatId);
        break;

      case "/help":
      case "/manual":
        sendTelegram(chatId, helpMenu());
        sendControlPanel(chatId);
        break;

      case "/history":
        sendTelegram(chatId, history());
        break;

      case "/stock":
        sendTelegram(chatId, getStock(args[1]));
        break;

      case "/search":
        sendTelegram(chatId, searchStock(args.slice(1).join(" ")));
        break;

      case "/allstock":

        let page = args[1] ? Number(args[1]) : 1;
        if (isNaN(page)) page = 1;

        sendStockPage(chatId, page);
        break;

      case "/lowstock":
        sendTelegram(chatId, lowStock());
        break;

      case "/restock":
        sendTelegram(chatId, restock(args[1], Number(args[2]), user));
        break;

      case "/topwithdraw":
        sendTelegram(chatId, topWithdraw());
        break;

      case "/dp-cbr":
        sendTelegram(chatId, withdraw(args[1], Number(args[2]), 5, "CBR", user));
        break;

      case "/dp-ccs":
        sendTelegram(chatId, withdraw(args[1], Number(args[2]), 6, "CCS", user));
        break;

      case "/dp-sko":
        sendTelegram(chatId, withdraw(args[1], Number(args[2]), 7, "SKO", user));
        break;

      case "/dp-ryg":
        sendTelegram(chatId, withdraw(args[1], Number(args[2]), 8, "RYG", user));
        break;

      case "/dp-trt":
        sendTelegram(chatId, withdraw(args[1], Number(args[2]), 9, "TRT", user));
        break;

      default:
        if (cmd.startsWith("/")) {
          sendTelegram(chatId,"❌ ไม่รู้จักคำสั่งนี้ครับ\nลองพิมพ์ /menu");
        }

    }

  } catch (err) {

    Logger.log(err);

  }
}

// ============================================================
// 3. UI FUNCTIONS (แผงควบคุม & การโต้ตอบ)
// ============================================================

function sendControlPanel(chatId) {
  const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`;
  
  // 1. สร้างโครงสร้างปุ่ม (Inline Keyboard)
  const keyboard = {
    "inline_keyboard": [
      [
        { "text": "📦 ดูสต็อกทั้งหมด", "callback_data": "/allstock" }, 
        { "text": "⚠️ ของใกล้หมด", "callback_data": "/lowstock" }
      ],
      [
        { "text": "📊 สรุปรายงาน", "callback_data": "/report" }, 
        { "text": "📈 ดู Dashboard", "url": "https://lookerstudio.google.com/reporting/468488a2-94a5-4db9-9945-8b6f1e2008dd" }
      ],
      [
        { "text": "📜 ประวัติล่าสุด", "callback_data": "/history" },
        { "text": "📄 ออกไฟล์ PDF", "callback_data": "/exportpdf" }
      ],
      [
        { "text": "📤 วิธีเบิก/เติม", "callback_data": "/help" },
        { "text": "📚 คู่มือใช้งาน", "callback_data": "/help" }
      ]
    ]
  };

  // 2. เตรียมข้อมูลที่จะส่ง
  const payload = {
    "chat_id": chatId ? chatId.toString() : CHAT_ID,
    "text": "🎮 <b>SYSTEM INNOVATION AND SUPPLY</b>\n\nยินดีต้อนรับ! เลือกรายการที่ต้องการจัดการด้านล่างครับ:",
    "parse_mode": "HTML",
    "reply_markup": keyboard 
  };

  // 3. ตั้งค่าการส่งข้อมูล
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  // 4. ส่งข้อมูลไปยัง Telegram
  try {
    UrlFetchApp.fetch(url, options);
  } catch (e) {
    Logger.log("Error sending control panel: " + e.message);
  }
}

function answerCallback(callbackId) {
  const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/answerCallbackQuery`;
  UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify({ callback_query_id: callbackId }) });
}

// ============================================================
// 4. STOCK LOGIC (คำนวณ & บันทึกผล)
// ============================================================

function getSheet() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("STOCK"); }

function restock(code, qty, user) {
  if (!code || isNaN(qty) || qty <= 0) return "⚠️ รูปแบบผิด: /restock [CODE] [QTY]";
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === code.toLowerCase()) {
      let newQty = Number(data[i][4]) + qty;
      sheet.getRange(i + 1, 5).setValue(newQty);
      writeLog(user, "Restock", data[i][0], qty, "Center", newQty);
      notifyActivity("เติมสินค้า | Restock", data[i][0], "IN", qty, newQty, "", user);
      return `✅ <b>Restock Success</b>\n📦 ${data[i][2]}\n👤 โดย: ${user}\n📊 Stock ใหม่: ${newQty}`;
    }
  }
  return "❌ ไม่พบรหัสสินค้า";
}

function withdraw(code, qty, colIndex, deptName, user) {
  if (!code || isNaN(qty) || qty <= 0) return `⚠️ รูปแบบผิด: /dp-xxx [CODE] [QTY]`;
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === code.toLowerCase()) {
      let currentStock = Number(data[i][4]);
      if (currentStock < qty) return `❌ <b>สต็อกไม่พอ!</b>\nคงเหลือ: ${currentStock}`;
      let newStock = currentStock - qty;
      let newDeptTotal = (Number(data[i][colIndex]) || 0) + qty;
      sheet.getRange(i + 1, 5).setValue(newStock); 
      sheet.getRange(i + 1, colIndex + 1).setValue(newDeptTotal);
      writeLog(user, "Withdraw", data[i][0], qty, deptName, newStock);
      notifyActivity(`เบิกสินค้า (${deptName})`, data[i][0], "OUT", qty, newStock, deptName, user);
      return `✅ <b>Success (${deptName})</b>\n📦 ${data[i][2]}\n👤 โดย: ${user}\n📉 เหลือ: ${newStock}`;
    }
  }
  return "❌ ไม่พบรหัสสินค้า";
}

function writeLog(user, action, code, qty, dept, remain) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("Logs");
  if (!logSheet) return;
  logSheet.appendRow([new Date(), user, action, code, qty, dept, remain]);
}

// ============================================================
// 5. QUERY & REPORT FUNCTIONS
// ============================================================

function getStock(code) {
  if (!code) return "⚠️ ระบุรหัส: /stock Q01";
  const data = getSheet().getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === code.toLowerCase()) {
      return `📦 <b>ข้อมูลสินค้า</b>\n🔖 Code: ${data[i][0]}\n📝 Desc: ${data[i][2]}\n📊 สต็อก: <b>${data[i][4]}</b>`;
    }
  }
  return "❌ ไม่พบข้อมูล";
}

function searchStock(keyword) {
  if (!keyword) return "⚠️ ใส่คำค้นหา: /search [ชื่อ]";
  const data = getSheet().getDataRange().getValues();
  let results = "";
  for (let i = 1; i < data.length; i++) {
    if (data[i][2].toString().toLowerCase().includes(keyword.toLowerCase())) {
      results += `• [${data[i][0]}] ${data[i][2]} (${data[i][4]})\n`;
    }
  }
  return results ? `🔎 <b>ผลการค้นหา: "${keyword}"</b>\n\n${results}` : "❌ ไม่พบข้อมูล";
}

function sendStockPage(chatId, page = 1) {

  if (!chatId) {
    Logger.log("❌ chatId is empty");
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("STOCK");
  const data = sheet.getDataRange().getValues();

  const perPage = 20;
  const totalPages = Math.ceil((data.length - 1) / perPage);

  if (page < 1) page = 1;
  if (page > totalPages) page = totalPages;

  const start = (page - 1) * perPage + 1;
  const end = Math.min(start + perPage, data.length);

  let msg = `📦 <b>STOCK LIST (Page ${page}/${totalPages})</b>\n`;
  msg += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n";
  msg += "<pre>";
  msg += "ID   QTY  DESCRIPTION\n";
  msg += "--------------------------------\n";

  for (let i = start; i < end; i++) {

    let id = (data[i][0] || "").toString().padEnd(4," ");
    let qty = (data[i][4] || 0).toString().padEnd(4," ");
    let desc = (data[i][2] || "").substring(0,25);

    msg += `${id} ${qty} ${desc}\n`;

  }

  msg += "</pre>";

  let keyboard = [];
  let row = [];

  if (page > 1) {
    row.push({
      text: "⬅️ ก่อนหน้า",
      callback_data: `/allstock_${page-1}`
    });
  }

  if (page < totalPages) {
    row.push({
      text: "➡️ ถัดไป",
      callback_data: `/allstock_${page+1}`
    });
  }

  if (row.length > 0) {
    keyboard.push(row);
  }

  // ปุ่มกลับเมนู
  keyboard.push([
    { text: "🏠 กลับเมนู", callback_data: "/menu" }
  ]);

  const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`;

  const payload = {
    chat_id: chatId.toString(),
    text: msg,
    parse_mode: "HTML",
    reply_markup: {
      inline_keyboard: keyboard
    }
  };

  UrlFetchApp.fetch(url,{
    method:"post",
    contentType:"application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions:true
  });

}

// ============================================================
// TOP WITHDRAW REPORT (สินค้าที่ถูกเบิกมากที่สุด)
// ============================================================

function topWithdraw() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("LOGS");

  if (!sheet) return "❌ ไม่พบชีต LOGS";

  const data = sheet.getDataRange().getValues();

  let withdrawMap = {};

  // 🔍 วนลูปอ่าน log ทั้งหมด
  for (let i = 1; i < data.length; i++) {

    const action = data[i][2]; // Restock / Withdraw
    const code = data[i][3];   // รหัสสินค้า
    const qty = Number(data[i][4]) || 0;

    // นับเฉพาะ Withdraw
    if (action === "Withdraw") {

      if (!withdrawMap[code]) {
        withdrawMap[code] = 0;
      }

      withdrawMap[code] += qty;
    }
  }

  // 🔄 แปลง object → array เพื่อ sort
  const sorted = Object.entries(withdrawMap)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10); // เอา Top 10

  if (sorted.length === 0) {
    return "📊 ยังไม่มีข้อมูลการเบิกสินค้า";
  }

  let msg = "🏆 <b>Top 10 สินค้าที่ถูกเบิกมากที่สุด</b>\n";
  msg += "━━━━━━━━━━━━━━━━━━\n";

  sorted.forEach((item, index) => {

    const code = item[0];
    const qty = item[1];

    msg += `${index + 1}. 📦 <code>${code}</code>\n`;
    msg += `   เบิกทั้งหมด: <b>${qty}</b> ชิ้น\n\n`;

  });

  return msg;
}

// ============================================================
// 6. TELEGRAM UTILITIES
// ============================================================

function sendTelegram(chatId, msg) {
  const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`;
  UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json; charset=utf-8",
    payload: JSON.stringify({ chat_id: chatId, text: msg, parse_mode: "HTML" }),
    muteHttpExceptions: true
  });
}

function notifyActivity(title, code, type, qty, remain, dept = "", user = "") {
  // 1. แยกสีและ Emoji ตามประเภทรายการให้ชัดเจน
  const statusEmoji = type === "IN" ? "🟢 [ เติมเข้า ]" : "🟠 [ เบิกออก ]";
  const mathSign = type === "IN" ? "➕" : "➖";

  // 2. ออกแบบข้อความใหม่
  let msg = `📌 <b>${title}</b>\n`;
  msg += `━━━━━━━━━━━━━━━━\n`;
  msg += `🆔 <b>รหัสสินค้า:</b> <code>${code}</code>\n`; // 
  msg += `📝 <b>สถานะ:</b> ${statusEmoji}\n`;
  msg += `🔢 <b>จำนวน:</b> ${mathSign} <b>${qty.toLocaleString()}</b> หน่วย\n`; // 
  
  if (dept) {
    msg += `🏢 <b>แผนก:</b> ${dept}\n`; // 
  }

  msg += `👤 <b>ผู้ทำรายการ:</b> ${user}\n`; // 
  msg += `━━━━━━━━━━━━━━━━\n`;
  msg += `📊 <b>สต็อกคงเหลือปัจจุบัน:</b>\n`;
  msg += `➡️ <b><u>${remain.toLocaleString()}</u></b> หน่วย`; // 

  sendTelegram(CHAT_ID, msg); // [cite: 85]
}

function helpMenu(){
return `
🧭<b>SYSTEM INNOVATION & SUPPLY</b>
━━━━━━━━━━━━━━━━━━━
📚 <b>คู่มือการใช้งาน (Smart Guide)</b>
━━━━━━━━━━━━━━━━━━━

🧭 <b>เริ่มต้นใช้งาน</b>
┌─────────────────
👉 <code>/menu</code>
เปิดเมนูควบคุม (กดปุ่มใช้ง่ายสุด)
└─────────────────

━━━━━━━━━━━━━━━━━━━
🔎 <b>ตรวจสอบสินค้า</b>
┌─────────────────
📦 <code>/stock CODE</code>
ดูสต็อกสินค้า

🔍 <code>/search NAME</code>
ค้นหาสินค้า

📋 <code>/allstock</code>
ดูสินค้าทั้งหมด

⚠️ <code>/lowstock</code>
ดูสินค้าใกล้หมด
└─────────────────

━━━━━━━━━━━━━━━━━━━
📥 <b>เติมสินค้า</b>
┌─────────────────
➕ <code>/restock CODE QTY</code>

📌 ตัวอย่าง:
<code>/restock Q001 50</code>
└─────────────────

━━━━━━━━━━━━━━━━━━━
📤 <b>เบิกสินค้า</b>
┌─────────────────
รูปแบบ:
<code>/dp-แผนก CODE QTY</code>

📌 ตัวอย่าง:
<code>/dp-cbr Q001 5</code>

🏢 แผนก:
• <code>cbr</code>
• <code>ccs</code>
• <code>sko</code>
• <code>ryg</code>
• <code>trt</code>
└─────────────────

━━━━━━━━━━━━━━━━━━━
📊 <b>รายงาน & วิเคราะห์</b>
┌─────────────────
📊 <code>/report</code>
ภาพรวมคลัง

📜 <code>/history</code>
ประวัติล่าสุด

🏆 <code>/topwithdraw</code>
สินค้าที่ถูกเบิกมากสุด

📄 <code>/exportpdf</code>
ดาวน์โหลด PDF
└─────────────────

`;
}

function history(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("LOGS");
  if(!sheet) return "❌ ไม่พบชีต LOGS";
  const data = sheet.getDataRange().getValues();
  let msg = "📜 <b>Stock History (10 รายการล่าสุด)</b>\n\n";
  let start = Math.max(1, data.length - 10);
  for(let i=start; i<data.length; i++){
    let date = Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "dd/MM HH:mm");
    msg += `🕒 ${date} | 👤 ${data[i][1]}\n📦 ${data[i][3]} (${data[i][2]}) ${data[i][4]} หน่วย\n📊 เหลือ ${data[i][6]}\n\n`;
  }
  return msg;
}

// ============================================================
// 7. AUTOMATION & TRIGGERS
// ============================================================

function sendDailySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName("STOCK");
  const logSheet = ss.getSheetByName("LOGS");
  if (!stockSheet || !logSheet) return;

  const stockData = stockSheet.getDataRange().getValues();
  const logData = logSheet.getDataRange().getValues();
  const today = new Date().toLocaleDateString('en-CA');
  
  let lowStockList = "";
  let lowCount = 0;
  let dailyIn = 0;
  let dailyOut = 0;

  for (let i = 1; i < stockData.length; i++) {
    if (Number(stockData[i][4]) <= LOW_STOCK_LIMIT) {
      lowStockList += `🚨 <code>${stockData[i][0]}</code> | เหลือ: <b>${stockData[i][4]}</b>\n`;
      lowCount++;
    }
  }

  for (let i = 1; i < logData.length; i++) {
    let logDate = new Date(logData[i][0]).toLocaleDateString('en-CA');
    if (logDate === today) {
      if (logData[i][2] === "Restock") dailyIn += Number(logData[i][4]);
      if (logData[i][2] === "Withdraw") dailyOut += Number(logData[i][4]);
    }
  }

  let msg = `📊 <b>สรุปคลังสินค้าประจำวัน</b>\n📅 วันที่: ${new Date().toLocaleDateString('th-TH')}\n━━━━━━━━━━━━━━━\n📥 เติมวันนี้: ${dailyIn}\n📤 เบิกวันนี้: ${dailyOut}\n⚠️ ใกล้หมด: ${lowCount} รายการ\n━━━━━━━━━━━━━━━\n${lowCount > 0 ? lowStockList : "✅ สต็อกปกติ"}`;
  sendTelegram(CHAT_ID, msg);
}

function onEdit(e) {

  const range = e.range;
  const sheet = range.getSheet();

  if (sheet.getName() !== "STOCK" || range.getRow() <= 1) return;

  const row = range.getRow();
  const col = range.getColumn();
  const editor = Session.getActiveUser().getEmail() || "Authorized User";

  const materialCode = sheet.getRange(row, 2).getValue();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colName = headers[col - 1] || "Unknown";

  const oldValue = e.oldValue || "ว่าง";
  const newValue = e.value || "ถูกลบ";

  // แจ้งเตือนการแก้ไข
  const msg =
`📝 <b>Manual Edit Alert</b>
━━━━━━━━━━━━━━━
🔖 รหัส: ${materialCode}
📍 ช่องที่แก้: ${colName}
❌ เดิม: ${oldValue}
✅ ใหม่: <b>${newValue}</b>
👤 โดย: ${editor}`;

  sendTelegram(CHAT_ID, msg);


  // ตรวจสอบ LOW STOCK (เฉพาะเมื่อแก้คอลัมน์ QTY)
  if (col === 5) {
    checkLowStockAlert(row);
  }

}


// ตรวจสอบของใกล้หมด
function checkLowStockAlert(row) {

  // ป้องกัน error ถ้า row ว่าง
  if (!row || isNaN(row)) {
    Logger.log("Invalid row: " + row);
    return;
  }

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("STOCK");

  const data = sheet
    .getRange(row, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  const code = data[0];
  const desc = data[2];
  const qty = Number(data[4]);
  const alertStatus = data[10]; // คอลัมน์ ALERT

  // ถ้าใกล้หมดและยังไม่เคยแจ้งเตือน
  if (qty <= LOW_STOCK_LIMIT && alertStatus !== "ALERTED") {

    const msg =
`🚨 <b>LOW STOCK ALERT</b>
━━━━━━━━━━━━━━━
🔖 Code: ${code}
📦 Item: ${desc}
⚠️ Remaining: <b>${qty}</b>
━━━━━━━━━━━━━━━
กรุณาเติมสินค้า`;

    sendTelegram(CHAT_ID, msg);

    // บันทึกว่าแจ้งเตือนแล้ว
    sheet.getRange(row, 11).setValue("ALERTED");

  }

  // ถ้าเติมสินค้าแล้ว reset alert
  if (qty > LOW_STOCK_LIMIT && alertStatus === "ALERTED") {
    sheet.getRange(row, 11).setValue("");
  }

}

function setWebhook() {
  const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/setWebhook?url=${WEB_APP_URL}`;
  const response = UrlFetchApp.fetch(url);
  Logger.log("Webhook Set: " + response.getContentText());
}

function exportStockPDF(chatId) { // ตรวจสอบว่ามี (chatId) ในวงเล็บ
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("STOCK");
  
  const url = ss.getUrl().replace(/edit$/, '') +
    'export?format=pdf&portrait=true&size=A4&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=true&gid=' 
    + sheet.getSheetId();

  const token = ScriptApp.getOAuthToken();

  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });

  const pdfBlob = response.getBlob().setName("Stock_Report.pdf");

  // ส่งค่า chatId ต่อไปที่ฟังก์ชันส่งไฟล์ด้วย
  sendPDFToTelegram(chatId, pdfBlob); 
}

function sendPDFToTelegram(chatId, pdfBlob) {
  const url = "https://api.telegram.org/bot" + TELEGRAM_TOKEN + "/sendDocument";

  // 🔥 กัน error ถ้า chatId undefined
  const finalChatId = chatId ? chatId.toString() : CHAT_ID;

  const payload = {
    "chat_id": finalChatId,
    "document": pdfBlob,
    "caption": "📦 Stock Report PDF"
  };

  const options = {
    "method": "post",
    "payload": payload,
    "muteHttpExceptions": true
  };

  UrlFetchApp.fetch(url, options);
}

// ============================================================
// 8. EMAIL AUTOMATION (ส่งเมลอัตโนมัติ)
// ============================================================

// ⚙️ ตั้งค่าอีเมลผู้รับ (ถ้ามีหลายคนให้คั่นด้วยลูกน้ำ , )
const EMAIL_RECIPIENTS = "okumakung2018@gmail.com"; 

function sendDailyEmailWithPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("STOCK");
  
  if (!sheet) return;

  // 1️⃣ รวบรวมข้อมูลสรุป
  let totalQty = 0;
  let lowItems = 0;
  let lowStockList = "";
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    let qty = Number(data[i][4]) || 0; 
    totalQty += qty;
    
    if (qty <= LOW_STOCK_LIMIT) {
      lowItems++;
      lowStockList += `- [${data[i][0]}] ${data[i][2]} (คงเหลือ: ${qty})\n`;
    }
  }
  
  // 2️⃣ สร้างเนื้อหาอีเมล (Body)
  let emailBody = `สวัสดีครับ,\n\nนี่คือสรุปรายงานคลังสินค้าประจำวันที่ ${new Date().toLocaleDateString('th-TH')}\n\n`;
  emailBody += `📦 รายการสินค้าทั้งหมด: ${data.length - 1} รายการ\n`;
  emailBody += `🔢 จำนวนชิ้นรวมในคลัง: ${totalQty} ชิ้น\n`;
  emailBody += `⚠️ สินค้าใกล้หมด: ${lowItems} รายการ\n\n`;
  
  if (lowItems > 0) {
    emailBody += `🚨 รายการที่ต้องสั่งซื้อเพิ่ม:\n${lowStockList}\n`;
  }
  
  emailBody += `\nกรุณาตรวจสอบรายละเอียดแบบเต็มในไฟล์ PDF ที่แนบมานี้ครับ\n\nด้วยความเคารพ,\nระบบจัดการสต็อกอัตโนมัติ`;

  // 3️⃣ สร้างไฟล์ PDF สต็อกล่าสุด
  const url = ss.getUrl().replace(/edit$/, '') +
    'export?format=pdf&portrait=true&size=A4&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=true&gid=' 
    + sheet.getSheetId();
    
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token }
  });
  
  const dateStr = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  const pdfBlob = response.getBlob().setName(`Daily_Stock_Report_${dateStr}.pdf`);
  
  // 4️⃣ ส่งอีเมล
  MailApp.sendEmail({
    to: EMAIL_RECIPIENTS,
    subject: `📊 สรุปรายงานคลังสินค้าประจำวัน (${dateStr})`,
    body: emailBody,
    attachments: [pdfBlob]
  });
  
  Logger.log("✅ ส่งอีเมลสำเร็จ!");
}

/* REPORT Func*/
function sendDailyStockPDF() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("STOCK");

  const url = ss.getUrl().replace(/edit$/, '') +
  'export?format=pdf&portrait=true&size=A4&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=true&gid=' 
  + sheet.getSheetId();

  const token = ScriptApp.getOAuthToken();

  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + token
    }
  });

  const pdfBlob = response.getBlob().setName("Daily_Stock_Report.pdf");

  const telegramUrl = "https://api.telegram.org/bot" + TELEGRAM_TOKEN + "/sendDocument";

  const payload = {
    chat_id: CHAT_ID,
    document: pdfBlob,
    caption: "📊 Daily Stock Report"
  };

  UrlFetchApp.fetch(telegramUrl, {
    method: "post",
    payload: payload
  });

}

// ============================================================
// 📊 REPORT DASHBOARD
// ฟังก์ชันนี้ใช้กับคำสั่ง /report
// แสดงภาพรวมคลังสินค้า
// ============================================================
function report(){

  // ดึงข้อมูลจาก Sheet STOCK
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();

  // ตัวแปรสำหรับสรุป
  let totalItems = data.length - 1; // จำนวนสินค้า (ไม่รวม header)
  let totalQty = 0;                 // จำนวนชิ้นรวม
  let lowStock = 0;                 // จำนวนสินค้าที่ใกล้หมด
  let outStock = 0;                 // จำนวนสินค้าที่หมด

  // วนลูปอ่านข้อมูลแต่ละสินค้า
  for(let i = 1; i < data.length; i++){

    // QTY อยู่ column 5
    let qty = Number(data[i][4]) || 0;

    totalQty += qty;

    // ตรวจสอบสินค้าใกล้หมด
    if(qty <= LOW_STOCK_LIMIT){
      lowStock++;
    }

    // ตรวจสอบสินค้าหมด
    if(qty === 0){
      outStock++;
    }

  }

  // ส่งข้อความกลับ Telegram
  return `
📊 <b>STOCK DASHBOARD</b>
━━━━━━━━━━━━━━━
📦 รายการสินค้า : ${totalItems}

📊 จำนวนชิ้นรวม : ${totalQty}

⚠️ ใกล้หมด : ${lowStock}

🚨 หมดสต็อก : ${outStock}
━━━━━━━━━━━━━━━
`;

}



// ============================================================
// ⚠️ LOW STOCK LIST
// ฟังก์ชันนี้ใช้กับคำสั่ง /lowstock
// แสดงรายการสินค้าที่ใกล้หมด
// ============================================================
function lowStock(){

  // ดึงข้อมูลจาก STOCK
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();

  // ข้อความเริ่มต้น
  let msg = "⚠️ <b>LOW STOCK LIST</b>\n━━━━━━━━━━━━━━━\n";

  let found = false;

  // วนลูปตรวจสอบสินค้า
  for(let i = 1; i < data.length; i++){

    let qty = Number(data[i][4]) || 0;

    // ถ้า QTY น้อยกว่าค่าที่กำหนด
    if(qty <= LOW_STOCK_LIMIT){

      msg += `🔖 ${data[i][0]} | ${data[i][2]}\n`;
      msg += `📊 เหลือ : <b>${qty}</b>\n\n`;

      found = true;

    }

  }

  // ถ้าไม่มีสินค้าใกล้หมด
  if(!found){
    return "✅ ไม่มีสินค้าใกล้หมด";
  }

  return msg;

}
