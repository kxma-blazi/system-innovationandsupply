// ============================================================
// 1. CONFIGURATION
// ============================================================
const TELEGRAM_TOKEN = "เเปด327163778:AAFM4aKpxT29WTB4z_StzKEEcRGrSBDS2_s";
const CHAT_ID = "-หนึ่ง003731290917"; 
const LOW_STOCK_LIMIT = 1;
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbw5KR2d8ukXl0DVUAj8Lo11JQ8F00vL-ax7pkqRbDeSv1NTLNIRy094narUUyz_TbXH/exec";


// ============================================================
// 2. CORE SYSTEM (Support ทั้งพิมพ์ และ กดปุ่ม)
// ============================================================

function doPost(e) {
  try {
    if (!e || !e.postData) return;
    const data = JSON.parse(e.postData.contents);

    // 1️⃣ CALLBACK BUTTON (เมื่อผู้ใช้กดปุ่ม Inline Keyboard)
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
      else if (callbackData === "/exportpdf") {
        sendTelegram(chatId, "📄 กำลังสร้าง PDF...");
        exportStockPDF(chatId);
        answerCallback(callbackId);
        return;
      } 
      else if (callbackData === "/menu") {
        sendControlPanel(chatId);
        answerCallback(callbackId);
        return;
      }
      else if (callbackData === "/lowstock") { responseText = lowStock(); } 
      else if (callbackData === "/report") { responseText = report(); } 
      else if (callbackData === "/history") { responseText = history(); } 
      else if (callbackData === "/help") { responseText = helpMenu(); }

      if (responseText) { sendTelegram(chatId, responseText); }
      answerCallback(callbackId);
      return;
    }

    // 2️⃣ TEXT COMMAND (เมื่อผู้ใช้พิมพ์ข้อความหรือกดคำสั่งจากเมนู)
    if (!data.message || !data.message.text) return;
    const chatId = data.message.chat.id;
    const text = data.message.text;
    const user = data.message.from.first_name || data.message.from.username || "Unknown";
    
    // ตัดคำสั่งให้เป็นตัวเล็ก และตัดชื่อบอทออก (เช่น /report@mybot -> /report)
    let cmd = text.split(" ")[0].toLowerCase();
    if (cmd.includes("@")) { cmd = cmd.split("@")[0]; }
    const args = text.split(" ");

    switch (cmd) {
      case "/exportpdf": 
        sendTelegram(chatId, "📄 กำลังสร้าง PDF..."); 
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

      case "/report": 
        // สั่งให้ทำสรุปรายงาน แล้วดึงข้อความมาส่งให้ user
        let reportMsg = report(); 
        sendTelegram(chatId, reportMsg); 
        break;

      case "/stock": 
        sendTelegram(chatId, getStock(args[1])); 
        break;

      case "/search": 
        sendTelegram(chatId, searchStock(args.slice(1).join(" "))); 
        break;

      case "/allstock": 
        let page = args[1] ? Number(args[1]) : 1;
        sendStockPage(chatId, isNaN(page) ? 1 : page); 
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

      case "/dp-cbr": sendTelegram(chatId, withdraw(args[1], Number(args[2]), 5, "CBR", user)); break;
      case "/dp-ccs": sendTelegram(chatId, withdraw(args[1], Number(args[2]), 6, "CCS", user)); break;
      case "/dp-sko": sendTelegram(chatId, withdraw(args[1], Number(args[2]), 7, "SKO", user)); break;
      case "/dp-ryg": sendTelegram(chatId, withdraw(args[1], Number(args[2]), 8, "RYG", user)); break;
      case "/dp-trt": sendTelegram(chatId, withdraw(args[1], Number(args[2]), 9, "TRT", user)); break;

      default:
        // ถ้าขึ้นต้นด้วย / แต่หาคำสั่งไม่เจอ (พิมพ์ผิด)
        if (cmd.startsWith("/")) {
          sendTelegram(chatId, "❌ ไม่รู้จักคำสั่งนี้ครับ\nลองเลือกจากเมนูด้านล่างนี้ดูนะครับ 👇");
          // ส่งเมนูหลักให้เขาเลย จะได้กดปุ่มแทน
          sendControlPanel(chatId); 
        }
        break;
    }

  } catch (err) {
    Logger.log("Error in doPost: " + err.toString());
  }
} // ปิด doPost

// ============================================================
// 3. UI & UTILITIES
// ============================================================

function sendControlPanel(chatId) {
  const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`;
  const keyboard = {
    "inline_keyboard": [
      [{ "text": "📦 ดูสต็อกทั้งหมด", "callback_data": "/allstock" }, { "text": "⚠️ ของใกล้หมด", "callback_data": "/lowstock" }],
      [{ "text": "📊 สรุปรายงาน", "callback_data": "/report" }, { "text": "📈 ดู Dashboard", "url": "https://lookerstudio.google.com/reporting/468488a2-94a5-4db9-9945-8b6f1e2008dd" }],
      [{ "text": "📜 ประวัติล่าสุด", "callback_data": "/history" }, { "text": "📄 ออกไฟล์ PDF", "callback_data": "/exportpdf" }],
      [{ "text": "📤 วิธีเบิก/เติม", "callback_data": "/help" }, { "text": "🏠 กลับเมนูหลัก", "callback_data": "/menu" }]
    ]
  };
  const payload = {
    "chat_id": chatId ? chatId.toString() : CHAT_ID,
    "text": "🎮 <b>SYSTEM INNOVATION AND SUPPLY</b>\n\nเลือกรายการที่ต้องการจัดการด้านล่างครับ:",
    "parse_mode": "HTML",
    "reply_markup": JSON.stringify(keyboard) 
  };
  UrlFetchApp.fetch(url, { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true });
}

function answerCallback(callbackId) {
  if (!callbackId) return;
  const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/answerCallbackQuery`;
  try { UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify({ callback_query_id: callbackId }), muteHttpExceptions: true }); } catch (e) {}
}

function sendTelegram(chatId, msg) {
  const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`;
  UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify({ chat_id: chatId, text: msg, parse_mode: "HTML" }), muteHttpExceptions: true });
}

// ============================================================
// 4. STOCK LOGIC
// ============================================================

function getSheet() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("STOCK"); }

function restock(code, qty, user) {
  if (!code || isNaN(qty) || qty <= 0) return "⚠️ รูปแบบผิด: /restock [CODE] [QTY]";
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === code.toLowerCase()) {
      let newQty = (Number(data[i][4]) || 0) + qty;
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
      let currentStock = Number(data[i][4]) || 0;
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

function writeLog(user, action, code, detail, sheetName, remain) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("Logs");
  if (!logSheet) return;
  logSheet.appendRow([new Date(), user, action, code, detail, sheetName, remain]);
}

function notifyActivity(title, code, type, qty, remain, dept = "", user = "") {
  const statusEmoji = type === "IN" ? "🟢 [ เติมเข้า ]" : "🟠 [ เบิกออก ]";
  const mathSign = type === "IN" ? "➕" : "➖";
  let msg = `📌 <b>${title}</b>\n━━━━━━━━━━━━━━━━\n🆔 <b>รหัส:</b> <code>${code}</code>\n📝 <b>สถานะ:</b> ${statusEmoji}\n🔢 <b>จำนวน:</b> ${mathSign} <b>${qty}</b>\n`;
  if (dept) msg += `🏢 <b>แผนก:</b> ${dept}\n`;
  msg += `👤 <b>โดย:</b> ${user}\n━━━━━━━━━━━━━━━━\n📊 <b>คงเหลือ:</b> <b><u>${remain}</u></b> หน่วย`;
  sendTelegram(CHAT_ID, msg);
}

// ============================================================
// 5. REPORTS & QUERIES
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

function report() {
  const data = getSheet().getDataRange().getValues();
  let totalItems = data.length - 1, totalQty = 0, lowCount = 0, outStock = 0;
  for (let i = 1; i < data.length; i++) {
    let qty = Number(data[i][4]) || 0;
    totalQty += qty;
    if (qty <= LOW_STOCK_LIMIT && qty > 0) lowCount++;
    if (qty === 0) outStock++;
  }
  return `📊 <b>STOCK DASHBOARD</b>\n━━━━━━━━━━━━━━━\n📦 รายการ: ${totalItems}\n📊 ชิ้นรวม: ${totalQty}\n⚠️ ใกล้หมด: ${lowCount}\n🚨 หมดสต็อก: ${outStock}\n━━━━━━━━━━━━━━━`;
}

function lowStock() {
  const data = getSheet().getDataRange().getValues();
  let msg = "⚠️ <b>LOW STOCK LIST</b>\n━━━━━━━━━━━━━━━\n", found = false;
  for (let i = 1; i < data.length; i++) {
    let qty = Number(data[i][4]) || 0;
    if (qty <= LOW_STOCK_LIMIT) {
      msg += `🔖 ${data[i][0]} | ${data[i][2]}\n📊 เหลือ: <b>${qty}</b>\n\n`;
      found = true;
    }
  }
  return found ? msg : "✅ ไม่มีสินค้าใกล้หมด";
}

function history() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs");
  if (!sheet) return "❌ ไม่พบชีต Logs";
  const data = sheet.getDataRange().getValues();
  let msg = "📜 <b>History (10 รายการล่าสุด)</b>\n\n", start = Math.max(1, data.length - 10);
  for (let i = start; i < data.length; i++) {
    let date = Utilities.formatDate(new Date(data[i][0]), "GMT+7", "dd/MM HH:mm");
    msg += `🕒 ${date} | 👤 ${data[i][1]}\n📦 ${data[i][3]} (${data[i][2]}) ${data[i][4]} หน่วย\n📊 เหลือ ${data[i][6]}\n\n`;
  }
  return msg;
}

function topWithdraw() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs");
  if (!sheet) return "❌ ไม่พบชีต Logs";
  const data = sheet.getDataRange().getValues();
  let withdrawMap = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === "Withdraw") {
      let code = data[i][3];
      withdrawMap[code] = (withdrawMap[code] || 0) + (Number(data[i][4]) || 0);
    }
  }
  const sorted = Object.entries(withdrawMap).sort((a, b) => b[1] - a[1]).slice(0, 10);
  if (sorted.length === 0) return "📊 ยังไม่มีข้อมูลการเบิก";
  let msg = "🏆 <b>Top 10 การเบิกสูงสุด</b>\n━━━━━━━━━━━━\n";
  sorted.forEach((item, idx) => { msg += `${idx + 1}. 📦 <code>${item[0]}</code>: <b>${item[1]}</b> ชิ้น\n`; });
  return msg;
}

// ============================================================
// 6. PDF & WEBHOOK
// ============================================================

function exportStockPDF(chatId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("STOCK");
  const url = ss.getUrl().replace(/edit$/, '') + 'export?format=pdf&portrait=true&size=A4&gridlines=true&gid=' + sheet.getSheetId();
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
  sendPDFToTelegram(chatId, response.getBlob().setName("Stock_Report.pdf"));
}

function sendPDFToTelegram(chatId, pdfBlob) {
  const payload = { "chat_id": chatId ? chatId.toString() : CHAT_ID, "document": pdfBlob, "caption": "📦 Stock Report PDF" };
  UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendDocument`, { "method": "post", "payload": payload, "muteHttpExceptions": true });
}

function setWebhook() {
  const response = UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/setWebhook?url=${WEB_APP_URL}`);
  Logger.log("Webhook Set: " + response.getContentText());
}

// ============================================================
// 7. TRIGGERS (onEdit - ปรับปรุงเพื่อรองรับหลายชีต)
// ============================================================

function onEdit(e) {
  try {
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    const row = e.range.getRow();
    const col = e.range.getColumn();
    
    // 1. ยกเว้นเฉพาะชีต "Logs" เพื่อป้องกันการแจ้งเตือนวนลูป
    if (sheetName === "Logs") return;

    // 2. ข้อมูลผู้ใช้งานและค่าที่เปลี่ยนแปลง
    const user = Session.getActiveUser().getEmail() || "Staff";
    const oldValue = e.oldValue || "ว่าง/ไม่มีข้อมูลเดิม";
    const newValue = e.value || "ถูกลบ/ค่าว่าง";

    // ถ้าค่าใหม่กับค่าเดิมเหมือนกันเป๊ะ (เช่น กด Enter เฉยๆ) ไม่ต้องส่งเตือน
    if (oldValue == newValue) return;

    // 3. ดึงข้อมูลหัวตาราง (Header) ของคอลัมน์ที่ถูกแก้ เพื่อให้รู้ว่าแก้ช่องอะไร
    const headerName = sheet.getRange(1, col).getValue() || "คอลัมน์ที่ " + col;

    // 4. ดึงข้อมูลอ้างอิงในแถวนั้น (ดึง 3 คอลัมน์แรกมาโชว์เพื่อให้รู้ว่าเป็นรายการไหน)
    const referenceData = sheet.getRange(row, 1, 1, 3).getValues()[0];
    const refText = referenceData.filter(String).join(" | ") || "ไม่พบข้อมูลอ้างอิง";

    // 5. บันทึก Log ลงชีต Logs (เก็บไว้ดูย้อนหลัง)
    if (typeof writeLog === 'function') {
      writeLog(user, `Edit [${sheetName}]`, refText, `${headerName}: ${oldValue} -> ${newValue}`, sheetName, "N/A");
    }

    // 6. ส่งข้อความเข้า Telegram
    let msg = `🔔 <b>แจ้งเตือนการแก้ไข [${sheetName}]</b>\n`;
    msg += `━━━━━━━━━━━━━━━━\n`;
    msg += `📍 <b>ตำแหน่ง:</b> แถว ${row} | ${headerName}\n`;
    msg += `📦 <b>รายการอ้างอิง:</b> <code>${refText}</code>\n`;
    msg += `❌ <b>เดิม:</b> ${oldValue}\n`;
    msg += `✅ <b>ใหม่:</b> <b>${newValue}</b>\n`;
    msg += `👤 <b>แก้ไขโดย:</b> ${user}\n`;
    msg += `━━━━━━━━━━━━━━━━`;

    sendTelegram(CHAT_ID, msg);

  } catch (err) {
    Logger.log("Error: " + err.toString());
  }
}


function checkLowStockAlert(row) {
  const sheet = getSheet();
  const rowData = sheet.getRange(row, 1, 1, 5).getValues()[0];
  const code = rowData[1]; // คอลัมน์ B
  const desc = rowData[2]; // คอลัมน์ C
  const qty = Number(rowData[4]); // คอลัมน์ E

  if (qty <= LOW_STOCK_LIMIT) {
    let msg = `🚨 <b>LOW STOCK ALERT</b>\n`;
    msg += `━━━━━━━━━━━━━━━━\n`;
    msg += `🔖 <b>Code:</b> <code>${code}</code>\n`;
    msg += `📦 <b>Item:</b> ${desc}\n`;
    msg += `⚠️ <b>คงเหลือ:</b> <b>${qty}</b> หน่วย\n`;
    msg += `━━━━━━━━━━━━━━━━`;
    sendTelegram(CHAT_ID, msg);
  }
}

function helpMenu() {
  return `🧭<b>STOCK GUIDE</b>\n━━━━━━━━━━\n👉 <code>/menu</code> - เมนูหลัก\n📦 <code>/stock CODE</code> - เช็คสต็อก\n➕ <code>/restock CODE QTY</code> - เติมของ\n📤 <code>/dp-แผนก CODE QTY</code> - เบิกของ\n📊 <code>/report</code> - สรุปคลัง`;
}

function sendStockPage(chatId, page = 1) {
  const data = getSheet().getDataRange().getValues();
  const perPage = 20, totalPages = Math.ceil((data.length - 1) / perPage);
  if (page < 1) page = 1; if (page > totalPages) page = totalPages;
  const start = (page - 1) * perPage + 1, end = Math.min(start + perPage, data.length);

  let msg = `📦 <b>STOCK LIST (${page}/${totalPages})</b>\n<pre>ID   QTY  DESCRIPTION\n`;
  for (let i = start; i < end; i++) {
    msg += `${(data[i][0]||"").toString().padEnd(4)} ${(data[i][4]||0).toString().padEnd(4)} ${(data[i][2]||"").substring(0,20)}\n`;
  }
  msg += "</pre>";

  let row = [];
  if (page > 1) row.push({ text: "⬅️", callback_data: `/allstock_${page-1}` });
  if (page < totalPages) row.push({ text: "➡️", callback_data: `/allstock_${page+1}` });
  
  sendTelegram(chatId, msg); // ปรับการส่งแบบง่ายเพื่อลด error
}

function dailyAutoReport() {
  const data = getSheet().getDataRange().getValues();
  let totalItems = data.length - 1;
  let lowItems = [];
  let outOfStock = 0;

  for (let i = 1; i < data.length; i++) {
    let qty = Number(data[i][4]) || 0;
    if (qty === 0) outOfStock++;
    else if (qty <= LOW_STOCK_LIMIT) {
      lowItems.push(`- ${data[i][0]}: ${data[i][2]} (เหลือ ${qty})`);
    }
  }

  let msg = `📢 <b>สรุปยอดรายวัน (${Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy")})</b>\n`;
  msg += `━━━━━━━━━━━━━━━\n`;
  msg += `📦 รายการทั้งหมด: ${totalItems} รายการ\n`;
  msg += `🚨 สินค้าหมด: ${outOfStock} รายการ\n`;
  msg += `⚠️ สินค้าใกล้หมด: ${lowItems.length} รายการ\n`;
  
  if (lowItems.length > 0) {
    msg += `\n<b>รายการที่ต้องเติมด่วน:</b>\n${lowItems.join("\n")}\n`;
  }
  
  msg += `━━━━━━━━━━━━━━━\n`;
  msg += ` `;

  sendTelegram(CHAT_ID, msg);
}

// ============================================================
// 9. DASHBOARD AUTO PUSH
// ============================================================
// เช็ค ข้อมูลเปลี่ยนไหม
function pushDashboardSmart() {

  const cache = CacheService.getScriptCache();
  const last = cache.get("dashboard");

  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();

  let totalQty = 0;

  for (let i = 1; i < data.length; i++) {
    totalQty += Number(data[i][4]) || 0;
  }

  if (last == totalQty.toString()) {
    return; // ❌ ไม่เปลี่ยน = ไม่ส่ง
  }

  cache.put("dashboard", totalQty.toString(), 21600); // 6 ชม

  pushDashboard(); // ยิงจริง
}
//ส่ง Noti
function pushDashboard() {

  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();

  let totalItems = data.length - 1;
  let totalQty = 0;
  let lowStock = 0;
  let outStock = 0;

  for (let i = 1; i < data.length; i++) {
    let qty = Number(data[i][4]) || 0;

    totalQty += qty;

    if (qty <= LOW_STOCK_LIMIT && qty > 0) lowStock++;
    if (qty === 0) outStock++;
  }

  let msg = `📊 <b>DASHBOARD UPDATE</b>\n`;
  msg += `━━━━━━━━━━━━━━━\n`;
  msg += `📦 รายการสินค้า: ${totalItems}\n`;
  msg += `📊 จำนวนรวม: ${totalQty}\n`;
  msg += `⚠️ ใกล้หมด: ${lowStock}\n`;
  msg += `🚨 หมดสต็อก: ${outStock}\n`;
  msg += `━━━━━━━━━━━━━━━`;

  sendTelegram(CHAT_ID, msg);
}

// Send Mail (แนบทั้ง PDF และ Excel ที่ใช้งานกับ Google Sheets ได้)
function sendDailyEmailWithPDF() {
  try {
    DriveApp.getRootFolder(); // ขอ permission

    // ✅ รองรับหลายอีเมลเพราะใส่คั่นด้วย "," (เพิ่มด้วย)
    const recipients = [
      "okumakung2018@gmail.com",
      "s65122250014@ssru.ac.th",
      // "user3@gmail.com"
    ].join(",");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("STOCK");

    if (!sheet) throw new Error("ไม่พบชีต STOCK");

    const ssId = ss.getId();
    const sheetId = sheet.getSheetId();
    const dateStr = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy");
    const token = ScriptApp.getOAuthToken();

    // ✅ PDF (ปรับ format ให้สวยขึ้น)
    const pdfUrl = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=pdf&gid=${sheetId}&size=A4&portrait=true&fitw=true&gridlines=false`;

    // ✅ Excel
    const excelUrl = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=xlsx&gid=${sheetId}`;

    const options = {
      headers: { Authorization: "Bearer " + token },
      muteHttpExceptions: true
    };

    const pdfResponse = UrlFetchApp.fetch(pdfUrl, options);
    const excelResponse = UrlFetchApp.fetch(excelUrl, options);

    // ❌ กันไฟล์โหลดพัง
    if (pdfResponse.getResponseCode() !== 200) {
      throw new Error("PDF export ล้มเหลว");
    }

    if (excelResponse.getResponseCode() !== 200) {
      throw new Error("Excel export ล้มเหลว");
    }

    MailApp.sendEmail({
      to: recipients,
      subject: `📊 รายงานสต็อกประจำวัน - ${dateStr}`,
      name: "🤖 SYSTEM INNOVATION AND SUPPLY",
      htmlBody: `
        <h2>📦 รายงานสต็อกประจำวัน</h2>
        <p>วันที่: <b>${dateStr}</b></p>

        <ul>
          <li>📄 แนบไฟล์ PDF (สำหรับดู)</li>
          <li>📊 แนบไฟล์ Excel (สำหรับใช้งานต่อ)</li>
        </ul>

        <p>
          🔗 <a href="${ss.getUrl()}">เปิดดู Google Sheets</a>
        </p>

        <hr>
        <p style="color:gray;">ระบบอัตโนมัติ</p>
      `,
      attachments: [
        pdfResponse.getBlob().setName(`Stock_Report_${dateStr}.pdf`),
        excelResponse.getBlob().setName(`Stock_Data_${dateStr}.xlsx`)
      ]
    });

    Logger.log("✅ ส่งเมลสำเร็จ");

  } catch (err) {
    Logger.log("❌ ERROR: " + err.message);
  }
}


// TEST CASE
function createTestCaseSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // ลบ sheet เก่าถ้ามี แล้วสร้างใหม่
  let sheet = ss.getSheetByName("TestCases");
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet("TestCases");

  // ============ HEADER ============
  const headers = ["#", "หมวด", "Test Case", "Input", "Expected Output", "Type", "ผลการทดสอบ", "หมายเหตุ"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground("#2B2B2B")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // ============ DATA ============
  const rows = [
    // TC-01 /stock
    ["TC-01-1", "/stock", "ค้นหา code ที่มีในระบบ",         "/stock Q001",         "แสดง Code, Desc, สต็อก",              "Positive", "", ""],
    ["TC-01-2", "/stock", "ค้นหาด้วยตัวพิมพ์เล็ก",          "/stock q001",         "ผลเหมือน TC-01-1 (case-insensitive)", "Edge",     "", ""],
    ["TC-01-3", "/stock", "ค้นหา code ที่ไม่มีในระบบ",       "/stock ZZZ",          "❌ ไม่พบข้อมูล",                      "Negative", "", ""],
    ["TC-01-4", "/stock", "ไม่ใส่ code",                    "/stock",              "⚠️ ระบุรหัส: /stock Q01",             "Negative", "", ""],

    // TC-02 /restock
    ["TC-02-1", "/restock", "เติมสินค้าปกติ",               "/restock Q001 10",    "✅ Restock Success, stock +10",        "Positive", "", ""],
    ["TC-02-2", "/restock", "จำนวน 0",                      "/restock Q001 0",     "⚠️ รูปแบบผิด",                        "Negative", "", ""],
    ["TC-02-3", "/restock", "จำนวนติดลบ",                   "/restock Q001 -5",    "⚠️ รูปแบบผิด",                        "Negative", "", ""],
    ["TC-02-4", "/restock", "ไม่ใส่ parameter",             "/restock",            "⚠️ รูปแบบผิด",                        "Negative", "", ""],
    ["TC-02-5", "/restock", "code ไม่มีในระบบ",             "/restock ZZZ 5",      "❌ ไม่พบรหัสสินค้า",                  "Negative", "", ""],

    // TC-03 /dp-xxx
    ["TC-03-1", "/dp-xxx", "เบิก CBR ปกติ",                 "/dp-cbr Q001 2",      "✅ Success (CBR), stock -2",           "Positive", "", ""],
    ["TC-03-2", "/dp-xxx", "เบิกเกินจำนวนที่มี",            "/dp-cbr Q001 9999",   "❌ สต็อกไม่พอ",                       "Negative", "", ""],
    ["TC-03-3", "/dp-xxx", "เบิก CCS ปกติ",                 "/dp-ccs Q001 1",      "✅ Success (CCS)",                     "Positive", "", ""],
    ["TC-03-4", "/dp-xxx", "จำนวน 0",                       "/dp-sko Q001 0",      "⚠️ รูปแบบผิด",                        "Negative", "", ""],
    ["TC-03-5", "/dp-xxx", "code ไม่มีในระบบ",              "/dp-cbr ZZZ 1",       "❌ ไม่พบรหัสสินค้า",                  "Negative", "", ""],

    // TC-04 /allstock
    ["TC-04-1", "/allstock", "กดปุ่มดูสต็อกทั้งหมด",        "กดปุ่ม ดูสต็อกทั้งหมด", "หน้า 1, มีปุ่ม ➡️ และ 🏠",         "Positive", "", ""],
    ["TC-04-2", "/allstock", "กดปุ่มถัดไป",                 "กดปุ่ม ➡️",           "หน้า 2, มีปุ่ม ⬅️ ➡️ และ 🏠",        "Positive", "", ""],
    ["TC-04-3", "/allstock", "กดปุ่มก่อนหน้า",              "กดปุ่ม ⬅️",           "กลับหน้า 1",                          "Positive", "", ""],
    ["TC-04-4", "/allstock", "กดปุ่มกลับเมนู",              "กดปุ่ม 🏠",           "แสดง Control Panel",                  "Positive", "", ""],
    ["TC-04-5", "/allstock", "อยู่หน้าสุดท้าย",             "กดถัดไปจนหน้าสุดท้าย","ไม่มีปุ่ม ➡️",                        "Edge",     "", ""],

    // TC-05 /lowstock
    ["TC-05-1", "/lowstock", "มีสินค้า stock ≤ 1",          "/lowstock",           "แสดงรายการสินค้าใกล้หมด",            "Positive", "", ""],
    ["TC-05-2", "/lowstock", "ไม่มีสินค้า stock ≤ 1",       "/lowstock",           "✅ ไม่มีสินค้าใกล้หมด",               "Edge",     "", ""],

    // TC-06 /report & /history
    ["TC-06-1", "/report",  "ดูรายงานสรุป",                 "/report",             "แสดงยอดรวม, ใกล้หมด, หมดสต็อก",     "Positive", "", ""],
    ["TC-06-2", "/history", "ดูประวัติล่าสุด",              "/history",            "แสดง log 10 รายการล่าสุด",           "Positive", "", ""],
    ["TC-06-3", "/history", "ตอน Logs ว่างเปล่า",           "/history",            "แสดงหัว ไม่ error",                   "Edge",     "", ""],

    // TC-07 คำสั่งผิด
    ["TC-07-1", "unknown cmd", "พิมพ์คำสั่งที่ไม่มี",       "/xyz",                "❌ ไม่รู้จักคำสั่ง + แสดง Control Panel", "Negative", "", ""],
    ["TC-07-2", "unknown cmd", "ส่งข้อความธรรมดา",          "hello",               "ไม่ตอบสนอง ไม่ crash",               "Edge",     "", ""],
    ["TC-07-3", "unknown cmd", "คำสั่งมีชื่อบอทต่อท้าย",   "/report@botname",     "ตัดชื่อบอทออก ทำงานปกติ",            "Edge",     "", ""],

    // TC-08 onEdit
    ["TC-08-1", "onEdit", "แก้ค่าใน Sheet STOCK",           "แก้ cell ใน STOCK",   "แจ้งเตือน Telegram แถว/คอลัมน์/ค่า", "Positive", "", ""],
    ["TC-08-2", "onEdit", "แก้ค่าใน Sheet Logs",            "แก้ cell ใน Logs",    "ไม่แจ้งเตือน (กัน loop)",            "Edge",     "", ""],
    ["TC-08-3", "onEdit", "กด Enter เฉยๆ ค่าไม่เปลี่ยน",   "กด Enter",            "ไม่แจ้งเตือน",                        "Edge",     "", ""],
  ];

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  // ============ STYLING ============
  // สีแถบสลับ
  for (let i = 0; i < rows.length; i++) {
    const bg = i % 2 === 0 ? "#F9F9F9" : "#FFFFFF";
    sheet.getRange(i + 2, 1, 1, headers.length).setBackground(bg);
  }

  // สีคอลัมน์ Type
  const typeCol = 6;
  for (let i = 0; i < rows.length; i++) {
    const type = rows[i][5];
    const cell = sheet.getRange(i + 2, typeCol);
    if (type === "Positive") { cell.setBackground("#D4EDDA").setFontColor("#155724"); }
    else if (type === "Negative") { cell.setBackground("#F8D7DA").setFontColor("#721C24"); }
    else if (type === "Edge") { cell.setBackground("#FFF3CD").setFontColor("#856404"); }
    cell.setHorizontalAlignment("center").setFontWeight("bold");
  }

  // คอลัมน์ ผลการทดสอบ — dropdown
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["✅ Pass", "❌ Fail", "⏳ Pending"], true).build();
  sheet.getRange(2, 7, rows.length, 1).setDataValidation(rule);
  sheet.getRange(2, 7, rows.length, 1)
    .setValue("⏳ Pending")
    .setHorizontalAlignment("center");

  // ปรับขนาดคอลัมน์
  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 250);
  sheet.setColumnWidth(6, 90);
  sheet.setColumnWidth(7, 110);
  sheet.setColumnWidth(8, 150);

  // Freeze header
  sheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert("✅ สร้าง TestCases Sheet สำเร็จ!");
}
