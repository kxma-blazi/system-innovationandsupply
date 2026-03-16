// ============================================================
// 1. CONFIGURATION
// ============================================================
const TELEGRAM_TOKEN = "8327163778:AAFM4aKpxT29WTB4z_StzKEEcRGrSBDS2_s";
const CHAT_ID = "-100373129091SEVEN"; 
const LOW_STOCK_LIMIT = 1;
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbw4F_rnJrQmzQSZMBsNKVHLncWWmAzI3_qiodVaDOf6GkTo1XeAOFxXZQQSwyjF7ZGg/exec";

// ============================================================
// 2. CORE SYSTEM (Support ทั้งพิมพ์ และ กดปุ่ม)
// ============================================================
function doPost(e) {
  try {
    const contents = e.postData.contents;
    const data = JSON.parse(contents);

    // --- ส่วนที่ 1: จัดการการกดปุ่ม (Callback Query) ---
    // ส่วนนี้จะทำงานเมื่อ User "จิ้ม" ที่ปุ่มสี่เหลี่ยมใต้ข้อความ
    if (data.callback_query) {
      const callbackData = data.callback_query.data; 
      const chatId = data.callback_query.message.chat.id;
      const callbackId = data.callback_query.id; 
      
      let responseText = "";
      if (callbackData === "/allstock") responseText = allStock();
      else if (callbackData === "/lowstock") responseText = lowStock();
      else if (callbackData === "/report") responseText = report();
      else if (callbackData === "/history") responseText = history();
      else if (callbackData === "/help") responseText = helpMenu();

      if (responseText) sendTelegram(chatId, responseText);
      
      // แจ้ง Telegram ว่าได้รับคำสั่งแล้ว (ช่วยให้บอทไม่หมุนค้าง)
      answerCallback(callbackId); 
      return;
    }

    // --- ส่วนที่ 2: จัดการการพิมพ์ (Message) ---
    if (!data.message || !data.message.text) return;

    const chatId = data.message.chat.id;
    const text = data.message.text;
    const user = data.message.from.first_name || data.message.from.username || "Unknown";
    
    // --- จุดที่ปรับปรุง: แก้ปัญหาชื่อบอทพ่วงมา (เช่น /menu@botname) ---
    let cmd = text.split(" ")[0].toLowerCase();
    if (cmd.includes("@")) {
      cmd = cmd.split("@")[0]; // ตัด @botname ออก ให้เหลือแค่ตัวคำสั่ง
    }
    const args = text.split(" ");
    // ---------------------------------------------------------
    switch(cmd) {
      case "/exportpdf":
      sendTelegram(chatId,"📄 กำลังสร้าง PDF...");
      exportStockPDF();
      break;

      case "/start": 
      case "/menu": 
        sendControlPanel(chatId); 
        break;
      
      case "/help":
      case "/manual": 
        sendTelegram(chatId, helpMenu()); 
        sendControlPanel(chatId); // ส่งปุ่มเมนูตามไปให้ด้วยเลย เพื่อความสะดวก         
        break;

      case "/history": sendTelegram(chatId, history()); break;
      case "/stock": sendTelegram(chatId, getStock(args[1])); break;
      case "/search": sendTelegram(chatId, searchStock(args.slice(1).join(" "))); break;
      case "/allstock": sendTelegram(chatId, allStock()); break;
      case "/lowstock": sendTelegram(chatId, lowStock()); break;
      case "/report": sendTelegram(chatId, report()); break;
      case "/restock": sendTelegram(chatId, restock(args[1], Number(args[2]), user)); break;
      
      case "/dp-cbr": sendTelegram(chatId, withdraw(args[1], Number(args[2]), 5, "CBR", user)); break;
      case "/dp-ccs": sendTelegram(chatId, withdraw(args[1], Number(args[2]), 6, "CCS", user)); break;
      case "/dp-sko": sendTelegram(chatId, withdraw(args[1], Number(args[2]), 7, "SKO", user)); break;
      case "/dp-ryg": sendTelegram(chatId, withdraw(args[1], Number(args[2]), 8, "RYG", user)); break;
      case "/dp-trt": sendTelegram(chatId, withdraw(args[1], Number(args[2]), 9, "TRT", user)); break;
      
      default: 
        // ถ้าขึ้นต้นด้วย / แต่ไม่ตรงกับเคสไหนเลย ให้บอกว่าไม่รู้จัก
        if (cmd.startsWith("/")) {
          sendTelegram(chatId, "❌ ไม่รู้จักคำสั่งนี้ครับ\nลองพิมพ์ หรือเลือกจาก /menu แทนนะ");
        }
    }
  } catch (err) {
    console.error("Error: " + err);
  }
}

// ============================================================
// 3. UI FUNCTIONS (แผงควบคุม & การโต้ตอบ)
// ============================================================

function sendControlPanel(chatId) {
  const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`;
  const payload = {
    "chat_id": chatId ? chatId.toString() : CHAT_ID,
    "text": "🎮 <b>Warehouse Management System</b>\nเลือกรายการที่ต้องการดำเนินการด้านล่าง:",
    "parse_mode": "HTML",
    "reply_markup": JSON.stringify({
      "inline_keyboard": [
        [
          { "text": "📦 ดูสต็อกทั้งหมด", "callback_data": "/allstock" }, 
          { "text": "⚠️ ของใกล้หมด", "callback_data": "/lowstock" }
        ],
        [
          { "text": "📊 สรุปรายงาน", "callback_data": "/report" }, 
          { "text": "📜 ประวัติล่าสุด", "callback_data": "/history" }
        ],
        [
          { "text": "📤 วิธีเบิก/เติม", "callback_data": "/help" },
          { "text": "📚 คู่มือใช้งาน", "callback_data": "/help" }
        ]
      ]
    })
  };
  
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  
  UrlFetchApp.fetch(url, options);
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
  let logSheet = ss.getSheetByName("LOGS");
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

function allStock() {
  const data = getSheet().getDataRange().getValues();
  let msg = "📦 <b>รายการทั้งหมด:</b>\n\n";
  for (let i = 1; i < data.length; i++) {
    msg += `<code>${data[i][0]}</code> | ${data[i][4]} | ${data[i][2].substring(0,15)}..\n`;
  }
  return msg;
}

function lowStock() {
  const data = getSheet().getDataRange().getValues();
  let msg = "⚠️ <b>ของใกล้หมด:</b>\n\n";
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][4] <= LOW_STOCK_LIMIT) {
      msg += `🚨 ${data[i][0]} | เหลือ: ${data[i][4]}\n`;
      count++;
    }
  }
  return count > 0 ? msg : "✅ สต็อกปกติ";
}

function report() {
  const data = getSheet().getDataRange().getValues();
  let totalItems = data.length - 1;
  let lowCount = data.filter((row, idx) => idx > 0 && row[4] <= LOW_STOCK_LIMIT).length;
  return `📊 <b>STOCK REPORT</b>\n━━━━━━━━━━━━━━━\n📦 ทั้งหมด: ${totalItems}\n⚠️ ใกล้หมด: ${lowCount}\n📅 ${new Date().toLocaleString('th-TH')}`;
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
  let msg = `🔔 <b>${title}</b>\n━━━━━━━━━━━━━━━\n🔖 Code: ${code}\n🔢 จำนวน: ${type === "IN" ? "➕" : "➖"} ${qty}\n👤 โดย: ${user}\n${dept ? "🏢 แผนก: " + dept + "\n" : ""}📊 เหลือ: <b>${remain}</b>\n━━━━━━━━━━━━━━━`;
  sendTelegram(CHAT_ID, msg);
}

function helpMenu() {
  return `📦 <b>Supply Stock Management</b>
📚 <b>คู่มือการใช้งาน STOCK BOT</b>
━━━━━━━━━━━━━━━
<b>🔍 หมวดหมู่: ตรวจสอบและค้นหา</b>
• <code>/stock [รหัส]</code> - ดูข้อมูลสินค้า
• <code>/search [ชื่อ]</code> - ค้นหาสินค้า
• <code>/allstock</code> - ดูรายการทั้งหมด
• <code>/lowstock</code> - ดูของที่ใกล้หมด
• <code>/history</code> - ดูประวัติล่าสุด
• <code>/report</code> - ดูสรุปภาพรวม

<b>📥 หมวดหมู่: การนำเข้า (Restock)</b>
• <code>/restock [รหัส] [จำนวน]</code> 

<b>📤 หมวดหมู่: การเบิกจ่าย (แผนก)</b>
• <code>/dp-cbr [รหัส] [จำนวน]</code>
• <code>/dp-ccs [รหัส] [จำนวน]</code>
• <code>/dp-sko [รหัส] [จำนวน]</code>
• <code>/dp-ryg [รหัส] [จำนวน]</code>
• <code>/dp-trt [รหัส] [จำนวน]</code>

<b>🎮 อื่นๆ</b>
• <code>/menu</code> - เปิดแผงควบคุมหลัก`;
}

function history(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("LOGS");
  if(!sheet) return "❌ ไม่พบชีต LOGS";
  const data = sheet.getDataRange().getValues();
  let msg = "📜 <b>Stock History (10 ล่าสุด)</b>\n\n";
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

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("STOCK");
  const data = sheet.getRange(row,1,1,sheet.getLastColumn()).getValues()[0];

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

    sendTelegram(CHAT_ID,msg);

    // บันทึกว่าแจ้งเตือนแล้ว
    sheet.getRange(row,11).setValue("ALERTED");

  }

  // ถ้าเติมสินค้าแล้ว reset alert
  if (qty > LOW_STOCK_LIMIT && alertStatus === "ALERTED") {
    sheet.getRange(row,11).setValue("");
  }

}

function setWebhook() {
  const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/setWebhook?url=${WEB_APP_URL}`;
  const response = UrlFetchApp.fetch(url);
  Logger.log("Webhook Set: " + response.getContentText());
}

function exportStockPDF() {
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

  sendPDFToTelegram(pdfBlob);
}

function sendPDFToTelegram(pdfBlob) {

  const url = "https://api.telegram.org/bot" + TELEGRAM_TOKEN + "/sendDocument";

  const payload = {
    chat_id: CHAT_ID,
    document: pdfBlob,
    caption: "📦 Stock Report PDF"
  };

  UrlFetchApp.fetch(url, {
    method: "post",
    payload: payload,
    muteHttpExceptions: true
  });
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
