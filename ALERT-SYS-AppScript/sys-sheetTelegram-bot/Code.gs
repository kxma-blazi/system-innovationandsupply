// ============================================================
// 1. CONFIGURATION
// ============================================================
const TELEGRAM_TOKEN = "8327163778:AAFM4aKpxT29WTB4z_StzKEEcRGrSBDS2_s";
const CHAT_ID = "-1003731290917"; 
const LOW_STOCK_LIMIT = 1;
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyOUBt4Fze-BIMChrOB_tBVm4MoFcsOzmTAGf3YO8GKgTD2joIlr7CXwX6AS-LAaoAs/exec";

// ============================================================
// 2. CORE SYSTEM (แก้ไข doPost ให้ส่งชื่อผู้ใช้งาน)
// ============================================================

function doPost(e) {
  try {
    const contents = e.postData.contents;
    const data = JSON.parse(contents);

    if (!data.message || !data.message.text) return;

    const chatId = data.message.chat.id;
    const text = data.message.text;
    
    // ดึงชื่อจาก Telegram (ถ้าไม่มีให้ใช้คำว่า Unknown)
    const user = data.message.from.first_name || data.message.from.username || "Unknown";
    
    const args = text.split(" ");
    const cmd = args[0].toLowerCase();

    switch(cmd) {
      case "/help": sendTelegram(chatId, helpMenu()); break;
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
        if (cmd.startsWith("/")) sendTelegram(chatId, "❌ ไม่รู้จักคำสั่ง\nTry /help");
    }
  } catch (err) {
    console.error("Error: " + err);
  }
}

// ============================================================
// 3. STOCK LOGIC (แก้ไขให้บันทึก Logs)
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

// ============================================================
// 4. บันทึกประวัติ (Write Logs)
// ============================================================

function writeLog(user, action, code, qty, dept, remain) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("LOGS");
  if (!logSheet) return; // กัน Error ถ้าหาชีตไม่เจอ
  
  logSheet.appendRow([
    new Date(), 
    user,       
    action,     
    code,       
    qty,        
    dept,       
    remain      
  ]);
}

// ============================================================
// 5. รายงานและแจ้งเตือน
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
  return `📦 <b>STOCK BOT MENU</b>\n━━━━━━━━━━━━━━━\n🔎 /stock [Code]\n🔍 /search [Text]\n📦 /allstock\n⚠️ /lowstock\n📊 /report\n\n📥 <b>เติม:</b> /restock [Code] [Qty]\n📤 <b>เบิก:</b> /dp-[dept] [Code] [Qty]\n(cbr, ccs, sko, ryg, trt)`;
}

function setWebhook() {
  const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/setWebhook?url=${WEB_APP_URL}`;
  const response = UrlFetchApp.fetch(url);
  Logger.log("ผลการตั้งค่า: " + response.getContentText());
}