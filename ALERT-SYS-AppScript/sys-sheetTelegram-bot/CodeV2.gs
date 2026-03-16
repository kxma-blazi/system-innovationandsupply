// ============================================================
// 1. CONFIGURATION
// ============================================================
const TELEGRAM_TOKEN = "8327163778:AAFM4aKpxT29WTB4z_StzKEEcRGrSBDS2_s";
const CHAT_ID = "-1003731290917"; 
const LOW_STOCK_LIMIT = 1;
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxhX-EacuJJeQsByoEywq-Z6xAZcI_8WLgjAt4BheofPDJ3seVQGrv6C4YrWc86USVv/exec"; // UPDATE ด้วยหลังอัพโค้ด

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
      case "/start": sendTelegram(chatId, startMessage()); break;
      case "/help": sendTelegram(chatId, helpMenu()); break;

      case "/history": sendTelegram(chatId, history()); break;
      
      case "/menu": sendTelegram(chatId, menuMessage()); break;
      case "/manual": sendTelegram(chatId, helpMenu()); break;

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

return `📦 <b>Warehouse Management System</b>
Stock Control Bot

📚 <b>User Manual</b>
━━━━━━━━━━━━━━━

🔎 <b>ตรวจสอบสินค้า</b>
/stock CODE

🔍 <b>ค้นหาสินค้า</b>
/search KEYWORD

📦 <b>ดู Stock ทั้งหมด</b>
/allstock

⚠️ <b>สินค้าใกล้หมด</b>
/lowstock

📊 <b>รายงานระบบ</b>
/report

━━━━━━━━━━━━━━━
📥 <b>เติมสินค้า</b>
/restock CODE QTY

📤 <b>เบิกสินค้า</b>
/dp-cbr CODE QTY
/dp-ccs CODE QTY
/dp-sko CODE QTY
/dp-ryg CODE QTY
/dp-trt CODE QTY

━━━━━━━━━━━━━━━
🤖 STOCK BOT SYSTEM
`;

}

function setWebhook() {
  const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/setWebhook?url=${WEB_APP_URL}`;
  const response = UrlFetchApp.fetch(url);
  Logger.log("ผลการตั้งค่า: " + response.getContentText());
}

function startMessage(){

return `👋 <b>Welcome to STOCK BOT</b>

📦 ระบบจัดการคลังสินค้า
Warehouse Management System

พิมพ์ /menu เพื่อดูเมนูหลัก
พิมพ์ /help เพื่อดูคำสั่งทั้งหมด`;

}

function menuMessage(){

return `📦 <b>STOCK BOT MENU</b>

1️⃣ ตรวจสอบสินค้า
/stock CODE

2️⃣ ค้นหาสินค้า
/search KEYWORD

3️⃣ ดู Stock ทั้งหมด
/allstock

4️⃣ สินค้าใกล้หมด
/lowstock

5️⃣ รายงานระบบ
/report

6️⃣ เติมสินค้า
/restock CODE QTY

7️⃣ เบิกสินค้า
/dp-cbr CODE QTY
/dp-ccs CODE QTY
/dp-sko CODE QTY
/dp-ryg CODE QTY
/dp-trt CODE QTY

📚 /manual คู่มือใช้งาน
`;

}

function history(){

const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName("LOGS");

if(!sheet) return "❌ ไม่พบชีต LOGS";

const data = sheet.getDataRange().getValues();

let msg = "📜 <b>Stock History</b>\n\n";

let start = Math.max(1, data.length - 10);

for(let i=start; i<data.length; i++){

let date = Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "dd/MM HH:mm");

msg += `🕒 ${date}
👤 ${data[i][1]}
📦 ${data[i][2]} ${data[i][3]}
🔢 ${data[i][4]}
📊 เหลือ ${data[i][6]}

`;

}

return msg;

}

function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  
  if (sheetName !== "STOCK" || range.getRow() <= 1) return;

  const row = range.getRow();
  const col = range.getColumn();
  const numRows = range.getHeight();
  const numCols = range.getWidth();

  // ดึง Email ของผู้แก้ไข (ต้องใช้ Installable Trigger)
  // หากดึงไม่ได้จะแสดงเป็น "Internal/Manual Edit"
  const editor = Session.getActiveUser().getEmail() || "Authorized User";

  if (numRows > 1 || numCols > 1) {
    let bulkMsg = `📦 <b>Bulk Update Notification</b>\n` +
                 `━━━━━━━━━━━━━━━\n` +
                 `📍 <b>ช่วงที่แก้ไข:</b> แถวที่ ${row} ถึง ${row + numRows - 1}\n` +
                 `📊 <b>จำนวนที่เปลี่ยน:</b> ${numRows * numCols} ช่อง\n` +
                 `━━━━━━━━━━━━━━━\n` +
                 `👤 <b>โดย:</b> ${editor}`;

    sendTelegram(CHAT_ID, bulkMsg);
    return;
  }

  const newValue = e.value !== undefined ? e.value : "ถูกลบ/ว่างเปล่า";
  const oldValue = e.oldValue !== undefined ? e.oldValue : "ไม่มีข้อมูลเดิม";
  
  const materialCode = sheet.getRange(row, 2).getValue(); 
  const description = sheet.getRange(row, 3).getValue();  
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colName = headers[col - 1] || "ไม่ทราบคอลัมน์";

  let msg = `📝 <b>มีการแก้ไขข้อมูลในชีต</b>\n` +
            `━━━━━━━━━━━━━━━\n` +
            `📍 <b>ตำแหน่ง:</b> แถวที่ ${row} [${colName}]\n` +
            `🔖 <b>รหัส:</b> ${materialCode}\n` +
            `📦 <b>สินค้า:</b> ${description.substring(0, 50)}${description.length > 50 ? "..." : ""}\n` +
            `━━━━━━━━━━━━━━━\n` +
            `❌ <b>ค่าเดิม:</b> ${oldValue}\n` +
            `✅ <b>เปลี่ยนเป็น:</b> <b>${newValue}</b>\n` +
            `━━━━━━━━━━━━━━━\n` +
            `👤 <b>ผู้เเก้ไข:</b> ${editor}`;

  sendTelegram(CHAT_ID, msg);
}

function testUser() {
  const email = Session.getActiveUser().getEmail();
  Logger.log("Email ของคุณคือ: " + email);
  if (!email) {
    Logger.log("⚠️ ไม่สามารถดึง Email ได้ (อาจเพราะยังไม่ได้กด Allow สิทธิ์)");
  }
}

// ฟังก์ชันส่งรายงานสรุปประจำวัน
function sendDailySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName("STOCK");
  const logSheet = ss.getSheetByName("LOGS");
  
  // 1. สรุปยอดสินค้าใกล้หมด
  const stockData = stockSheet.getDataRange().getValues();
  let lowStockList = "";
  let lowCount = 0;
  for (let i = 1; i < stockData.length; i++) {
    if (stockData[i][4] <= LOW_STOCK_LIMIT) {
      lowStockList += `🚨 ${stockData[i][0]} | เหลือ: ${stockData[i][4]}\n`;
      lowCount++;
    }
  }

  // 2. สรุปกิจกรรมของวันนี้ (นับจาก LOGS)
  const logData = logSheet.getDataRange().getValues();
  const today = new Date().toLocaleDateString();
  let dailyIn = 0;
  let dailyOut = 0;
  
  for (let i = 1; i < logData.length; i++) {
    let logDate = new Date(logData[i][0]).toLocaleDateString();
    if (logDate === today) {
      if (logData[i][2] === "Restock") dailyIn += Number(logData[i][4]);
      if (logData[i][2] === "Withdraw") dailyOut += Number(logData[i][4]);
    }
  }

  // 3. สร้างข้อความ
  let msg = `📊 <b>Daily Stock Summary</b>\n`;
  msg += `📅 วันที่: ${new Date().toLocaleDateString('th-TH')}\n`;
  msg += `━━━━━━━━━━━━━━━\n`;
  msg += `📥 เติมของวันนี้: +${dailyIn} รายการ\n`;
  msg += `📤 เบิกของวันนี้: -${dailyOut} รายการ\n`;
  msg += `━━━━━━━━━━━━━━━\n`;
  msg += `⚠️ <b>สินค้าใกล้หมด (${lowCount} รายการ):</b>\n`;
  msg += lowCount > 0 ? lowStockList : "✅ สต็อกปกติทุกรายการ\n";
  msg += `━━━━━━━━━━━━━━━\n`;
  msg += `🤖 <i>Auto-report by Stock System</i>`;

  // ส่งไปที่ Chat ID กลุ่ม
  sendTelegram(CHAT_ID, msg);
}

// ฟังก์ชันสำหรับส่งรายงานสรุปประจำวันอัตโนมัติ
function sendDailySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName("STOCK");
  const logSheet = ss.getSheetByName("LOGS");
  
  if (!stockSheet || !logSheet) return; // กัน Error ถ้าหาชีตไม่เจอ

  const stockData = stockSheet.getDataRange().getValues();
  const logData = logSheet.getDataRange().getValues();
  const today = new Date().toLocaleDateString('en-CA'); // รูปแบบ YYYY-MM-DD สำหรับเช็คเงื่อนไข
  
  let lowStockList = "";
  let lowCount = 0;
  let dailyIn = 0;
  let dailyOut = 0;

  // 1. ตรวจสอบสินค้าใกล้หมด
  for (let i = 1; i < stockData.length; i++) {
    let currentQty = Number(stockData[i][4]);
    if (currentQty <= LOW_STOCK_LIMIT) {
      lowStockList += `🚨 <code>${stockData[i][0]}</code> | เหลือ: <b>${currentQty}</b>\n   └ ${stockData[i][2].substring(0, 30)}\n`;
      lowCount++;
    }
  }

  // 2. สรุปกิจกรรมของวันนี้จากหน้า LOGS
  for (let i = 1; i < logData.length; i++) {
    let logDate = new Date(logData[i][0]).toLocaleDateString('en-CA');
    if (logDate === today) {
      if (logData[i][2] === "Restock") dailyIn += Number(logData[i][4]);
      if (logData[i][2] === "Withdraw") dailyOut += Number(logData[i][4]);
    }
  }

  // 3. สร้างข้อความรายงานแบบใหม่
  let msg = `📊 <b>REPORT: สรุปสถานะคลังสินค้า</b>\n`;
  msg += `📅 ประจำวันที่: ${new Date().toLocaleDateString('th-TH', { 
    year: 'numeric', month: 'long', day: 'numeric' 
  })}\n`;
  msg += `📍 <i>อัปเดตล่าสุด: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm")} น.</i>\n`;
  msg += `━━━━━━━━━━━━━━━\n\n`;

  msg += `📈 <b>ภาพรวมกิจกรรมวันนี้</b>\n`;
  msg += `📥 เติมสินค้าเพิ่ม:  <code>${dailyIn.toLocaleString()}</code> รายการ\n`;
  msg += `📤 เบิกไปใช้งาน:  <code>${dailyOut.toLocaleString()}</code> รายการ\n`;
  msg += `🔄 รวมความเคลื่อนไหว: <b>${(dailyIn + dailyOut).toLocaleString()}</b> รายการ\n\n`;

  msg += `⚠️ <b>สินค้าใกล้หมด (${lowCount})</b>\n`;
  msg += `━━━━━━━━━━━━━━━\n`;
  if (lowCount > 0) {
    msg += lowStockList;
  } else {
    msg += `✅ สินค้าทุกรายการมีจำนวนเพียงพอ\n`;
  }
  msg += `━━━━━━━━━━━━━━━\n`;
  msg += `🤖 <b>STOCK BOT AUTOMATION</b>`;

  // ส่งข้อความเข้ากลุ่ม Telegram
  sendTelegram(CHAT_ID, msg);
}

/*
 Backup
// ฟังก์ชันนี้จะทำงานอัตโนมัติเมื่อมีการแก้ไขเซลล์ใน Google Sheets
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  
  // ตรวจสอบว่าแก้ในชีต STOCK และไม่ใช่แถวหัวข้อ
  if (sheetName === "STOCK" && range.getRow() > 1) {
    
    const row = range.getRow();
    const col = range.getColumn();
    
    // ดึงค่าใหม่ และ ค่าเก่า
    const newValue = e.value || "ถูกลบ/ว่างเปล่า";
    const oldValue = e.oldValue || "ไม่มีข้อมูลเดิม";
    
    // ดึงชื่อสินค้า (คอลัมน์ C) และรหัส (คอลัมน์ B) มาแสดงเพื่อให้รู้ว่าแก้ที่รายการไหน
    const materialCode = sheet.getRange(row, 2).getValue(); // คอลัมน์ B (Material Code)
    const description = sheet.getRange(row, 3).getValue();  // คอลัมน์ C (Description)
    
    // แปลงเลขคอลัมน์เป็นชื่อหัวข้อ (Header)
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colName = headers[col - 1] || "ไม่ทราบคอลัมน์";

    // สร้างข้อความแจ้งเตือน
    let msg = `📝 <b>มีการแก้ไขข้อมูลในชีต</b>\n`;
    msg += `━━━━━━━━━━━━━━━\n`;
    msg += `📍 <b>ตำแหน่ง:</b> แถวที่ ${row} [${colName}]\n`;
    msg += `🔖 <b>รหัส:</b> ${materialCode}\n`;
    msg += `📦 <b>สินค้า:</b> ${description}\n`;
    msg += `━━━━━━━━━━━━━━━\n`;
    msg += `❌ <b>ค่าเดิม:</b> ${oldValue}\n`;
    msg += `✅ <b>เปลี่ยนเป็น:</b> <b>${newValue}</b>\n`;
    msg += `━━━━━━━━━━━━━━━\n`;
    msg += `👤 <b>โดย:</b> Admin (Manual Edit)`;

    // เรียกฟังก์ชันส่ง Telegram (ต้องมีฟังก์ชัน sendTelegram ในสคริปต์ด้วย)
    sendTelegram(CHAT_ID, msg);
  }
}
*/


function sendControlPanel(chatId) {
  const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`;
  const payload = {
    chat_id: chatId,
    text: "🎮 <b>แผงควบคุมสต็อก</b>\nเลือกรายการที่ต้องการตรวจสอบ:",
    parse_mode: "HTML",
    reply_markup: {
      inline_keyboard: [
        [{ text: "📦 ดูสต็อกทั้งหมด", callback_data: "/allstock" }, { text: "⚠️ ของใกล้หมด", callback_data: "/lowstock" }],
        [{ text: "📊 สรุปรายงาน", callback_data: "/report" }, { text: "📜 ประวัติล่าสุด", callback_data: "/history" }]
      ]
    }
  };
  UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify(payload) });
}
