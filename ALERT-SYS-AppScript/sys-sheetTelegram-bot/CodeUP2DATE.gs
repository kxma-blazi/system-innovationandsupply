// ============================================================
// 1. CONFIGURATION API https://aistudio.google.com/api-keys?project=gen-lang-client-0093320233
// ============================================================
const TELEGRAM_TOKEN = "Eight327163778:AAFM4aKpxT29WTB4z_StzKEEcRGrSBDS2_s";
const CHAT_ID = "-0NE003731290917"; 
const LOW_STOCK_LIMIT = 1;
const GEMINI_API_KEY = "AIzaSyDzLqjhgUeIzVO-ZoCcGFr_a7lhoq6JDYg";

// ============================================================
// 2. CORE SYSTEM (doPost)
// ============================================================

function doPost(e) {
  try {
    if (!e || !e.postData) return;
    const data = JSON.parse(e.postData.contents);

    if (data.callback_query) {
      handleCallback(data.callback_query);
      return;
    }

    if (!data.message || !data.message.text) return;
    const chatId = data.message.chat.id;
    const text = data.message.text.trim();
    const user = data.message.from.first_name || "Unknown";
    const args = text.split(/\s+/);
    const command = args[0].toLowerCase();

    switch (command) {
      case "/start":
      case "/menu": sendControlPanel(chatId); break;
      case "/stock":
        if (args.length < 2) return sendTelegram(chatId, "⚠️ ระบุรหัส: `/stock [ID] [ชื่อชีต]`");
        sendTelegram(chatId, getStockInfo(args[1], args[2] || "STOCK"));
        break;
      case "/search":
        if (args.length < 2) return sendTelegram(chatId, "⚠️ ระบุคำค้น: `/search [คำ]`");
        sendTelegram(chatId, searchProduct(args[1], args[2] || "STOCK"));
        break;
      case "/allstock":
        sendAllStock(chatId, parseInt(args[1]) || 1, args[2] || "STOCK");
        break;
      case "/lowstock":
        sendTelegram(chatId, getLowStock(args[1] || "STOCK"));
        break;
      case "/report":
        sendTelegram(chatId, getReportSummary(args[1] || "STOCK"));
        break;
      case "/history":
        sendTelegram(chatId, getHistory());
        break;
      case "/restock":
        if (args.length < 3) return sendTelegram(chatId, "⚠️ รูปแบบ: `/restock [ID] [จำนวน] [ชื่อชีต]`");
        sendTelegram(chatId, restock(args[1], Number(args[2]), user, args[3] || "STOCK"));
        break;
      case "/exportpdf": 
        sendDailyStockPDF(chatId); break;
      
      // แผนกเบิกจ่าย
      case "/dp-cbr": withdrawCmd(chatId, args, 5, "CBR", user); break;
      case "/dp-ccs": withdrawCmd(chatId, args, 6, "CCS", user); break;
      case "/dp-sko": withdrawCmd(chatId, args, 7, "SKO", user); break;
      case "/dp-ryg": withdrawCmd(chatId, args, 8, "RYG", user); break;
      case "/dp-trt": withdrawCmd(chatId, args, 9, "TRT", user); break;

      default:
        // ถ้าไม่ใช่คำสั่งระบบ ให้ AI ตอบโดยใช้บริบทจากสต็อกจริง
        sendTelegram(chatId, callGemini(text));
    }
  } catch (err) {
    Logger.log("Error: " + err.toString());
  }
}

function withdrawCmd(chatId, args, col, dept, user) {
  if (args.length < 3) return sendTelegram(chatId, `⚠️ รูปแบบ: /dp-${dept.toLowerCase()} [ID] [จำนวน] [ชื่อชีต]`);
  sendTelegram(chatId, withdraw(args[1], Number(args[2]), col, dept, user, args[3] || "STOCK"));
}

// ============================================================
// 3. STOCK LOGIC
// ============================================================

function withdraw(code, qty, colIndex, deptName, user, targetSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) return "❌ ไม่พบชีต: " + targetSheetName;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === code.toLowerCase()) {
      let currentStock = Number(data[i][4]);
      if (currentStock < qty) return `❌ สต็อกไม่พอใน ${targetSheetName} (เหลือ ${currentStock})`;

      let newStock = currentStock - qty;
      let newDeptTotal = (Number(data[i][colIndex]) || 0) + qty;

      sheet.getRange(i + 1, 5).setValue(newStock); 
      sheet.getRange(i + 1, colIndex + 1).setValue(newDeptTotal);

      writeLog(user, "Withdraw", data[i][0], qty, `${targetSheetName} (${deptName})`, newStock);
      
      let msg = `✅ *เบิกสำเร็จ*\n📦 ${data[i][2]}\n📂 ชีต: ${targetSheetName}\n📍 แผนก: ${deptName}\n📉 จำนวน: -${qty}\n📊 คงเหลือ: ${newStock}`;
      if (newStock <= LOW_STOCK_LIMIT) msg += `\n\n⚠️ *ALERT: สินค้าใกล้หมด!* ⚠️`;
      
      notifyActivity(msg);
      return msg;
    }
  }
  return "❌ ไม่พบรหัสในชีต " + targetSheetName;
}

function restock(code, qty, user, targetSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) return "❌ ไม่พบชีต";

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === code.toLowerCase()) {
      let newQty = (Number(data[i][4]) || 0) + qty;
      sheet.getRange(i + 1, 5).setValue(newQty);
      writeLog(user, "Restock", data[i][0], qty, targetSheetName, newQty);

      let msg = `✅ *เติมสต็อกสำเร็จ*\n📦 ${data[i][2]}\n➕ เพิ่ม: +${qty}\n📊 ยอดใหม่: ${newQty}`;
      notifyActivity(msg);
      return msg;
    }
  }
  return "❌ ไม่พบรหัสสินค้า";
}

// ============================================================
// 4. REPORTS & LOGS
// ============================================================

function writeLog(user, action, code, qty, note, remain) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("Logs") || ss.insertSheet("Logs");
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(["วันเวลา", "ชื่อผู้ใช้", "รายการ", "รหัสสินค้า", "จำนวน", "เเผนก/ชื่อชีต", "ยอดเหลือสุทธิ"]);
  }
  logSheet.appendRow([new Date(), user, action, code, qty, note, remain]);
}

function getHistory() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs");
  if (!sheet) return "❌ ไม่พบประวัติ";
  const data = sheet.getDataRange().getValues();
  const lastLogs = data.slice(-10).reverse();
  let msg = "📜 *ประวัติ 10 รายการล่าสุด*\n\n";
  lastLogs.forEach(r => {
    if (r[0] === "วันเวลา") return;
    msg += `🕒 ${Utilities.formatDate(new Date(r[0]), "GMT+7", "dd/MM HH:mm")} | ${r[1]}\n${r[2]} [${r[3]}] x${r[4]} (${r[5]})\n---\n`;
  });
  return msg;
}

function getReportSummary(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return "❌ ไม่พบชีต";
  const data = sheet.getDataRange().getValues();
  let totalItems = data.length - 1;
  let totalQty = 0;
  let low = 0;
  data.forEach((r, i) => { if (i > 0) { totalQty += Number(r[4]) || 0; if (Number(r[4]) <= LOW_STOCK_LIMIT) low++; } });
  return `📊 *สรุปรายงาน (${sheetName})*\n\n📦 รายการทั้งหมด: ${totalItems}\n🔢 ชิ้นรวม: ${totalQty}\n⚠️ ใกล้หมด: ${low}`;
}

// ============================================================
// 5. UI & UTILITIES
// ============================================================

function handleCallback(query) {
  const chatId = query.message.chat.id;
  const data = query.data;
  if (data === "main_menu") sendControlPanel(chatId);
  else if (data.startsWith("page_")) {
    const parts = data.split("_");
    sendAllStock(chatId, parseInt(parts[1]), parts[2]);
  } else if (data === "view_logs") sendTelegram(chatId, getHistory());
  else if (data === "check_low") sendTelegram(chatId, getLowStock("STOCK"));
}

function sendControlPanel(chatId) {
  const payload = {
    chat_id: chatId,
    text: "📦 *SYSTEM INNOVATION AND SUPPLY*\nเลือกเมนูการใช้งานด้านล่าง:",
    parse_mode: "Markdown",
    reply_markup: {
      inline_keyboard: [
        [{ text: "📦 ดูสต็อกทั้งหมด", callback_data: "page_1_STOCK" }],
        [{ text: "⚠️ ของใกล้หมด", callback_data: "check_low" }],
        [{ text: "📜 ประวัติล่าสุด", callback_data: "view_logs" }],
        [{ text: "🏠 กลับเมนูหลัก", callback_data: "main_menu" }]
      ]
    }
  };
  sendToTelegram("sendMessage", payload);
}

function sendAllStock(chatId, page, sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return;
  const rows = sheet.getDataRange().getValues().slice(1);
  const pageSize = 10;
  const totalPages = Math.ceil(rows.length / pageSize);
  const currentRows = rows.slice((page - 1) * pageSize, page * pageSize);

  let msg = `📦 *รายการสต็อก (${sheetName})* หน้า ${page}/${totalPages}\n\n`;
  currentRows.forEach(r => msg += `🔹 \`${r[0]}\`: ${r[2].toString().substring(0,15)}... | *${r[4]}*\n`);

  const nav = [];
  if (page > 1) nav.push({ text: "⬅️", callback_data: `page_${page-1}_${sheetName}` });
  if (page < totalPages) nav.push({ text: "➡️", callback_data: `page_${page+1}_${sheetName}` });

  sendToTelegram("sendMessage", {
    chat_id: chatId, text: msg, parse_mode: "Markdown",
    reply_markup: { inline_keyboard: [nav, [{ text: "🏠 เมนูหลัก", callback_data: "main_menu" }]] }
  });
}

function searchProduct(keyword, sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  let res = data.filter(r => r[2].toString().toLowerCase().includes(keyword.toLowerCase()) || r[0].toString().toLowerCase() === keyword.toLowerCase());
  return res.length > 0 ? `🔍 ผลการค้นหา "${keyword}":\n\n` + res.map(r => `✅ \`${r[0]}\` - ${r[2]}\nคงเหลือ: *${r[4]}*`).join("\n\n") : "❌ ไม่พบสินค้า";
}

function getLowStock(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  let list = data.filter((r, i) => i > 0 && Number(r[4]) <= LOW_STOCK_LIMIT).map(r => `⚠️ \`${r[0]}\`: ${r[2]} (${r[4]})`);
  return list.length > 0 ? `⚠️ *สินค้าใกล้หมด (${sheetName})*\n\n` + list.join("\n") : "✅ ไม่มีของใกล้หมด";
}

function getStockInfo(code, sheetName) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === code.toLowerCase()) 
      return `📦 *ข้อมูลสินค้า (${sheetName})*\nID: ${data[i][0]}\nชื่อ: ${data[i][2]}\nคงเหลือ: ${data[i][4]}`;
  }
  return "❌ ไม่พบสินค้า";
}

// ============================================================
// 6. TELEGRAM & PDF & NOTIFY
// ============================================================

function sendTelegram(chatId, text) {
  sendToTelegram("sendMessage", { chat_id: chatId, text: text, parse_mode: "Markdown" });
}

function notifyActivity(msg) {
  sendToTelegram("sendMessage", { chat_id: CHAT_ID, text: "📢 *Activity Report*\n" + msg, parse_mode: "Markdown" });
}

function sendToTelegram(method, payload) {
  const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/${method}`;
  return UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify(payload) });
}

function sendDailyStockPDF(chatId = CHAT_ID) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const blob = DriveApp.getFileById(ss.getId()).getAs('application/pdf').setName("Stock_Report_" + Utilities.formatDate(new Date(), "GMT+7", "dd-MM-yyyy") + ".pdf");
    const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendDocument`;
    const payload = { chat_id: chatId, document: blob, caption: "📄 รายงานสต็อกสินค้าประจำวัน (PDF)" };
    UrlFetchApp.fetch(url, { method: "post", payload: payload });
  } catch (e) { sendTelegram(chatId, "❌ เกิดข้อผิดพลาดในการสร้าง PDF"); }
}

// ============================================================
// 7. TRIGGERS (Auto-Notify)
// ============================================================

function onEdit(e) {
  const range = e.range;
  const sheetName = range.getSheet().getName();
  if (sheetName === "Logs") return;
  const oldValue = e.oldValue || "ว่างเปล่า";
  const newValue = e.value || "ว่างเปล่า";
  const user = Session.getActiveUser().getEmail() || "ผู้ใช้นอกระบบ";
  const msg = `⚠️ *แก้ไขไฟล์โดยตรง*\n📍 ชีต: ${sheetName}\n🎯 เซลล์: ${range.getA1Notation()}\n🔄 ${oldValue} ➡️ ${newValue}\n👤 โดย: ${user}`;
  notifyActivity(msg);
}

function sendDailyReport() {
  const summary = getReportSummary("STOCK");
  const lowStock = getLowStock("STOCK");
  notifyActivity(`☀️ *สรุปรายงานประจำวัน*\n\n${summary}\n\n${lowStock}`);
}

// ============================================================
// 8. AI SYSTEM (Gemini) - ฉบับอ่านสต็อก Real-time
// ============================================================

function callGemini(userMessage) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("STOCK");
    let stockContext = "ข้อมูลสต็อกปัจจุบัน:\n";

    if (sheet) {
      const data = sheet.getDataRange().getValues();
      // อ่านข้อมูล 50 รายการแรกมาให้ AI เป็นบริบท
      for (let i = 1; i < Math.min(data.length, 51); i++) {
        stockContext += `- ${data[i][2]} (ID: ${data[i][0]}) คงเหลือ: ${data[i][4]}\n`;
      }
    }

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=${GEMINI_API_KEY}`;
    
    const systemPrompt = `คุณคือ SmartStock AI ของบริษัท System Innovation and Supply 
    คุณมีข้อมูลสต็อกสินค้าอยู่ด้านล่างนี้ หากคนถามถึงจำนวนสินค้า ให้เช็คจากข้อมูลที่ให้ไป 
    ตอบสั้นๆ เป็นกันเอง และลงท้ายด้วยการแนะนำให้พิมพ์ /menu หากต้องการทำรายการเบิกจ่าย`;
    
    const payload = {
      "contents": [{
        "parts": [{
          "text": `${systemPrompt}\n\n${stockContext}\n\nคำถามจากผู้ใช้: ${userMessage}`
        }]
      }]
    };

    const options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    const response = UrlFetchApp.fetch(url, options);
    const resText = response.getContentText();
    const json = JSON.parse(resText);

    if (json.candidates && json.candidates[0] && json.candidates[0].content) {
      return "🤖 " + json.candidates[0].content.parts[0].text.trim();
    } else {
      return "🤖 ขออภัยครับ ผมเข้าถึงข้อมูลไม่ได้ชั่วคราว ลองพิมพ์ /menu ดูนะครับ";
    }
  } catch (e) {
    return "🤖 ระบบ AI พักผ่อนครับ ลองใช้เมนูปกติก่อนนะ";
  }
}
