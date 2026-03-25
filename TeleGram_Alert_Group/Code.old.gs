// ============================================================
// 1. CONFIGURATION
// ============================================================
const TELEGRAM_TOKEN  = "8327163778:AAFM4aKpxT29WTB4z_StzKEEcRGrSBDS2_s";
const CHAT_ID         = "-1003731290917";
const LOW_STOCK_LIMIT = 3;   // ← ปรับขีดเตือนได้ที่นี่
const GEMINI_API_KEY  = "AIzaSyDzLqjhgUeIzVO-ZoCcGFr_a7lhoq6JDYg";
const SHEET_ID        = "1MyAWKuCtmBclqVWALUEhEZjF7jKLMHs2sqL5oXjhA0w";

// ============================================================
// 2. WEB APP ENTRY POINT (รวมเป็นอันเดียว)
// ============================================================
function doGet(e) {
  var id = (e && e.parameter && e.parameter.id) ? e.parameter.id : "";
  var tmp = HtmlService.createTemplateFromFile("index");
  tmp.targetId = id; 
  return tmp.evaluate()
    .setTitle("📦 SYSTEM INNOVATION Inventory Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// 3. DO POST
// ============================================================
function doPost(e) {
  try {
    if (!e || !e.postData) return;
    const update = JSON.parse(e.postData.contents);
    if (update.callback_query) { handleCallback(update.callback_query); return; }
    if (!update.message) return;

    const msg     = update.message;
    const chatId  = msg.chat.id;
    const user    = msg.from.first_name || "Unknown";

    if (!msg.text) return;
    const text    = msg.text.trim();
    const args    = text.split(/\s+/);
    const command = args[0].toLowerCase();

    switch (command) {
      case "/start":
      case "/menu":   sendMainMenu(chatId, user); break;

      // ── ค้นหา ──
      case "/stock":
        if (args.length < 2) return sendMsg(chatId, "⚠️ รูปแบบ: `/stock [ID]`");
        sendMsg(chatId, getStockInfo(args[1], args[2] || "STOCK")); break;
      case "/search":
        if (args.length < 2) return sendMsg(chatId, "⚠️ รูปแบบ: `/search [คำค้น]`");
        sendSearchResult(chatId, args.slice(1).join(" ")); break;

      // ── สต็อก ──
      case "/allstock":
        sendAllStock(chatId, 1, args[1] || "STOCK"); break;
      case "/lowstock":
        sendMsg(chatId, getLowStock(args[1] || "STOCK")); break;

      // ── เบิก/เติม ──
      case "/withdraw":
      case "/w":
        if (args.length < 3) return sendMsg(chatId, "⚠️ รูปแบบ: `/w [ID] [จำนวน] [แผนก]`\nแผนก: cbr, ccs, sko, ryg, trt");
        handleWithdrawText(chatId, args, user); break;
      case "/restock":
      case "/r":
        if (args.length < 3) return sendMsg(chatId, "⚠️ รูปแบบ: `/r [ID] [จำนวน]`");
        sendMsg(chatId, restock(args[1], Number(args[2]), user, args[3] || "STOCK")); break;

      // ── แผนกเบิก ──
      case "/dp-cbr": handleDeptWithdraw(chatId, args, 6, "CBR", user); break;
      case "/dp-ccs": handleDeptWithdraw(chatId, args, 7, "CCS", user); break;
      case "/dp-sko": handleDeptWithdraw(chatId, args, 8, "SKO", user); break;
      case "/dp-ryg": handleDeptWithdraw(chatId, args, 9, "RYG", user); break;
      case "/dp-trt": handleDeptWithdraw(chatId, args, 10, "TRT", user); break; // สมมติว่า TRT อยู่ต่อกัน

      // ── รายงาน ──
      case "/report":   sendMsg(chatId, getReportSummary("STOCK")); break;
      case "/history":  sendMsg(chatId, getHistory(20)); break;
      case "/daily":    sendDailySummary(chatId); break;
      case "/weekly":   sendWeeklySummary(chatId); break;
      case "/exportpdf": sendDailyStockPDF(chatId); break;

      default:
        sendMsg(chatId, callGemini(text));
    }
  } catch (err) {
    Logger.log("doPost Error: " + err);
  }
}

// ============================================================
// 4. MAIN MENU
// ============================================================
function sendMainMenu(chatId, user) {
  const name = user || "คุณ";
  sendToTelegram("sendMessage", {
    chat_id:    chatId,
    parse_mode: "Markdown",
    text: `👋 สวัสดีครับ *${name}*!\n\n📦 *SYSTEM INNOVATION AND SUPPLY*\n━━━━━━━━━━━━━━━━━━\n🤖 ระบบจัดการสต็อกอัจฉริยะ\nเลือกเมนูด้านล่างได้เลยครับ`,
    reply_markup: {
      inline_keyboard: [
        [
          { text: "📦 ดูสต็อก",        callback_data: "page_1_STOCK" },
          { text: "⚠️ ของใกล้หมด",    callback_data: "check_low"    }
        ],
        [
          { text: "🔍 ค้นหารายการ",    callback_data: "search_prompt" },
          { text: "📜 ประวัติ",        callback_data: "history_20"   }
        ],
        [
          { text: "📊 รายงานวันนี้",    callback_data: "daily_report" },
          { text: "📈 รายงานสัปดาห์",  callback_data: "weekly_report"}
        ],
        [
          { text: "➕ เติมสต็อก",       callback_data: "restock_prompt" },
          { text: "➖ เบิกรายการ",      callback_data: "withdraw_prompt"}
        ],
        [
          { text: "📄 Export PDF",      callback_data: "export_pdf"   }
        ]
      ]
    }
  });
}

// ============================================================
// 5. CALLBACK HANDLER
// ============================================================
function handleCallback(query) {
  const chatId  = query.message.chat.id;
  const data    = query.data;
  const user    = query.from.first_name || "Unknown";

  answerCallback(query.id);

  if (data === "main_menu")         { sendMainMenu(chatId, user); return; }
  if (data === "check_low")         { sendMsg(chatId, getLowStock("STOCK")); return; }
  if (data === "history_20")        { sendMsg(chatId, getHistory(20)); return; }
  if (data === "daily_report")      { sendDailySummary(chatId); return; }
  if (data === "weekly_report")     { sendWeeklySummary(chatId); return; }
  if (data === "export_pdf")        { sendDailyStockPDF(chatId); return; }
  if (data === "search_prompt")     { sendMsg(chatId, "🔍 พิมพ์ `/search [ชื่อรายการ หรือ ID หรือ แผนก]` เพื่อค้นหาครับ\n\n*ตัวอย่าง:*\n`/search CPRI`\n`/search Q001`\n`/search CBR`"); return; }
  if (data === "restock_prompt")    { sendMsg(chatId, "➕ *เติมสต็อก*\nพิมพ์: `/r [ID] [จำนวน]`\n\n*ตัวอย่าง:*\n`/r Q001 10`"); return; }
  if (data === "withdraw_prompt")   { sendMsg(chatId, "➖ *เบิกรายการ*\nพิมพ์: `/w [ID] [จำนวน] [แผนก]`\nแผนก: cbr, ccs, sko, ryg, trt\n\n*ตัวอย่าง:*\n`/w Q001 2 cbr`"); return; }

  if (data.startsWith("page_")) {
    const p = data.split("_");
    sendAllStock(chatId, parseInt(p[1]), p[2] || "STOCK"); return;
  }

  if (data.startsWith("wd_")) {
    const parts = data.split("_");
    const code = parts[1], dept = parts[2].toUpperCase();
    const colMap = { CBR: 6, CCS: 7, SKO: 8, RYG: 9, TRT: 10 };
    const col = colMap[dept] || 6;
    const result = withdraw(code, 1, col, dept, user, "STOCK");
    sendMsg(chatId, result); return;
  }

  if (data.startsWith("info_")) {
    const code = data.split("_")[1];
    sendStockDetail(chatId, code); return;
  }
}

function answerCallback(callbackId) {
  try {
    UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/answerCallbackQuery`, {
      method: "post", contentType: "application/json",
      payload: JSON.stringify({ callback_query_id: callbackId })
    });
  } catch(e) {}
}

// ============================================================
// 6. STOCK LOGIC
// ============================================================
function withdraw(code, qty, colIndex, deptName, user, targetSheetName) {
  const sheet = getSpreadsheet().getSheetByName(targetSheetName);
  if (!sheet) return "❌ ไม่พบชีต: " + targetSheetName;
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === code.toLowerCase()) {
      
      const cur = Number(data[i][5]); 
      const itemName = data[i][3];   
      
      if (cur < qty) return `❌ *สต็อกไม่พอ*\nรายการ: ${itemName}\nคงเหลือ: ${cur} ชิ้น`;
      
      const newQty   = cur - qty;
      const newDept  = (Number(data[i][colIndex]) || 0) + qty;
      
      sheet.getRange(i + 1, 6).setValue(newQty); 
      sheet.getRange(i + 1, colIndex + 1).setValue(newDept);
      
      writeLog(user, "Withdraw", data[i][0], qty, `${targetSheetName} (${deptName})`, newQty);
      
      let msg = `✅ *เบิกสำเร็จ!*\n━━━━━━━━━━━━━━━\n📦 ${itemName}\n🆔 ID: \`${data[i][0]}\`\n📍 แผนก: ${deptName}\n📉 เบิก: *-${qty}* ชิ้น\n📊 คงเหลือ: *${newQty}* ชิ้น`;
      
      if (newQty <= LOW_STOCK_LIMIT) msg += `\n\n⚠️ *รายการใกล้หมด!* เหลือแค่ ${newQty} ชิ้น`;
      notifyActivity(msg);
      return msg;
    }
  }
  return `❌ ไม่พบรหัส \`${code}\` ในชีต ${targetSheetName}`;
}

function restock(code, qty, user, targetSheetName) {
  const sheet = getSpreadsheet().getSheetByName(targetSheetName);
  if (!sheet) return "❌ ไม่พบชีต";
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === code.toLowerCase()) {
      const newQty = (Number(data[i][5]) || 0) + qty;
      sheet.getRange(i + 1, 6).setValue(newQty);
      
      writeLog(user, "Restock", data[i][0], qty, targetSheetName, newQty);
      
      const msg = `✅ *เติมสต็อกสำเร็จ!*\n━━━━━━━━━━━━━━━\n📦 ${data[i][3]}\n🆔 ID: \`${data[i][0]}\`\n➕ เพิ่ม: *+${qty}* ชิ้น\n📊 ยอดใหม่: *${newQty}* ชิ้น`;
      notifyActivity(msg);
      return msg;
    }
  }
  return `❌ ไม่พบรหัส \`${code}\``;
}

function handleWithdrawText(chatId, args, user) {
  const code = args[1];
  const qty  = Number(args[2]);
  const dept = (args[3] || "CBR").toUpperCase();
  const colMap = { CBR: 6, CCS: 7, SKO: 8, RYG: 9, TRT: 10 };
  const col = colMap[dept];
  if (!col) return sendMsg(chatId, "❌ แผนกไม่ถูกต้อง\nใช้: cbr, ccs, sko, ryg, trt");
  sendMsg(chatId, withdraw(code, qty, col, dept, user, "STOCK"));
}

function handleDeptWithdraw(chatId, args, col, dept, user) {
  if (args.length < 3) return sendMsg(chatId, `⚠️ รูปแบบ: /dp-${dept.toLowerCase()} [ID] [จำนวน]`);
  sendMsg(chatId, withdraw(args[1], Number(args[2]), col, dept, user, args[3] || "STOCK"));
}

// ============================================================
// 7. SEARCH
// ============================================================
function sendSearchResult(chatId, keyword) {
  const sheet = getSpreadsheet().getSheetByName("STOCK");
  if (!sheet) return sendMsg(chatId, "❌ ไม่พบชีต STOCK");
  const data  = sheet.getDataRange().getValues();
  const kw    = keyword.toLowerCase();

  const deptColMap = { cbr: 6, ccs: 7, sko: 8, ryg: 9, trt: 10 };
  const isDeptSearch = Object.keys(deptColMap).includes(kw);

  let results = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const matchName = row[3].toString().toLowerCase().includes(kw); // แก้เป็น row[3]
    const matchId   = row[0].toString().toLowerCase() === kw;
    const matchDept = isDeptSearch && Number(row[deptColMap[kw]]) > 0;
    if (matchName || matchId || matchDept) results.push(row);
  }

  if (!results.length) return sendMsg(chatId, `❌ ไม่พบรายการที่ตรงกับ *"${keyword}"*`);

  const top = results.slice(0, 8);
  let msg = `🔍 *ผลการค้นหา "${keyword}"* (${results.length} รายการ)\n━━━━━━━━━━━━━━━\n\n`;
  top.forEach(r => {
    const alert = Number(r[5]) <= LOW_STOCK_LIMIT ? " ⚠️" : " ✅"; // แก้เป็น r[5]
    msg += `🔹 \`${r[0]}\` *${r[3].toString().substring(0,30)}*\n   คงเหลือ: *${r[5]}* ชิ้น${alert}\n\n`; // แก้เป็น r[3], r[5]
  });
  if (results.length > 8) msg += `_...และอีก ${results.length - 8} รายการ_\n`;

  const buttons = top.map(r => [{ text: `📦 ${r[0]} – ดูรายละเอียด`, callback_data: `info_${r[0]}` }]);
  buttons.push([{ text: "🏠 เมนูหลัก", callback_data: "main_menu" }]);

  sendToTelegram("sendMessage", { chat_id: chatId, text: msg, parse_mode: "Markdown", reply_markup: { inline_keyboard: buttons } });
}

// ============================================================
// 8. STOCK DETAIL
// ============================================================
function sendStockDetail(chatId, code) {
  const sheet = getSpreadsheet().getSheetByName("STOCK");
  if (!sheet) return sendMsg(chatId, "❌ ไม่พบชีต");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === code.toLowerCase()) {
      const r = data[i];
      const alert = Number(r[5]) <= LOW_STOCK_LIMIT ? "⚠️ ใกล้หมด" : "✅ ปกติ"; // แก้เป็น r[5]
      const msg = `📦 *รายละเอียดรายการ*\n━━━━━━━━━━━━━━━\n🆔 ID: \`${r[0]}\`\n📝 ชื่อ: ${r[3]}\n🔢 QTY: *${r[5]}* ชิ้น\n🏷 สถานะ: ${alert}\n\n*การเบิกตามแผนก:*\n🏢 CBR: ${r[6] || 0} | CCS: ${r[7] || 0}\n🏢 SKO: ${r[8] || 0} | RYG: ${r[9] || 0}`; // ขยับ Index ทั้งหมด
      const buttons = [
        [
          { text: "➖ เบิก CBR",  callback_data: `wd_${r[0]}_cbr` },
          { text: "➖ เบิก CCS",  callback_data: `wd_${r[0]}_ccs` }
        ],
        [
          { text: "➖ เบิก SKO",  callback_data: `wd_${r[0]}_sko` },
          { text: "➖ เบิก RYG",  callback_data: `wd_${r[0]}_ryg` }
        ],
        [{ text: "🏠 เมนูหลัก", callback_data: "main_menu" }]
      ];
      sendToTelegram("sendMessage", { chat_id: chatId, text: msg, parse_mode: "Markdown", reply_markup: { inline_keyboard: buttons } });
      return;
    }
  }
  sendMsg(chatId, `❌ ไม่พบรหัส \`${code}\``);
}

// ============================================================
// 9. REPORTS
// ============================================================
function getStockInfo(code, sheetName) {
  const data = getSpreadsheet().getSheetByName(sheetName).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === code.toLowerCase())
      return `📦 *ข้อมูลรายการ*\n━━━━━━━━━━━━━━━\n🆔 ID: \`${data[i][0]}\`\nชื่อ: ${data[i][3]}\nคงเหลือ: *${data[i][5]}* ชิ้น`; // แก้เป็น [3] และ [5]
  }
  return "❌ ไม่พบรายการ";
}

function getLowStock(sheetName) {
  const data = getSpreadsheet().getSheetByName(sheetName).getDataRange().getValues();
  const list = data.filter((r, i) => i > 0 && r[0] && Number(r[5]) <= LOW_STOCK_LIMIT)
                   .map(r => `⚠️ \`${r[0]}\` ${r[3].toString().substring(0,25)} *(${r[5]})*`);
  
  return list.length > 0
    ? `⚠️ *รายการใกล้หมด* (${list.length} รายการ)\n━━━━━━━━━━━━━━━\n\n` + list.join("\n")
    : "✅ ไม่มีรายการใกล้หมดครับ";
}

function getReportSummary(sheetName) {
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return "❌ ไม่พบชีต";
  const data = sheet.getDataRange().getValues();
  let totalQty = 0, low = 0;
  
  data.forEach((r, i) => { 
    if (i > 0 && r[0]) { 
      let qty = Number(r[5]) || 0; 
      totalQty += qty; 
      if (qty <= LOW_STOCK_LIMIT) low++; 
    } 
  });
  
  return `📊 *สรุปรายงานสต็อก*\n━━━━━━━━━━━━━━━\n📦 รายการทั้งหมด: *${data.length - 1}* รายการ\n🔢 QTY รวม: *${totalQty}* ชิ้น\n⚠️ ใกล้หมด: *${low}* รายการ`;
}

function sendDailySummary(chatId) {
  const ss       = getSpreadsheet();
  const logSheet = ss.getSheetByName("Logs");
  const today    = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy");

  let withdrawCount = 0, restockCount = 0, withdrawQty = 0, restockQty = 0;
  if (logSheet) {
    const logs = logSheet.getDataRange().getValues().slice(1);
    logs.forEach(r => {
      if (!r[0]) return;
      const logDate = Utilities.formatDate(new Date(r[0]), "GMT+7", "dd/MM/yyyy");
      if (logDate !== today) return;
      if (String(r[2]).toLowerCase().includes("withdraw")) { withdrawCount++; withdrawQty += Number(r[4]) || 0; }
      if (String(r[2]).toLowerCase().includes("restock"))  { restockCount++;  restockQty  += Number(r[4]) || 0; }
    });
  }

  const summary = getReportSummary("STOCK");
  const msg = `☀️ *สรุปรายงานประจำวัน*\n📅 ${today}\n━━━━━━━━━━━━━━━\n\n${summary}\n\n📋 *กิจกรรมวันนี้:*\n📉 เบิก: *${withdrawCount}* ครั้ง (${withdrawQty} ชิ้น)\n📈 เติม: *${restockCount}* ครั้ง (${restockQty} ชิ้น)`;
  sendMsg(chatId, msg);
}

function sendWeeklySummary(chatId) {
  const ss       = getSpreadsheet();
  const logSheet = ss.getSheetByName("Logs");
  const now      = new Date();
  const weekAgo  = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);

  let totalW = 0, totalR = 0, wQty = 0, rQty = 0;
  const itemMap = {};

  if (logSheet) {
    const logs = logSheet.getDataRange().getValues().slice(1);
    logs.forEach(r => {
      if (!r[0]) return;
      const d = new Date(r[0]);
      if (d < weekAgo) return;
      const action = String(r[2]).toLowerCase();
      const qty    = Number(r[4]) || 0;
      const code   = String(r[3]);
      if (action.includes("withdraw")) { totalW++; wQty += qty; itemMap[code] = (itemMap[code] || 0) + qty; }
      if (action.includes("restock"))  { totalR++; rQty += qty; }
    });
  }

  const topItems = Object.entries(itemMap).sort((a,b) => b[1]-a[1]).slice(0,3)
    .map((x,i) => `${i+1}. \`${x[0]}\` — ${x[1]} ชิ้น`).join("\n");

  const msg = `📈 *สรุปรายสัปดาห์ (7 วันที่ผ่านมา)*\n━━━━━━━━━━━━━━━\n📉 เบิกทั้งหมด: *${totalW}* ครั้ง (*${wQty}* ชิ้น)\n📈 เติมทั้งหมด: *${totalR}* ครั้ง (*${rQty}* ชิ้น)\n\n🏆 *Top 3 รายการที่เบิกมากสุด:*\n${topItems || "ไม่มีข้อมูล"}`;
  sendMsg(chatId, msg);
}

function getHistory(limit) {
  const sheet = getSpreadsheet().getSheetByName("Logs");
  if (!sheet) return "❌ ไม่พบประวัติ";
  const rows    = sheet.getDataRange().getValues().slice(1).reverse().slice(0, limit || 10);
  let msg = `📜 *ประวัติ ${limit || 10} รายการล่าสุด*\n━━━━━━━━━━━━━━━\n\n`;
  rows.forEach(r => {
    if (!r[0] || r[0] === "วันเวลา") return;
    const icon = String(r[2]).toLowerCase().includes("withdraw") ? "📉" : "📈";
    msg += `${icon} ${Utilities.formatDate(new Date(r[0]), "GMT+7", "dd/MM HH:mm")} | *${r[1]}*\n   ${r[2]} [\`${r[3]}\`] x${r[4]} → เหลือ ${r[6]}\n\n`;
  });
  return msg;
}

function sendAllStock(chatId, page, sheetName) {
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return;
  const rows       = sheet.getDataRange().getValues().slice(1).filter(r => r[0]);
  const totalPages = Math.ceil(rows.length / 10);
  const current    = rows.slice((page - 1) * 10, page * 10);
  let msg = `📦 *สต็อกทั้งหมด* หน้า ${page}/${totalPages}\n━━━━━━━━━━━━━━━\n\n`;
  current.forEach(r => {
    const alert = Number(r[5]) <= LOW_STOCK_LIMIT ? "⚠️" : "✅"; // แก้เป็น r[5]
    msg += `${alert} \`${r[0]}\` *${r[3].toString().substring(0, 20)}*\n   QTY: *${r[5]}* ชิ้น\n\n`; // แก้เป็น r[3], r[5]
  });
  const nav = [];
  if (page > 1)          nav.push({ text: "◀️ ก่อนหน้า", callback_data: `page_${page-1}_${sheetName}` });
  if (page < totalPages) nav.push({ text: "ถัดไป ▶️",    callback_data: `page_${page+1}_${sheetName}` });
  sendToTelegram("sendMessage", {
    chat_id: chatId, text: msg, parse_mode: "Markdown",
    reply_markup: { inline_keyboard: [nav, [{ text: "🏠 เมนูหลัก", callback_data: "main_menu" }]] }
  });
}

// ============================================================
// 10. LOGS & TELEGRAM HELPERS
// ============================================================
function writeLog(user, action, code, qty, note, remain) {
  const ss  = getSpreadsheet();
  const log = ss.getSheetByName("Logs") || ss.insertSheet("Logs");
  if (log.getLastRow() === 0)
    log.appendRow(["วันเวลา", "ชื่อผู้ใช้", "รายการ", "รหัสรายการ", "จำนวน", "เเผนก/ชื่อชีต", "ยอดเหลือสุทธิ"]);
  log.appendRow([new Date(), user, action, code, qty, note, remain]);
}

function sendMsg(chatId, text) {
  sendToTelegram("sendMessage", { chat_id: chatId, text: text, parse_mode: "Markdown" });
}

function notifyActivity(msg) {
  sendToTelegram("sendMessage", { chat_id: CHAT_ID, text: "📢 *Activity*\n" + msg, parse_mode: "Markdown" });
}

function sendToTelegram(method, payload) {
  return UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/${method}`, {
    method: "post", contentType: "application/json", payload: JSON.stringify(payload)
  });
}

function sendDailyStockPDF(chatId = CHAT_ID) {
  try {
    const blob = DriveApp.getFileById(getSpreadsheet().getId())
      .getAs("application/pdf")
      .setName("Stock_Report_" + Utilities.formatDate(new Date(), "GMT+7", "dd-MM-yyyy") + ".pdf");
    UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendDocument`, {
      method: "post", payload: { chat_id: chatId, document: blob, caption: "📄 รายงานสต็อกประจำวัน" }
    });
  } catch (e) { sendMsg(chatId, "❌ ไม่สามารถสร้าง PDF ได้"); }
}

// ============================================================
// 11. TRIGGERS
// ============================================================
function onEdit(e) {
  const range     = e.range;
  const sheetName = range.getSheet().getName();
  if (sheetName === "Logs") return;
  const user = Session.getActiveUser().getEmail() || "Unknown";
  notifyActivity(`✏️ *แก้ไขโดยตรง*\nชีต: ${sheetName} | เซลล์: ${range.getA1Notation()}\n${e.oldValue||"-"} → ${e.value||"-"}\nโดย: ${user}`);
}

function sendDailyReport() {
  const summary  = getReportSummary("STOCK");
  const lowStock = getLowStock("STOCK");
  notifyActivity(`☀️ *รายงานประจำวัน*\n\n${summary}\n\n${lowStock}`);
}

function sendWeeklyReport() {
  sendWeeklySummary(CHAT_ID);
}

// ============================================================
// 12. AI (Gemini)
// ============================================================
function callGemini(userMessage) {
  try {
    const sheet = getSpreadsheet().getSheetByName("STOCK");
    let ctx = "ข้อมูลสต็อกปัจจุบัน:\n";
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < Math.min(data.length, 51); i++)
        ctx += `- ${data[i][3]} (ID: ${data[i][0]}) คงเหลือ: ${data[i][5]}\n`; // แก้เป็น [3] และ [5]
    }
    const prompt = `คุณคือ SmartStock AI ของบริษัท System Innovation and Supply มีข้อมูลสต็อกรายการด้านล่างนี้ ตอบสั้นๆ กระชับ เป็นกันเอง ใช้ emoji ประกอบ ลงท้ายแนะนำ /menu\n\n${ctx}\n\nคำถาม: ${userMessage}`;
    const res  = UrlFetchApp.fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=${GEMINI_API_KEY}`,
      { method: "post", contentType: "application/json", payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }), muteHttpExceptions: true }
    );
    const json = JSON.parse(res.getContentText());
    return json.candidates?.[0]?.content
      ? "🤖 " + json.candidates[0].content.parts[0].text.trim()
      : "🤖 ขออภัยครับ ระบบ AI ไม่พร้อม ลองพิมพ์ /menu ครับ";
  } catch (e) {
    return "🤖 ระบบ AI พักผ่อนครับ";
  }
}

// ============================================================
// 13. DASHBOARD DATA
// ============================================================
function getStockData() {
  try {
    const sheet = getSpreadsheet().getSheetByName("STOCK");
    if (!sheet) return { error: "ไม่พบชีต STOCK" };

    const items = sheet.getDataRange().getValues().slice(1)
      .filter(r => r[0] !== "" && r[0] !== null)
      .map(r => ({
        id:           String(r[0] || ""),
        materialCode: String(r[2] || ""),
        description:  String(r[3] || ""),
        sn:           String(r[4] || ""),
        qty:          Number(r[5]) || 0,
        dpCBR:        Number(r[6]) || 0,
        dpCCS:        Number(r[7]) || 0,
        dpSKO:        Number(r[8]) || 0,
        dpRYG:        Number(r[9]) || 0,
        alert:        String(r[10] || "✅ ปกติ")
      }));

    const kpi = {
      totalItems:  items.length,
      totalQty:    items.reduce((s, r) => s + r.qty, 0),
      lowStock:    items.filter(r => r.alert.includes("⚠️")).length,
      normalStock: items.filter(r => r.alert.includes("✅")).length
    };

    const top10 = [...items].sort((a, b) => b.qty - a.qty).slice(0, 10)
      .map(r => ({
        label: r.description.substring(0, 25) + (r.description.length > 25 ? "…" : ""),
        value: r.qty
      }));

    const deptData = {
      CBR: items.reduce((s, r) => s + r.dpCBR, 0),
      CCS: items.reduce((s, r) => s + r.dpCCS, 0),
      SKO: items.reduce((s, r) => s + r.dpSKO, 0),
      RYG: items.reduce((s, r) => s + r.dpRYG, 0)
    };

    return { kpi, top10, deptData, items };
  } catch (e) {
    return { error: e.toString() };
  }
}

function getLogsData() {
  try {
    const sheet = getSpreadsheet().getSheetByName("Logs");
    if (!sheet) return { error: "ไม่พบชีต Logs" };

    const logs = sheet.getDataRange().getValues().slice(1)
      .reverse()
      .filter(r => r[0])
      .slice(0, 100)
      .map(r => ({
        timestamp: r[0] ? Utilities.formatDate(new Date(r[0]), "GMT+7", "dd/MM/yyyy HH:mm") : "-",
        username:  String(r[1] || "-"),
        action:    String(r[2] || "-"),
        code:      String(r[3] || "-"),
        qty:       Number(r[4]) || 0,
        dept:      String(r[5] || "-"),
        balance:   Number(r[6]) || 0
      }));

    return { logs };
  } catch (e) {
    return { error: e.toString() };
  }
}

function getLogsData() {
  try {
    const sheet = getSpreadsheet().getSheetByName("Logs");
    if (!sheet) return { error: "ไม่พบชีต Logs" };
    const logs = sheet.getDataRange().getValues().slice(1).reverse()
      .filter(r => r[0]).slice(0, 100)
      .map(r => ({
        timestamp: r[0] ? Utilities.formatDate(new Date(r[0]), "GMT+7", "dd/MM/yyyy HH:mm") : "-",
        username: String(r[1]||"-"), action: String(r[2]||"-"), code: String(r[3]||"-"),
        qty: Number(r[4])||0, dept: String(r[5]||"-"), balance: Number(r[6])||0
      }));
    return { logs };
  } catch(e) { return { error: e.toString() }; }
}

// ============================================================
// 14. SET WEBHOOK
// ============================================================
function setWebhook() {
  const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzdtBkRglc5Uuu8dgkWX_jZPuS28TchsfhyXIvp9LS5CSgV7Hsj1x5UqfU7tM3H-U0N/exec";  
  const res = UrlFetchApp.fetch(
    `https://api.telegram.org/bot${TELEGRAM_TOKEN}/setWebhook?url=${encodeURIComponent(WEB_APP_URL)}`
  );
  Logger.log(res.getContentText());
}

// ============================================================
// 15. AUTO LOW STOCK ALERT
// ============================================================
function checkAndAlertLowStock() {
  const sheet = getSpreadsheet().getSheetByName("STOCK");
  if (!sheet) return;

  const data     = sheet.getDataRange().getValues();
  const lowItems = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const qty = Number(row[5]); // แก้เป็น row[5]
    if (qty <= LOW_STOCK_LIMIT) {
      lowItems.push({ id: row[0], name: row[3], qty: qty }); // แก้เป็น row[3]
    }
  }

  if (lowItems.length === 0) return; 

  let msg = `🚨 *แจ้งเตือน: รายการใกล้หมด!*\n`;
  msg += `📅 ${Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm")}\n`;
  msg += `━━━━━━━━━━━━━━━\n\n`;

  lowItems.forEach(item => {
    const bar = item.qty <= 0 ? "🔴 หมดแล้ว!" : `🟡 เหลือ *${item.qty}* ชิ้น`;
    msg += `⚠️ \`${item.id}\` ${String(item.name).substring(0, 30)}\n   ${bar}\n\n`;
  });

  msg += `━━━━━━━━━━━━━━━\n`;
  msg += `📦 รวม *${lowItems.length}* รายการที่ต้องเติม`;

  sendToTelegram("sendMessage", {
    chat_id:    CHAT_ID,
    text:       msg,
    parse_mode: "Markdown",
    reply_markup: {
      inline_keyboard: [[
        { text: "📋 ดูรายการทั้งหมด", callback_data: "check_low" }
      ]]
    }
  });
}

// ============================================================
// 16. INSTALL EDIT TRIGGER
// ============================================================
function setupEditTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "onEdit") ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger("onEdit")
    .forSpreadsheet(SHEET_ID)
    .onEdit()
    .create();

  Logger.log("✅ ติดตั้ง onEdit Trigger สำเร็จ");
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function getDashboardStats() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("STOCK");
  const data = sheet.getDataRange().getValues();
  
  // ตัดแถวหัวตารางออก
  const rows = data.slice(1); 
  
  // 1. รายการทั้งหมด (นับจำนวนแถวที่มี ID)
  const totalSKU = rows.filter(row => row[0] !== "").length;
  
  // 2. QTY รวมทั้งหมด (บวกค่าใน Column F หรือ Index 5)
  const totalQty = rows.reduce((sum, row) => {
    const qty = parseFloat(row[5]);
    return sum + (isNaN(qty) ? 0 : qty);
  }, 0);

  return {
    totalSKU: totalSKU,
    totalQty: totalQty
  };
}

// เพิ่มฟังก์ชันสำหรับหน้า History (Logs)
function getLogsData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("Logs");
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1).reverse().slice(0, 100); // เอา 100 รายการล่าสุด

  const logs = rows.map(row => ({
    timestamp: Utilities.formatDate(new Date(row[0]), "GMT+7", "dd/MM/yyyy HH:mm"),
    username: row[1],
    action: row[2],
    code: row[3],
    qty: row[4],
    dept: row[5],
    balance: row[6]
  }));

  return { logs: logs };
}
