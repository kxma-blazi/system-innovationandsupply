// ============================================================
// 1. CONFIGURATION (ตั้งค่าระบบ)
// ============================================================
const TELEGRAM_TOKEN = "8327163778:AAG2uPRh8V1F77ot03X1_DDRAsCNmdk4Wgo";
const CHAT_ID = "-1003731290917"; 
const LOW_STOCK_LIMIT = 1;
const GEMINI_API_KEY = "AIzaSyDkquk7-7tIHU3vzqFZXh-Eq1DkBWIcZ1w";
const SHEET_ID = "1MyAWKuCtmBclqVWALUEhEZjF7jKLMHs2sqL5oXjhA0w";

// ============================================================
// 2. WEB DASHBOARD (หน้าเว็บแสดงผล)
// ============================================================
function doGet(e) {
  const page   = (e && e.parameter.page)   || "stock";
  const search = (e && e.parameter.search) || "";
  
  const ss         = SpreadsheetApp.openById(SHEET_ID);
  const stockSheet = ss.getSheetByName("STOCK");
  const logSheet   = ss.getSheetByName("Logs") || ss.insertSheet("Logs");

  const stockData = stockSheet ? stockSheet.getDataRange().getValues() : [];
  const logData   = logSheet   ? logSheet.getDataRange().getValues()   : [];

  // 1. กรองข้อมูลสต็อก
  let rows = stockData.slice(1).filter(r => r[0]);
  if (search) {
    const kw = search.toLowerCase();
    rows = rows.filter(r => 
      r[0].toString().toLowerCase().includes(kw) || 
      r[3].toString().toLowerCase().includes(kw)
    );
  }

  // 2. สร้างแถวตารางสต็อก
  const stockRows = rows.map(r => {
    const qty = Number(r[5]) || 0;
    const badge = qty <= 0 ? `<span class="badge red">หมด</span>` : qty <= LOW_STOCK_LIMIT ? `<span class="badge yellow">ใกล้หมด</span>` : `<span class="badge green">ปกติ</span>`;
    return `<tr><td><code>${r[0]}</code></td><td>${r[3] || "-"}</td><td class="qty">${qty}</td><td>${r[6] || 0}</td><td>${r[7] || 0}</td><td>${r[8] || 0}</td><td>${r[9] || 0}</td><td>${r[10] || 0}</td><td>${badge}</td></tr>`;
  }).join("");

  const allStockRows = stockData.slice(1).filter(r => r[0]);
  const totalItems = allStockRows.length;
  const lowCount   = allStockRows.filter(r => { const q = Number(r[5]); return q > 0 && q <= LOW_STOCK_LIMIT; }).length;
  const outCount   = allStockRows.filter(r => Number(r[5]) <= 0).length;

  const logRows = logData.slice(1).reverse().slice(0, 50).map(r => {
    const dt = r[0] ? Utilities.formatDate(new Date(r[0]), "GMT+7", "dd/MM/yy HH:mm") : "-";
    const ac = r[2] === "Withdraw" ? "withdraw" : "restock";
    return `<tr><td>${dt}</td><td>${r[1] || "-"}</td><td><span class="action ${ac}">${r[2] || "-"}</span></td><td><code>${r[3] || "-"}</code></td><td>${r[4] || 0}</td><td>${r[5] || "-"}</td><td class="qty">${r[6] || 0}</td></tr>`;
  }).join("");

  // 5. แก้ไขการสร้าง deliveryRows ให้ดึงข้อมูลมาโชว์
  let deliveryRows = "";
  if (page === 'delivery') {
    const deliveryData = getDeliveryRecords(100);
    deliveryRows = (deliveryData.records || []).map(r => {
      return `<tr>
        <td><code>${r.code}</code></td>
        <td><code>${r.serial}</code></td>
        <td>${r.dateDelivered}</td>
        <td><span class="action withdraw">${r.note}</span></td>
        <td>${r.deliveredBy}</td>
      </tr>`;
    }).join("") || `<tr><td colspan="5" class="empty">ยังไม่มีการนำจ่ายซีเรียล</td></tr>`;
  }

  const tmpl = HtmlService.createTemplateFromFile("Index");
  tmpl.page        = page;
  tmpl.search      = search;
  tmpl.stockRows   = stockRows;
  tmpl.logRows     = logRows;
  tmpl.deliveryRows = deliveryRows; // ส่งตัวแปรนี้ไปที่หน้าเว็บ
  tmpl.totalItems  = totalItems;
  tmpl.lowCount    = lowCount;
  tmpl.outCount    = outCount;
  tmpl.resultCount = rows.length;

  return tmpl.evaluate().setTitle("SYS Stock Dashboard").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
  } catch (e) { return { records: [] }; }
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

    const deptMap = { "/dp-cbr": 6, "/dp-ccs": 7, "/dp-sko": 8, "/dp-ryg": 9, "/dp-trt": 10 };

    if (cmd === "/start" || cmd === "/help" || cmd === "/menu") {
      const welcomeMsg = "📦 *ระบบจัดการสต็อก SYS (Auto-Detect)*\n\n" +
                         "🔹 *วิธีเบิกสินค้า:*\n" +
                         "• ใส่ซีเรียลหรือรหัส: \`/dp-xxx [SN หรือ ID]\`\n" +
                         "• ระบุจำนวนเพิ่ม: \`/dp-xxx [ID] [จำนวน]\` \n\n" +
                         "*(ระบบจะค้นหาซีเรียลก่อน ถ้าไม่เจอจะเบิกเป็นจำนวนให้เอง)*";
      sendTelegram(chatId, welcomeMsg);
    }
    else if (deptMap[cmd]) {
      const input = args[1]; 
      if (!input) {
        sendTelegram(chatId, `⚠️ กรุณาระบุรหัสสินค้าหรือซีเรียล\nเช่น: \`${cmd} SN999\` หรือ \`${cmd} Q001\``);
        return;
      }

      // --- ขั้นตอนที่ 1: ลองเบิกแบบซีเรียลก่อน ---
      const serialResult = withdrawBySerial(input, deptMap[cmd], cmd.replace("/dp-","").toUpperCase(), user);

      // ถ้าเจอซีเรียล (เบิกสำเร็จ หรือ ถูกเบิกไปแล้ว)
      if (serialResult !== "NOT_FOUND") {
        sendTelegram(chatId, serialResult);
        if (typeof CHAT_ID !== 'undefined' && !serialResult.includes("❌")) {
          sendTelegram(CHAT_ID, `📣 *บันทึก (SN):* ${serialResult}\n(โดย: ${user})`);
        }
        return;
      }

      // --- ขั้นตอนที่ 2: ถ้าไม่ใช่ซีเรียล (NOT_FOUND) ให้เบิกแบบรหัสสินค้า (ID) ---
      let qtyMatch = (args[2] || "1").match(/\d+/); 
      let qty = qtyMatch ? parseInt(qtyMatch[0], 10) : 1;

      const result = withdraw(input, qty, deptMap[cmd], cmd.replace("/dp-","").toUpperCase(), user);
      sendTelegram(chatId, result);
      
      if (typeof CHAT_ID !== 'undefined' && !result.includes("❌")) {
        sendTelegram(CHAT_ID, `📣 *บันทึก (QTY):* ${result}\n(โดย: ${user})`);
      }
    } 
    // ... ส่วนของ /stock และ /restock ใช้ของเดิมที่เฮียส่งมาได้เลยครับ ...
    else if (cmd === "/stock") {
      sendTelegram(chatId, getStockInfo(args[1]));
    } 
    else if (cmd === "/restock") {
       if (args.length < 3) {
         sendTelegram(chatId, "⚠️ รูปแบบผิด! ต้องเป็น: `/restock [รหัส] [จำนวน]`");
       } else {
         let qtyMatch = args[2].match(/\d+/); 
         let qty = qtyMatch ? parseInt(qtyMatch[0], 10) : NaN;
         sendTelegram(chatId, restock(args[1], qty, user));
       }
    } 
    else {
      sendTelegram(chatId, callGemini(text, chatId));
    }
  } catch (err) {
    Logger.log("doPost Error: " + err.toString());
  }
}

// ฟังก์ชันเบิกสินค้า (แกนหลัก)
function withdraw(code, qty, colIndex, deptName, user) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("STOCK");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    // เปลี่ยนจาก [0] เป็น [2] เพื่อให้ค้นหา Material Code ในคอลัมน์ C
    if (data[i][2].toString().toLowerCase() === code.toString().toLowerCase()) {
      const currentTotal = Number(data[i][5]); 
      if (currentTotal < qty) return `❌ ${code} ของไม่พอ (คงเหลือ ${currentTotal})`;
      
      const newTotal = currentTotal - qty;
      const newDeptTotal = (Number(data[i][colIndex]) || 0) + qty;
      
      sheet.getRange(i+1, 6).setValue(newTotal);
      sheet.getRange(i+1, colIndex+1).setValue(newDeptTotal);
      
      writeLog(user, "Withdraw", code, qty, deptName, newTotal);
      
      let resMsg = `✅ *เบิกสำเร็จ*\n📦 รายการ: ${data[i][3]}\n📍 แผนก: ${deptName}\n📉 จำนวน: -${qty}\n📊 เหลือรวม: ${newTotal}`;
      if (newTotal <= LOW_STOCK_LIMIT) resMsg += `\n\n⚠️ *ALERT: สินค้าใกล้หมดแล้ว!*`;
      return resMsg;
    }
  }
  return `❌ ไม่พบรหัสสินค้า "${code}" ในระบบ`;
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
// AI Gemini (Version: Memory Support)
// ============================================================
function callGemini(msg, chatId) {
  try {
    // ✅ [แก้ไขแล้ว] ให้ลองดึงจาก Script Properties ก่อน ถ้าไม่มีให้ใช้ตัวแปร GEMINI_API_KEY ด้านบนสุด
    const API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || GEMINI_API_KEY;

    if (!API_KEY) return "❌ ยังไม่ได้ตั้งค่า API KEY";

    const ss = SpreadsheetApp.openById(SHEET_ID);
    let memSheet = ss.getSheetByName("AI_Memory");

    if (!memSheet) {
      memSheet = ss.insertSheet("AI_Memory");
      memSheet.appendRow(["Chat ID", "History", "Last Updated"]);
    }

    const data = memSheet.getDataRange().getValues();
    let userRow = -1;
    let history = [];

    for (let i = 1; i < data.length; i++) {
      if (!data[i] || !data[i][0]) continue; 
      const rowChatId = data[i][0];
      if (rowChatId === undefined || rowChatId === null) continue; 
      if (rowChatId.toString() === (chatId || "").toString()) {
        userRow = i + 1;
        try {
          history = JSON.parse(data[i][1] || "[]");
        } catch (e) {
          history = [];
        }
        break;
      }
    }

    if (history.length > 10) history = history.slice(-10);

    const systemPrompt = `คุณคือ SYS Bot ตอบไทย สั้น กระชับ เป็นกันเอง`;

    let contents = history.map(h => ({
      role: h.role,
      parts: [{ text: h.text }]
    }));

    contents.push({
      role: "user",
      parts: [{ text: msg }]
    });

    const payload = {
      contents: contents,
      systemInstruction: {
        parts: [{ text: systemPrompt }]
      }
    };

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${API_KEY}`;

    const res = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const status = res.getResponseCode();
    const body = res.getContentText();

    if (status !== 200) {
      Logger.log("❌ Gemini Error: " + body);
      return "🤖 AI ใช้งานไม่ได้ (" + status + ")";
    }

    const json = JSON.parse(body);
    const aiText = json?.candidates?.[0]?.content?.parts?.[0]?.text;

    if (!aiText) {
      Logger.log("⚠️ Gemini ไม่มีคำตอบ: " + body);
      return "🤖 งงนิดหน่อย ลองใหม่อีกทีนะ";
    }

    // --- save memory ---
    history.push({ role: "user", text: msg });
    history.push({ role: "model", text: aiText });

    if (userRow !== -1) {
      memSheet.getRange(userRow, 2).setValue(JSON.stringify(history));
      memSheet.getRange(userRow, 3).setValue(new Date());
    } else {
      memSheet.appendRow([chatId, JSON.stringify(history), new Date()]);
    }

    return "🤖 " + aiText.trim();

  } catch (e) {
    Logger.log("🔥 callGemini Error: " + e.toString());
    return "🤖 ระบบ AI ล่มชั่วคราว";
  }
}