// ============================================================
// 1. CONFIGURATION (ตั้งค่าระบบ)
// ============================================================
const TELEGRAM_TOKEN = "8327163778:AAG2uPRh8V1F77ot03X1_DDRAsCNmdk4Wgo";
const CHAT_ID = "-1003731290917"; 
const LOW_STOCK_LIMIT = 1;
const GEMINI_API_KEY = "AIzaSyDzLqjhgUeIzVO-ZoCcGFr_a7lhoq6JDYg";
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

  // กรองข้อมูลสต็อก
  let rows = stockData.slice(1).filter(r => r[0]);
  if (search) {
    const kw = search.toLowerCase();
    rows = rows.filter(r => 
      r[0].toString().toLowerCase().includes(kw) || 
      r[3].toString().toLowerCase().includes(kw)
    );
  }

  // สร้างแถวตารางสต็อก
  const stockRows = rows.map(r => {
    const qty = Number(r[5]) || 0; // คอลัมน์ F
    const badge = qty <= 0 ? `<span class="badge red">หมด</span>` : qty <= LOW_STOCK_LIMIT ? `<span class="badge yellow">ใกล้หมด</span>` : `<span class="badge green">ปกติ</span>`;
    return `<tr>
      <td><code>${r[0]}</code></td>
      <td>${r[3] || "-"}</td>
      <td class="qty">${qty}</td>
      <td>${r[6] || 0}</td><td>${r[7] || 0}</td>
      <td>${r[8] || 0}</td><td>${r[9] || 0}</td><td>${r[10] || 0}</td>
      <td>${badge}</td>
    </tr>`;
  }).join("") || `<tr><td colspan="9" class="empty">ไม่พบข้อมูล</td></tr>`;

  // คำนวณตัวเลขสรุป Dashboard
  const allStockRows = stockData.slice(1).filter(r => r[0]);
  const totalItems = allStockRows.length;
  const lowCount   = allStockRows.filter(r => {
    const q = Number(r[5]);
    return q > 0 && q <= LOW_STOCK_LIMIT;
  }).length;
  const outCount   = allStockRows.filter(r => Number(r[5]) <= 0).length;

  // สร้างแถวตาราง Log
  const logRows = logData.slice(1).reverse().slice(0, 50).map(r => {
    const dt = r[0] ? Utilities.formatDate(new Date(r[0]), "GMT+7", "dd/MM/yy HH:mm") : "-";
    const ac = r[2] === "Withdraw" ? "withdraw" : "restock";
    return `<tr>
      <td>${dt}</td>
      <td>${r[1] || "-"}</td>
      <td><span class="action ${ac}">${r[2] || "-"}</span></td>
      <td><code>${r[3] || "-"}</code></td>
      <td>${r[4] || 0}</td>
      <td>${r[5] || "-"}</td>
      <td class="qty">${r[6] || 0}</td>
    </tr>`;
  }).join("") || `<tr><td colspan="7" class="empty">ยังไม่มีประวัติ</td></tr>`;

  const tmpl = HtmlService.createTemplateFromFile("Index");
  tmpl.page        = page;
  tmpl.search      = search;
  tmpl.stockRows   = stockRows;
  tmpl.logRows     = logRows;
  tmpl.totalItems  = totalItems;
  tmpl.lowCount    = lowCount;
  tmpl.outCount    = outCount;
  tmpl.resultCount = rows.length;

  return tmpl.evaluate().setTitle("SYS Stock Dashboard").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// 3. TELEGRAM LOGIC (ระบบสั่งการผ่านบอท)
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.callback_query) return; // ข้ามปุ่มกดถ้ายังไม่ได้ใช้
    if (!data.message || !data.message.text) return;
    
    const chatId = data.message.chat.id;
    const text = data.message.text.trim();
    const user = data.message.from.first_name || "User";
    const args = text.split(/\s+/);
    const cmd = args[0].toLowerCase();

    // Mapping แผนกกับคอลัมน์
    const deptMap = { 
    "/dp-cbr": 6, 
    "/dp-ccs": 7, 
    "/dp-sko": 8, 
    "/dp-ryg": 9, 
    "/dp-trt": 10 // คอลัมน์ K (Index 10)
  };

    if (cmd === "/start" || cmd === "/menu") {
      const welcomeMsg = "📦 *ระบบจัดการสต็อก SYS*\n\n" +
                         "🔹 *การเบิกสินค้า:* (ระบุรหัสและจำนวน)\n" +
                         "• `/dp-cbr [ID] [จำนวน]`\n" +
                         "• `/dp-ccs [ID] [จำนวน]`\n" +
                         "• `/dp-sko [ID] [จำนวน]`\n" +
                         "• `/dp-ryg [ID] [จำนวน]`\n" +
                         "• `/dp-trt [ID] [จำนวน]`\n\n" +
                         "🔹 *อื่นๆ:*\n" +
                         "• `/stock [ID]` - เช็คสต็อก\n" +
                         "• `/restock [ID] [จำนวน]` - เติมของ";
      sendTelegram(chatId, welcomeMsg);
    } 
    else if (deptMap[cmd]) {
      if (args.length < 3) {
        sendTelegram(chatId, `⚠️ รูปแบบผิด! ต้องเป็น: \`${cmd} [รหัส] [จำนวน]\``);
      } else {
        const result = withdraw(args[1], Number(args[2]), deptMap[cmd], cmd.replace("/dp-","").toUpperCase(), user);
        sendTelegram(chatId, result);
        sendTelegram(CHAT_ID, `📣 *บันทึก:* ${result}\n(ทำรายการโดย: ${user})`); // แจ้งเข้ากลุ่มหลัก
      }
    } 
    else if (cmd === "/stock") {
      sendTelegram(chatId, getStockInfo(args[1]));
    } 
    else if (cmd === "/restock") {
      if (args.length < 3) {
        sendTelegram(chatId, "⚠️ รูปแบบผิด! ต้องเป็น: `/restock [รหัส] [จำนวน]`");
      } else {
        sendTelegram(chatId, restock(args[1], Number(args[2]), user));
      }
    } 
    else {
      // ถ้าไม่มีคำสั่งตรง ให้ส่งไปถาม Gemini AI
      sendTelegram(chatId, callGemini(text));
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
    if (data[i][0].toString().toLowerCase() === code.toString().toLowerCase()) {
      const currentTotal = Number(data[i][5]); // คอลัมน์ F
      if (currentTotal < qty) return `❌ ${code} ของไม่พอ (คงเหลือ ${currentTotal})`;
      
      const newTotal = currentTotal - qty;
      const newDeptTotal = (Number(data[i][colIndex]) || 0) + qty;
      
      // บันทึกลงชีต
      sheet.getRange(i+1, 6).setValue(newTotal); // ช่องคงเหลือรวม
      sheet.getRange(i+1, colIndex+1).setValue(newDeptTotal); // ช่องแผนก
      
      // บันทึก Log
      writeLog(user, "Withdraw", code, qty, deptName, newTotal);
      
      let resMsg = `✅ *เบิกสำเร็จ*\n📦 รายการ: ${data[i][3]}\n📍 แผนก: ${deptName}\n📉 จำนวน: -${qty}\n📊 เหลือรวม: ${newTotal}`;
      if (newTotal <= LOW_STOCK_LIMIT) resMsg += `\n\n⚠️ *ALERT: สินค้าใกล้หมดแล้ว!*`;
      return resMsg;
    }
  }
  return `❌ ไม่พบรหัสสินค้า "${code}" ในระบบ`;
}

// ฟังก์ชันเติมสต็อก
function restock(code, qty, user) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("STOCK");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === code.toString().toLowerCase()) {
      const newQty = (Number(data[i][5]) || 0) + qty;
      sheet.getRange(i+1, 6).setValue(newQty);
      writeLog(user, "Restock", code, qty, "STOCK", newQty);
      return `✅ เติมสต็อกสำเร็จ\n📦 ${data[i][3]}\n➕ เพิ่ม: ${qty}\n📊 ยอดรวมใหม่: ${newQty}`;
    }
  }
  return "❌ ไม่พบรหัสสินค้าเพื่อเติมสต็อก";
}

// ============================================================
// 4. UTILITY FUNCTIONS (ฟังก์ชันช่วยทำงาน)
// ============================================================

function writeLog(user, act, code, qty, note, rem) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("Logs") || ss.insertSheet("Logs");
  sheet.appendRow([new Date(), user, act, code, qty, note, rem]);
}

function getStockInfo(id) {
  if (!id) return "⚠️ ระบุรหัสสินค้า เช่น `/stock Q001`";
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const data = ss.getSheetByName("STOCK").getDataRange().getValues();
  
  const r = data.find(x => x[0].toString().trim().toLowerCase() === id.toString().trim().toLowerCase());
  if (!r) return "❌ ไม่พบรหัสสินค้า: " + id;

  const name   = r[3] || "ไม่มีชื่อ";
  const total  = (r[5] === "" || isNaN(r[5])) ? 0 : r[5];
  const cbr    = (r[6] === "" || isNaN(r[6])) ? 0 : r[6];
  const ccs    = (r[7] === "" || isNaN(r[7])) ? 0 : r[7];
  const sko    = (r[8] === "" || isNaN(r[8])) ? 0 : r[8];
  const ryg    = (r[9] === "" || isNaN(r[9])) ? 0 : r[9];
  const trt    = (r[10] === "" || isNaN(r[10])) ? 0 : r[10]; // แก้จุดนี้: ถ้าว่างให้เป็น 0
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

function callGemini(msg) {
  try {
    const payload = { contents: [{ parts: [{ text: `คุณคือ SmartStock AI ตอบสั้นๆ เป็นกันเอง แนะนำให้ใช้คำสั่ง /menu คำถาม: ${msg}` }] }] };
    const res = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${GEMINI_API_KEY}`, { method: "post", contentType: "application/json", payload: JSON.stringify(payload) });
    return "🤖 " + JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
  } catch (e) {
    return "🤖 รับทราบครับ (Gemini ไม่ว่างตอบ)";
  }
}

// แจ้งเตือนเมื่อแก้ชีตโดยตรง
function onEdit(e) {
  try {
    const range = e.range;
    const sheetName = range.getSheet().getName();
    if (sheetName === "Logs") return;
    const msg = `⚠️ *มีการแก้ไขข้อมูลโดยตรง*\n📍 ชีต: ${sheetName}\n📌 ช่อง: ${range.getA1Notation()}\n🔄 เปลี่ยนเป็น: ${e.value || "ว่าง"}`;
    sendTelegram(CHAT_ID, msg);
  } catch (err) {}
}
