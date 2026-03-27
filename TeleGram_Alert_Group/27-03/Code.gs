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
  const page = (e && e.parameter.page) || "stock";
  const search = (e && e.parameter.search) || "";

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const stockSheet = ss.getSheetByName("STOCK");
  const logSheet = ss.getSheetByName("Logs") || ss.insertSheet("Logs");

  const stockData = stockSheet ? stockSheet.getDataRange().getValues() : [];
  const logData = logSheet ? logSheet.getDataRange().getValues() : [];

  // กรองข้อมูลสต็อก
  let rows = stockData.slice(1).filter((r) => r[0]);
  if (search) {
    const kw = search.toLowerCase();
    rows = rows.filter(
      (r) =>
        r[0].toString().toLowerCase().includes(kw) ||
        r[3].toString().toLowerCase().includes(kw),
    );
  }

  // สร้างแถวตารางสต็อก
  const stockRows =
    rows
      .map((r) => {
        const qty = Number(r[5]) || 0; // คอลัมน์ F
        const badge =
          qty <= 0
            ? `<span class="badge red">หมด</span>`
            : qty <= LOW_STOCK_LIMIT
              ? `<span class="badge yellow">ใกล้หมด</span>`
              : `<span class="badge green">ปกติ</span>`;
        return `<tr>
      <td><code>${r[0]}</code></td>
      <td>${r[3] || "-"}</td>
      <td class="qty">${qty}</td>
      <td>${r[6] || 0}</td><td>${r[7] || 0}</td>
      <td>${r[8] || 0}</td><td>${r[9] || 0}</td><td>${r[10] || 0}</td>
      <td>${badge}</td>
    </tr>`;
      })
      .join("") || `<tr><td colspan="9" class="empty">ไม่พบข้อมูล</td></tr>`;

  // คำนวณตัวเลขสรุป Dashboard
  const allStockRows = stockData.slice(1).filter((r) => r[0]);
  const totalItems = allStockRows.length;
  const lowCount = allStockRows.filter((r) => {
    const q = Number(r[5]);
    return q > 0 && q <= LOW_STOCK_LIMIT;
  }).length;
  const outCount = allStockRows.filter((r) => Number(r[5]) <= 0).length;

  // สร้างแถวตาราง Log
  const logRows =
    logData
      .slice(1)
      .reverse()
      .slice(0, 50)
      .map((r) => {
        const dt = r[0]
          ? Utilities.formatDate(new Date(r[0]), "GMT+7", "dd/MM/yy HH:mm")
          : "-";
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
      })
      .join("") ||
    `<tr><td colspan="7" class="empty">ยังไม่มีประวัติ</td></tr>`;

  const tmpl = HtmlService.createTemplateFromFile("Index");
  tmpl.page = page;
  tmpl.search = search;
  tmpl.stockRows = stockRows;
  tmpl.logRows = logRows;
  tmpl.totalItems = totalItems;
  tmpl.lowCount = lowCount;
  tmpl.outCount = outCount;
  tmpl.resultCount = rows.length;

  return tmpl
    .evaluate()
    .setTitle("SYS Stock Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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

    // Mapping แผนกกับคอลัมน์
    const deptMap = {
      "/dp-cbr": 6,
      "/dp-ccs": 7,
      "/dp-sko": 8,
      "/dp-ryg": 9,
      "/dp-trt": 10,
    };

    if (cmd === "/start" || cmd === "/help" || cmd === "/menu") {
      const welcomeMsg =
        "📦 *ระบบจัดการสต็อก SYS*\n\n" +
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
    } else if (deptMap[cmd]) {
      if (args.length < 3) {
        sendTelegram(
          chatId,
          `⚠️ รูปแบบผิด! ต้องเป็น: \`${cmd} [รหัส] [จำนวน]\``,
        );
      } else {
        const result = withdraw(
          args[1],
          Number(args[2]),
          deptMap[cmd],
          cmd.replace("/dp-", "").toUpperCase(),
          user,
        );
        sendTelegram(chatId, result);
        // แจ้งเข้ากลุ่มหลัก (ตรวจสอบว่ามีตัวแปร CHAT_ID ที่เป็น ID กลุ่มอยู่ด้านบนโค้ดนะครับ)
        if (typeof CHAT_ID !== "undefined") {
          sendTelegram(
            CHAT_ID,
            `📣 *บันทึก:* ${result}\n(ทำรายการโดย: ${user})`,
          );
        }
      }
    } else if (cmd === "/stock") {
      sendTelegram(chatId, getStockInfo(args[1]));
    } else if (cmd === "/restock") {
      if (args.length < 3) {
        sendTelegram(
          chatId,
          "⚠️ รูปแบบผิด! ต้องเป็น: `/restock [รหัส] [จำนวน]`",
        );
      } else {
        sendTelegram(chatId, restock(args[1], Number(args[2]), user));
      }
    } else {
      // ✅ ส่งไปหา AI Gemini เมื่อไม่ใช่คำสั่ง
      // const aiResponse = callGemini(text);

      const aiResponse = callGemini(text, chatId);
      sendTelegram(chatId, aiResponse);
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
      if (currentTotal < qty)
        return `❌ ${code} ของไม่พอ (คงเหลือ ${currentTotal})`;

      const newTotal = currentTotal - qty;
      const newDeptTotal = (Number(data[i][colIndex]) || 0) + qty;

      // บันทึกลงชีต
      sheet.getRange(i + 1, 6).setValue(newTotal); // ช่องคงเหลือรวม
      sheet.getRange(i + 1, colIndex + 1).setValue(newDeptTotal); // ช่องแผนก

      // บันทึก Log
      writeLog(user, "Withdraw", code, qty, deptName, newTotal);

      let resMsg = `✅ *เบิกสำเร็จ*\n📦 รายการ: ${data[i][3]}\n📍 แผนก: ${deptName}\n📉 จำนวน: -${qty}\n📊 เหลือรวม: ${newTotal}`;
      if (newTotal <= LOW_STOCK_LIMIT)
        resMsg += `\n\n⚠️ *ALERT: สินค้าใกล้หมดแล้ว!*`;
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
      sheet.getRange(i + 1, 6).setValue(newQty);
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
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName("Logs");

    // 1. ถ้ายังไม่มี Sheet ชื่อ Logs ให้สร้างใหม่พร้อมใส่หัวตาราง
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

    // 2. บันทึกข้อมูล (เพิ่มคำสั่ง flush เพื่อให้เขียนลงไฟล์ทันที)
    sheet.appendRow([new Date(), user, act, code, qty, note, rem]);
    SpreadsheetApp.flush();

    Logger.log("✅ บันทึก Log สำเร็จ: " + code);
  } catch (e) {
    // ถ้าบันทึกไม่ได้ ให้แจ้งเตือนใน Logger ของระบบ
    Logger.log("❌ Error writeLog: " + e.toString());
  }
}

function getStockInfo(id) {
  if (!id) return "⚠️ ระบุรหัสสินค้า เช่น `/stock Q001`";
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const data = ss.getSheetByName("STOCK").getDataRange().getValues();

  const r = data.find(
    (x) =>
      x[0].toString().trim().toLowerCase() ===
      id.toString().trim().toLowerCase(),
  );
  if (!r) return "❌ ไม่พบรหัสสินค้า: " + id;

  const name = r[3] || "ไม่มีชื่อ";
  const total = r[5] === "" || isNaN(r[5]) ? 0 : r[5];
  const cbr = r[6] === "" || isNaN(r[6]) ? 0 : r[6];
  const ccs = r[7] === "" || isNaN(r[7]) ? 0 : r[7];
  const sko = r[8] === "" || isNaN(r[8]) ? 0 : r[8];
  const ryg = r[9] === "" || isNaN(r[9]) ? 0 : r[9];
  const trt = r[10] === "" || isNaN(r[10]) ? 0 : r[10]; // แก้จุดนี้: ถ้าว่างให้เป็น 0
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

    // ชีตที่ "ไม่ต้องการ" แจ้งเตือนสามารถใส่เพิ่มได้เรื่อยๆ
    const skipSheets = ["Logs", "Summary", "Config"];
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

// ฟังก์ชันสร้างและส่ง PDF รายงานสต็อกรายวัน
function sendDailyStockPDF() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("STOCK");
    const data = sheet.getDataRange().getValues();

    // 1. สร้าง HTML สำหรับแปลงเป็น PDF
    let html = "<h2>รายงานสรุปรายวัน</h2>";
    html +=
      "<table border='1' style='border-collapse: collapse; width: 100%;'>";
    html +=
      "<tr style='background-color: #f2f2f2;'><th>ID</th><th>Description</th><th>Total</th><th>CBR</th><th>CCS</th><th>SKO</th><th>RYG</th><th>TRT</th></tr>";

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === "") continue;
      html += `<tr>
                <td>${data[i][0]}</td>
                <td>${data[i][3]}</td>
                <td style='text-align:center;'>${data[i][5]}</td>
                <td style='text-align:center;'>${data[i][6]}</td>
                <td style='text-align:center;'>${data[i][7]}</td>
                <td style='text-align:center;'>${data[i][8]}</td>
                <td style='text-align:center;'>${data[i][9]}</td>
                <td style='text-align:center;'>${data[i][10]}</td>
               </tr>`;
    }
    html +=
      "</table><p>วันที่รายงาน: " +
      Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm") +
      "</p>";

    // 2. แปลง HTML เป็น Blob PDF
    const blob = HtmlService.createHtmlOutput(html)
      .getAs("application/pdf")
      .setName("Daily_Stock_Report.pdf");

    // 3. ส่งไฟล์ไปที่ Telegram
    const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendDocument`;
    const payload = {
      chat_id: CHAT_ID,
      document: blob,
      caption: "📊 รายงานสรุปสต็อกรายวันมาแล้วครับ!",
    };

    const options = {
      method: "post",
      payload: payload,
      muteHttpExceptions: true,
    };

    UrlFetchApp.fetch(url, options);
    Logger.log("ส่ง PDF เรียบร้อย");
  } catch (e) {
    Logger.log("Error ในการส่ง PDF: " + e.toString());
  }
}

// https://script.google.com/macros/s/AKfycbyuwxAzdc-GU1vWzuB8kTg737TBola4vBsMU380s9_fiHAoCYprjbvFxifIUDSnH0C7/exec?page=stock

// ฟังก์ชันระวังและแจ้งเตือนสินค้าใกล้หมด (รันอัตโนมัติจาก Trigger)
function checkAndAlertLowStock() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("STOCK");
    const data = sheet.getDataRange().getValues();

    let lowStockItems = [];

    // เริ่มวนลูปจากแถวที่ 2 (i=1) ข้ามหัวตาราง
    for (let i = 1; i < data.length; i++) {
      const id = data[i][0]; // คอลัมน์ A (รหัส)
      const name = data[i][3]; // คอลัมน์ D (ชื่อสินค้า)
      const qty = Number(data[i][5]); // คอลัมน์ F (คงเหลือรวม)

      if (id !== "" && qty <= LOW_STOCK_LIMIT) {
        const status = qty <= 0 ? "❌ หมดแล้ว" : "⚠️ ใกล้หมด";
        lowStockItems.push(`• ${id}: ${name} (คงเหลือ: ${qty}) [${status}]`);
      }
    }

    // ถ้ามีรายการที่ใกล้หมด ให้ส่งข้อความแจ้งเตือน
    if (lowStockItems.length > 0) {
      const alertMsg =
        "🚨 *แจ้งเตือนสถานะสต็อก (Low Stock Alert)*\n\n" +
        lowStockItems.join("\n") +
        "\n\n🔗 [ดู Dashboard เพิ่มเติม](https://script.google.com/macros/s/AKfycbyuwxAzdc-GU1vWzuB8kTg737TBola4vBsMU380s9_fiHAoCYprjbvFxifIUDSnH0C7/exec?page=stock)";

      sendTelegram(CHAT_ID, alertMsg);
      Logger.log("✅ ส่งแจ้งเตือน Low Stock เรียบร้อย");
    } else {
      Logger.log("ℹ️ ไม่พบสินค้าที่ใกล้หมด");
    }
  } catch (e) {
    Logger.log("❌ Error checkAndAlertLowStock: " + e.toString());
  }
}

// ============================================================
// AI Gemini (Version: Memory Support)
// ============================================================
function callGemini(msg, chatId) {
  try {
    const API_KEY = "AIzaSyB7VAmJ-2wPFml4lRAZzkHjsdJJjm0LR18";
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let memSheet = ss.getSheetByName("AI_Memory");

    // สร้าง Sheet ถ้ายังไม่มี
    if (!memSheet) {
      memSheet = ss.insertSheet("AI_Memory");
      memSheet.appendRow(["Chat ID", "History", "Last Updated"]);
    }

    // --- ส่วนดึงความจำเก่า ---
    const data = memSheet.getDataRange().getValues();
    let userRow = -1;
    let history = [];

    // ค้นหาประวัติการคุยตาม Chat ID
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === chatId.toString()) {
        userRow = i + 1;
        try {
          history = JSON.parse(data[i][1]); // แปลง JSON กลับเป็น Array
        } catch (e) {
          history = [];
        }
        break;
      }
    }

    // จำกัดความจำไว้ที่ 10 ประโยคล่าสุด (เพื่อไม่ให้ข้อมูลเยอะเกินไป)
    if (history.length > 10) history = history.slice(-10);

    const systemPrompt = `คุณคือ "SYS Bot" ผู้ช่วยประจำทีม System Innovation and Supply
ตอบภาษาไทยเป็นกันเอง สั้นกระชับ อารมณ์ดี 
ถ้าถามเรื่องสต็อกให้แนะนำคำสั่ง /menu ด้วย`;

    // --- เตรียมโครงสร้าง Payload สำหรับ Gemini ---
    // สร้างประวัติการคุยในรูปแบบที่ Gemini เข้าใจ
    let contents = history.map((h) => ({
      role: h.role, // "user" หรือ "model"
      parts: [{ text: h.text }],
    }));

    // เพิ่มคำถามปัจจุบันเข้าไปในลิสต์
    contents.push({ role: "user", parts: [{ text: msg }] });

    const payload = {
      contents: contents,
      systemInstruction: { parts: [{ text: systemPrompt }] }, // ส่งคำสั่งระบบแยกต่างหาก
    };

    const res = UrlFetchApp.fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${API_KEY}`,
      {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      },
    );

    const json = JSON.parse(res.getContentText());

    if (json.candidates && json.candidates[0] && json.candidates[0].content) {
      const aiText = json.candidates[0].content.parts[0].text.trim();

      // --- ส่วนบันทึกความจำใหม่ ---
      // เก็บทั้งคำถาม (user) และคำตอบ (model)
      history.push({ role: "user", text: msg });
      history.push({ role: "model", text: aiText });

      if (userRow !== -1) {
        // อัปเดตแถวเดิม
        memSheet.getRange(userRow, 2).setValue(JSON.stringify(history));
        memSheet.getRange(userRow, 3).setValue(new Date());
      } else {
        // เพิ่มแถวใหม่สำหรับ User คนนี้
        memSheet.appendRow([chatId, JSON.stringify(history), new Date()]);
      }

      return "🤖 " + aiText;
    } else {
      Logger.log("Gemini Error: " + res.getContentText());
      return "🤖 ขออภัยครับ ผมมึนๆ นิดหน่อย ลองใหม่อีกทีนะ";
    }
  } catch (e) {
    Logger.log("Memory System Error: " + e.toString());
    return "🤖 ระบบความจำขัดข้องชั่วคราวครับ";
  }
}

/*
// ============================================================
// AI (Gemini) V.1
// ============================================================
// ฟังก์ชันเรียกใช้ AI Gemini (รุ่นฟรี 1.5 Flash)
function callGemini(msg) {
  try {
    const API_KEY = "AIzaSyB7VAmJ-2wPFml4lRAZzkHjsdJJjm0LR18"; // คีย์ใหม่
    const systemPrompt = `คุณคือ "SYS Bot" ผู้ช่วยประจำทีม System Innovation and Supply
ตอบภาษาไทยเป็นกันเอง สั้นกระชับ อารมณ์ดี
ถ้าถามเรื่องสต็อกให้แนะนำคำสั่ง /menu ด้วย`;

    const payload = {
      contents: [{
        parts: [{ text: `${systemPrompt}\n\nคำถาม: ${msg}` }]
      }]
    };

    // เปลี่ยนรุ่นเป็น gemini-pro (ตัวนี้เสถียรสุดใน v1)
    const res = UrlFetchApp.fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${API_KEY}`,
      {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );

    const json = JSON.parse(res.getContentText());
    
    if (json.candidates && json.candidates[0] && json.candidates[0].content) {
      return "🤖 " + json.candidates[0].content.parts[0].text.trim();
    } else {
      Logger.log("Gemini API Error: " + res.getContentText());
      return "🤖 ขออภัยครับ ผมยังงงๆ กับคำถามนี้ ลองถามใหม่อีกทีนะ";
    }
  } catch (e) {
    Logger.log("Error: " + e.toString());
    return "🤖 ระบบขัดข้องชั่วคราวครับ";
  }
}
*/

// ฟังก์ชันสำหรับกด Test ดูว่าคีย์ผ่านไหม (กด Run อันนี้ได้เลย)
function debugKey() {
  const test = callGemini("สวัสดี ทดสอบหน่อยครับ");
  Logger.log("ผลลัพธ์จาก AI: " + test);
}

function testAI() {
  const result = callGemini("สวัสดี ทดสอบหน่อย");
  Logger.log(result);
}

function setBotCommands() {
  const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/setMyCommands`;
  const commands = [
    { command: "menu", description: "ดูเมนูการใช้งานทั้งหมด" },
    { command: "stock", description: "เช็คสต็อกสินค้า (ระบุรหัส)" },
    { command: "restock", description: "เติมสินค้าเข้าสต็อก" },
  ];

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ commands: commands }),
  };
  UrlFetchApp.fetch(url, options);
}
// ดูสถานะ
function debugKey() {
  Logger.log("Key ที่ใช้อยู่คือ: " + GEMINI_API_KEY);
  const test = callGemini("ทักทายหน่อย");
  Logger.log("ผลลัพธ์: " + test);
}

function checkMyModels() {
  const API_KEY = "AIzaSyB7VAmJ-2wPFml4lRAZzkHjsdJJjm0LR18"; // คีย์ใหม่ของคุณ
  const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${API_KEY}`;

  try {
    const res = UrlFetchApp.fetch(url);
    const json = JSON.parse(res.getContentText());
    Logger.log("รายชื่อรุ่นที่บัญชีคุณใช้ได้:");
    json.models.forEach((m) => {
      if (m.supportedGenerationMethods.includes("generateContent")) {
        Logger.log("- " + m.name);
      }
    });
  } catch (e) {
    Logger.log("Error เช็คโมเดล: " + e.toString());
  }
}
