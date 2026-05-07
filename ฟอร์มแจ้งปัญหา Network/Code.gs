/**
 * ตั้งค่าเบื้องต้น - ตรวจสอบ ID ให้ถูกต้อง
 */
const SHEET_ID = "1DeM68CrkXnY-xWBFd21UfKAofWA0ZkGlmBFpAZbi6Ro"; 
const FOLDER_ID = "1XcJ1zjWKPXg_vj5FDFv7hn_SqVu2eFYc"; 
const ADMIN_EMAIL = "okumakung2018@gmail.com";

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle('Network Alert System');
}

function processAll(data) {
  try {
    // 1. เชื่อมต่อ Spreadsheet (ใช้ ID เพื่อป้องกันข้อผิดพลาด getActive)
    const ss = SpreadsheetApp.openById(SHEET_ID); 
    const sheet = ss.getSheets()[0]; 
    
    let imageUrl = "ไม่มีรูปภาพ";

    // 2. จัดการอัปโหลดรูปภาพเข้า Drive
    if (data.imageFile && data.imageFile.base64) {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const contentType = data.imageFile.type;
      const base64Data = data.imageFile.base64.split(',')[1];
      const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, `Alert_${data.siteId}_${new Date().getTime()}.jpg`);
      const file = folder.createFile(blob);
      imageUrl = file.getUrl();
    }

    // 3. บันทึกข้อมูลลง Google Sheets
    sheet.appendRow([
      new Date(),           // คอลัมน์ A: วันที่ทำรายการ
      data.datetime,        // คอลัมน์ B: วันเวลาที่เลือกในฟอร์ม
      data.fullname,        // คอลัมน์ C: ชื่อ-นามสกุล
      data.department,      // คอลัมน์ D: หน่วยงาน
      data.phone,           // คอลัมน์ E: เบอร์ติดต่อ
      data.email,           // คอลัมน์ F: อีเมล
      data.siteId,          // คอลัมน์ G: Site ID
      data.issueType,       // คอลัมน์ H: ประเภทปัญหา
      data.description,     // คอลัมน์ I: รายละเอียด
      imageUrl              // คอลัมน์ J: ลิงก์รูปภาพ
    ]);

    // 4. ส่งอีเมลแจ้งเตือน
    sendNotification(data, imageUrl);
    
    return { status: "success", message: "ส่งข้อมูลสำเร็จ! ระบบบันทึกข้อมูลและแจ้ง IT เรียบร้อยแล้ว" };
  } catch (e) {
    return { status: "error", message: "เกิดข้อผิดพลาด: " + e.toString() };
  }
}

function sendNotification(data, imageUrl) {
  const subject = `🚨 แจ้งปัญหา Network: Site ${data.siteId} โดย ${data.fullname}`;
  const htmlBody = `
    <div style="font-family: 'Sarabun', sans-serif; padding: 20px; border: 1px solid #2ecc71; border-radius: 10px;">
      <h2 style="color: #27ae60;">บันทึกการแจ้งปัญหา Network</h2>
      <p><strong>ผู้แจ้ง:</strong> ${data.fullname} (${data.department})</p>
      <p><strong>เบอร์ติดต่อ:</strong> ${data.phone}</p>
      <p><strong>อีเมล:</strong> ${data.email || '-'}</p>
      <p><strong>Site ID:</strong> ${data.siteId}</p>
      <p><strong>ประเภทปัญหา:</strong> ${data.issueType}</p>
      <p><strong>รายละเอียด:</strong> ${data.description}</p>
      <p><strong>รูปหลักฐาน:</strong> <a href="${imageUrl}">คลิกเพื่อดูรูปภาพ</a></p>
      <hr>
      <p style="color: #7f8c8d; font-size: 12px;">แจ้งจากระบบอัตโนมัติ Network Alert System</p>
    </div>
  `;
  
  MailApp.sendEmail({
    to: ADMIN_EMAIL,
    subject: subject,
    htmlBody: htmlBody
  });
}


// TEST
function testAuth() {
  DriveApp.getRootFolder();
  MailApp.getRemainingDailyQuota();
  SpreadsheetApp.openById(SHEET_ID);
  Logger.log("ถ้าเห็นข้อความนี้ แสดงว่าสิทธิ์ผ่านแล้ว!");
}

function forceAuth() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  folder.createFile("test.txt", "test"); // บังคับให้ระบบเช็คสิทธิ์ createFile
}

// เพิ่มฟังก์ชันนี้ไว้ล่างสุดของไฟล์ รหัส.gs
function getServiceUrl() {
  return ScriptApp.getService().getUrl();
}