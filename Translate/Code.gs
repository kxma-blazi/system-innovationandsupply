function doPost(e) {
  try {
    // รับข้อความภาษาจีนที่ส่งมาจากหน้าเว็บ
    var text = e.parameter.text;
    
    if (!text && e.postData && e.postData.contents) {
      text = JSON.parse(e.postData.contents).text;
    }

    if (!text || text.trim() === "") {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: "ไม่พบข้อความภาษาจีน" }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // สั่งแปลจาก จีน (zh) เป็น ไทย (th)
    var translatedText = LanguageApp.translate(text, 'zh', 'th');
    
    // ส่งผลลัพธ์กลับไปที่หน้าเว็บ
    var response = {
      success: true,
      originalText: text,
      translatedText: translatedText
    };
    
    return ContentService.createTextOutput(JSON.stringify(response))
                         .setMimeType(ContentService.MimeType.JSON);
                         
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput("API แปลภาษาพร้อมใช้งานแล้วครับ!");
}