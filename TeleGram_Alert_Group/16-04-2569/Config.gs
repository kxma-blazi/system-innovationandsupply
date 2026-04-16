// ============================================================
// CONFIGURATION (รวมการตั้งค่าระบบ)
// ============================================================

// 1. การเชื่อมต่อ Telegram
var TELEGRAM_TOKEN = "8327163778:AAG2uPRh8V1F77ot03X1_DDRAsCNmdk4Wgo";
var CHAT_ID = "-1003731290917";

// 2. การเชื่อมต่อ Google Sheet
var SHEET_ID = "1MyAWKuCtmBclqVWALUEhEZjF7jKLMHs2sqL5oXjhA0w";

// 3. ตั้งค่าแผนก (Column Index ในหน้า STOCK)
// CBR=6, CCS=7, SKO=8, RYG=9, TRT=10
var DEPT_MAP = {
  "/dp-cbr": 6,
  "/dp-ccs": 7,
  "/dp-sko": 8,
  "/dp-ryg": 9,
  "/dp-trt": 10,
};
// https://generativelanguage.googleapis.com/v1beta/models?key=KEY-API-AI
// 4. อื่น
var GROQ_API_KEY = "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";

var LOW_STOCK_LIMIT = 1;
