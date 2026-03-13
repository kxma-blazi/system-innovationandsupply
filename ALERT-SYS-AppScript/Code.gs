// Setting Variable (Bot Token)
const TELEGRAM_TOKEN = "8327163778:AAFM4aKpxT29WTB4z_StzKEEcRGrSBDS2_s";
const CHAT_ID = "-1003731290917";
const LOW_STOCK_LIMIT = 1;

// Used Func for Test Connect
function test(){
sendTelegram("-1003731290917","TEST STOCK BOT");
}


function sendTelegram(chatId,msg){

  if(!msg) return;

  const url = "https://api.telegram.org/bot"+TELEGRAM_TOKEN+"/sendMessage";

  const payload = {
    chat_id:chatId,
    text:msg,
    parse_mode:"HTML"
  };

  UrlFetchApp.fetch(url,{
    method:"post",
    contentType:"application/json",
    payload:JSON.stringify(payload)
  });

}

function getSheet(){
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("STOCK");
}
  // doPost Func
function doPost(e){

  const data = JSON.parse(e.postData.contents);

  if(!data.message) return;

  const chatId = data.message.chat.id;
  const text = data.message.text;

  const args = text.split(" ");
  const cmd = args[0].toLowerCase();

  if(cmd==="/help"){
    sendTelegram(chatId,helpMenu());
  }

  if(cmd==="/stock"){
    sendTelegram(chatId,getStock(args[1]));
  }

  if(cmd==="/search"){
    sendTelegram(chatId,searchStock(args[1]));
  }

  if(cmd==="/allstock"){
    sendTelegram(chatId,allStock());
  }

  if(cmd==="/lowstock"){
    sendTelegram(chatId,lowStock());
  }

  if(cmd==="/restock"){
    sendTelegram(chatId,restock(args[1],Number(args[2])));
  }

  if(cmd==="/dp-cbr"){
    sendTelegram(chatId,withdraw(args[1],Number(args[2]),5));
  }

  if(cmd==="/dp-ccs"){
    sendTelegram(chatId,withdraw(args[1],Number(args[2]),6));
  }

  if(cmd==="/dp-sko"){
    sendTelegram(chatId,withdraw(args[1],Number(args[2]),7));
  }

  if(cmd==="/dp-ryg"){
    sendTelegram(chatId,withdraw(args[1],Number(args[2]),8));
  }

  if(cmd==="/dp-trt"){
    sendTelegram(chatId,withdraw(args[1],Number(args[2]),9));
  }

  if(cmd==="/report"){
    sendTelegram(chatId,report());
  }

}

function helpMenu(){

return `
📦 STOCK BOT

/stock CODE
/search KEYWORD

/allstock
/lowstock
/report

/dp-cbr CODE QTY
/dp-ccs CODE QTY
/dp-sko CODE QTY
/dp-ryg CODE QTY
/dp-trt CODE QTY

/restock CODE QTY
`;

}

function getStock(code){

const data=getSheet().getDataRange().getValues();

for(let i=1;i<data.length;i++){

if(data[i][0]==code){

return `📦 ${data[i][0]}
${data[i][2]}
QTY : ${data[i][4]}`;

}

}

return "❌ ไม่พบสินค้า";

}

function restock(code,qty){

const sheet=getSheet();
const data=sheet.getDataRange().getValues();

for(let i=1;i<data.length;i++){

if(data[i][0]==code){

let newQty=data[i][4]+qty;

sheet.getRange(i+1,5).setValue(newQty);

return "✅ Restock "+code+" +"+qty;

}

}

return "❌ ไม่พบสินค้า";

}

function withdraw(code,qty,col){

const sheet=getSheet();
const data=sheet.getDataRange().getValues();

for(let i=1;i<data.length;i++){

if(data[i][0]==code){

let stock=data[i][4];

if(stock<qty){
return "❌ Stock ไม่พอ";
}

sheet.getRange(i+1,5).setValue(stock-qty);
sheet.getRange(i+1,col+1).setValue(data[i][col]+qty);

return "📦 เบิก "+code+" "+qty;

}

}

return "❌ ไม่พบสินค้า";

}

function allStock(){

const data=getSheet().getDataRange().getValues();

let msg="📦 ALL STOCK\n\n";

for(let i=1;i<data.length;i++){

msg+=data[i][0]+" | "+data[i][4]+"\n";

}

return msg;

}

function lowStock(){

const data=getSheet().getDataRange().getValues();

let msg="⚠️ LOW STOCK\n\n";

for(let i=1;i<data.length;i++){

if(data[i][4]<=LOW_STOCK_LIMIT){

msg+=data[i][0]+" | "+data[i][4]+"\n";

}

}

return msg;

}

function report(){

const data=getSheet().getDataRange().getValues();

let total=data.length-1;
let low=0;

for(let i=1;i<data.length;i++){

if(data[i][4]<=LOW_STOCK_LIMIT){
low++;
}

}

return `

📊 STOCK REPORT

Items : ${total}
Low Stock : ${low}`;

}

function searchStock(keyword){

const data=getSheet().getDataRange().getValues();
let msg="🔎 SEARCH RESULT\n\n";

for(let i=1;i<data.length;i++){

if(String(data[i][2]).toLowerCase().includes(keyword.toLowerCase())){

msg+=data[i][0]+" | "+data[i][4]+"\n";

}

}

return msg || "❌ ไม่พบข้อมูล";

}

function onEdit(e){

if(!e) return;

const oldValue = e.oldValue || "ไม่มีค่าเดิม";
const sheetName = "STOCK";
const sheet = e.source.getActiveSheet();

if(sheet.getName() !== sheetName) return;

const row = e.range.getRow();
const col = e.range.getColumn();
const cell = e.range.getA1Notation();
const value = e.range.getValue();

// ดึงชื่อ column header
const header = sheet.getRange(1,col).getValue();

const data = sheet.getRange(row,1,1,5).getValues()[0];

const id = data[0];
const desc = data[2];
const qty = data[4];

// msg = Output
let msg = `
 📦 มีการอัปเดตสต็อก

 🔖 ID: ${id}
 📦 Description : ${desc}

 ✏️ Edited : ${header}
 🔢 New Value : ${value}
 📊 ค่าเดิมก่อนแก้ไข : ${oldValue}

 📍 ตำแหน่งในชีต : ${cell}`;
 

// ============ENG VERSION===============
// let msg = `
// 📦 STOCK UPDATED

// 🔖 Item ID : ${id}
// 📦 Item Name : ${desc}

// ✏️ Updated Field : ${header}
// 🔢 New Value : ${value}
// 📊 Current Stock : ${qty}

// 📍 Sheet Location
// Row : ${row}
// Column : ${cell}`;

sendTelegram(CHAT_ID,msg); // Send message Func

}