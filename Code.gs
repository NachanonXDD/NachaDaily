// ==========================================
// NACHA DAILY - GOOGLE APPS SCRIPT BACKEND
// ==========================================

// 1. ใส่ Token และ Chat ID ของคุณที่นี่
const TELEGRAM_BOT_TOKEN = "YOUR_BOT_TOKEN_HERE"; 
const TELEGRAM_CHAT_ID = "YOUR_CHAT_ID_HERE";

// ชื่อชีตที่จะใช้เก็บข้อมูล
const SHEET_NAME = "Tasks";

// ฟังก์ชันนี้กดปุ่ม [เรียกใช้] (Run) 1 ครั้งแรกเพื่อตั้งค่าเริ่มต้น
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["ID", "Title", "Type", "Priority", "Date", "Time", "Completed"]);
    // Freeze header
    sheet.setFrozenRows(1);
    sheet.getRange("A1:G1").setFontWeight("bold");
  }
  
  createTimeDrivenTriggers();
  sendTelegramNotification("✅ Nacha Daily Backend ติดตั้งสำเร็จเรียบร้อยแล้ว!");
}

// ==========================================
// API ENDPOINTS สำหรับให้เว็บเรียกใช้
// ==========================================

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  
  const tasks = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    tasks.push({
      id: row[0], title: row[1], type: row[2], 
      priority: row[3], date: row[4], time: row[5], 
      completed: row[6] === true || row[6] === "true" || row[6] === "TRUE"
    });
  }
  return ContentService.createTextOutput(JSON.stringify(tasks)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const requestData = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    // เคลียร์ข้อมูลเก่าออกให้หมดแล้วเซฟใหม่
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).clearContent();
    }
    
    if (requestData.length > 0) {
      const rows = requestData.map(task => [
        task.id, task.title, task.type, task.priority, 
        task.date, task.time, task.completed
      ]);
      sheet.getRange(2, 1, rows.length, 7).setValues(rows);
    }
    return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({error: error.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// TELEGRAM NOTIFICATIONS
// ==========================================

function sendTelegramNotification(message) {
  if (TELEGRAM_BOT_TOKEN === "YOUR_BOT_TOKEN_HERE" || !TELEGRAM_BOT_TOKEN) return;
  
  const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;
  const payload = {
    chat_id: TELEGRAM_CHAT_ID,
    text: message,
    parse_mode: "HTML"
  };
  
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  UrlFetchApp.fetch(url, options);
}

// ==========================================
// TRIGGERS (ตั้งเวลาแจ้งเตือน)
// ==========================================

function createTimeDrivenTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  // แจ้งเตือนเช้า 8:00
  ScriptApp.newTrigger("sendMorningBrief").timeBased().atHour(8).nearMinute(0).everyDays(1).create();
  // เช็คงานล่วงหน้าทุกๆ 15 นาที
  ScriptApp.newTrigger("checkUpcomingTasks").timeBased().everyMinutes(15).create();
  // สรุปงานตอนสิ้นวัน 23:00
  ScriptApp.newTrigger("sendDailySummary").timeBased().atHour(23).nearMinute(0).everyDays(1).create();
}

function getTodayString() {
  const today = new Date();
  return Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function sendMorningBrief() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if(!sheet) return;
  const data = sheet.getDataRange().getValues();
  
  let taskCount = 0;
  let msg = "🌅 <b>Good Morning! งานที่ต้องทำในวันนี้:</b>\n\n";
  const todayStr = getTodayString();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const type = row[2];
    const date = row[4];
    const completed = row[6];
    
    // ไม่เอางาน Routine, ไม่เอางานที่เสร็จแล้ว
    if (type === "task" && (completed === false || completed === "false") && (date === todayStr || date === "")) {
      const priority = row[3] === "urgent" ? "🔥" : "📌";
      const timeStr = row[5] ? ` (${row[5]})` : "";
      msg += `${priority} <b>${row[1]}</b>${timeStr}\n`;
      taskCount++;
    }
  }
  
  if (taskCount > 0) sendTelegramNotification(msg);
}

function checkUpcomingTasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if(!sheet) return;
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const todayStr = getTodayString();
  
  const props = PropertiesService.getScriptProperties();
  let alerted = props.getProperty('alertedTasks');
  alerted = alerted ? JSON.parse(alerted) : {};
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = row[0];
    const title = row[1];
    const type = row[2];
    const date = row[4];
    const time = row[5];
    const completed = row[6] === true || row[6] === "true" || row[6] === "TRUE";
    
    if (type !== "task" || completed || !time || (date && date !== todayStr)) continue;
    
    const [hours, minutes] = time.split(':');
    const taskDate = new Date();
    taskDate.setHours(parseInt(hours, 10));
    taskDate.setMinutes(parseInt(minutes, 10));
    taskDate.setSeconds(0);
    
    const diffMs = taskDate.getTime() - now.getTime();
    const diffMins = Math.floor(diffMs / 60000);
    
    if (!alerted[id]) alerted[id] = {};
    
    // 1 Hour Before
    if (diffMins <= 60 && diffMins > 30 && !alerted[id]['60m']) {
      sendTelegramNotification(`⏳ <b>กำลังจะถึงใน 1 ชั่วโมง:</b>\n📌 ${title}\n⏰ ${time}`);
      alerted[id]['60m'] = true;
    }
    // 30 Mins Before
    if (diffMins <= 30 && diffMins > 0 && !alerted[id]['30m']) {
      sendTelegramNotification(`⚠️ <b>อีก 30 นาที:</b>\n📌 ${title}\n⏰ ${time}`);
      alerted[id]['30m'] = true;
    }
    // Exact Time (0 mins)
    if (diffMins <= 0 && diffMins > -15 && !alerted[id]['0m']) {
      sendTelegramNotification(`🚨 <b>ถึงเวลาแล้ว!</b>\n🔥 <b>${title}</b>\n⏰ กำหนดเวลา: ${time}`);
      alerted[id]['0m'] = true;
    }
  }
  
  // ล้างการแจ้งเตือนตอนเที่ยงคืน
  if (now.getHours() === 0 && now.getMinutes() < 15) {
     props.setProperty('alertedTasks', JSON.stringify({}));
  } else {
     props.setProperty('alertedTasks', JSON.stringify(alerted));
  }
}

function sendDailySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if(!sheet) return;
  const data = sheet.getDataRange().getValues();
  
  let totalTasks = 0;
  let completedTasks = 0;
  let pendingMsg = "";
  const todayStr = getTodayString();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const type = row[2];
    const title = row[1];
    const date = row[4];
    const completed = row[6] === true || row[6] === "true" || row[6] === "TRUE";
    
    if (type === "task" && (date === todayStr || date === "")) {
      totalTasks++;
      if (completed) {
        completedTasks++;
      } else {
        pendingMsg += `- ${title}\n`;
      }
    }
  }
  
  if (totalTasks === 0) return;
  const percent = Math.round((completedTasks / totalTasks) * 100);
  let msg = `📊 <b>สรุปผลประจำวัน</b>\n\n`;
  msg += `ความสำเร็จ: <b>${percent}%</b> (${completedTasks}/${totalTasks} งาน)\n\n`;
  
  if (percent === 100) {
    msg += `🏆 ยอดเยี่ยมมาก! วันนี้คุณจัดการงานได้หมดเกลี้ยง`;
  } else {
    msg += `📉 <b>งานที่ยังทำไม่เสร็จ:</b>\n${pendingMsg}\nพักผ่อนให้เต็มที่แล้วมาลุยต่อพรุ่งนี้นะครับ!`;
  }
  
  sendTelegramNotification(msg);
}
