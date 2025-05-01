/**
 * Centralized constants and reusable utilities for email filtering and processing
 */
function getPublicSheetData(sheetName) {
  var sheetUrl = getEnv('DATA_SHEET_URL'); // Ensure this returns a valid URL
  var ss = SpreadsheetApp.openByUrl(sheetUrl);

  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("Error: Sheet '" + sheetName + "' not found.");
    return [];
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log("No data found or only headers present.");
    return [];
  }

  let formattedData = data.slice(1).map(item => ({
    from: item[0] || "", 
    subject: item[1] || "", 
    plainBody: item[2] || ""
  }));

  return formattedData;
}

let GLOBAL_BAN_LIST = [
  { from: Session.getActiveUser().getEmail(), subject: "", plainBody: "" }
];

GLOBAL_BAN_LIST = GLOBAL_BAN_LIST.concat(getPublicSheetData("global"));

function exportSheetToFolder() {
  const sheetFile = SpreadsheetApp.getActiveSpreadsheet();
  const folderId = getEnv('FOLDER_ID'); // <-- Replace with your folder ID
  const folder = DriveApp.getFolderById(folderId);

  // Define export format: "application/pdf", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", etc.
  const exportMime = MimeType.PDF; // or MimeType.MICROSOFT_EXCEL

  const url = `https://www.googleapis.com/drive/v3/files/${sheetFile.getId()}/export?mimeType=${encodeURIComponent(exportMime)}`;

  const token = ScriptApp.getOAuthToken();

  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + token,
    },
    muteHttpExceptions: true,
  });

  // Create file in the target folder
  folder.createFile(response.getBlob()).setName(sheetFile.getName() + '_exported');
}

function getMergedBanList(localList = []) {
  return [...GLOBAL_BAN_LIST, ...localList];
}

function isBannedEmail(from, subject, plainBody, mergedList) {
  return mergedList.some(ban => from.toLowerCase().includes(ban.from.toLowerCase()) && subject.toLowerCase().includes(ban.subject.toLowerCase()) && plainBody.toLowerCase().includes(ban.plainBody.toLowerCase()));
}

function getEnv(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

function sendMsgToSlack(payload) {
  const slackWebhookUrl = getEnv('SLACK_WEBHOOK');
  UrlFetchApp.fetch(slackWebhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: payload
  });
}

function removeDuplicateMessages(messages) {
  const seen = new Set();
  return messages.filter(msg => {
    const identifier = `${msg.from}|${msg.subject}`;
    if (seen.has(identifier)) return false;
    seen.add(identifier);
    return true;
  });
}

function fetchMessagesFromThreads(threads) {
  return threads.flatMap(thread => thread.getMessages()).map(message => ({
    date: message.getDate(),
    from: message.getFrom(),
    subject: message.getSubject(),
    plainBody: message.getPlainBody().replace(/\n/g, ''),
    gmailMessage: message
  }));
}

function categorizeEasyApplyMessages(messages, rules) {
  const easyApply = [];
  const remaining = [];

  for (const msg of messages) {
    const matched = rules.some(rule =>
      msg.subject.toLowerCase().includes(rule.subject.toLowerCase()) &&
      msg.from.toLowerCase().includes(rule.from.toLowerCase())
    );
    if (matched) easyApply.push(msg);
    else remaining.push(msg);
  }

  return { easyApply, remaining };
}

function logMessagesToSheet(sheet, title, messages, countColumn = false) {
  sheet.appendRow([title]);
  sheet.getRange(sheet.getLastRow(), 1).setFontWeight("bold").setFontColor("blue");

  const startRow = sheet.getLastRow() + 1;

  const manualList = getPublicSheetData("manual");

  messages.forEach(msg => {
    const row = [msg.date, msg.from, msg.subject];
    if (countColumn) row.push(messages.length);
    if (title === "Manual" && manualList.some(rule => msg.from.toLowerCase().includes(rule.from.toLowerCase()) && msg.subject.toLowerCase().includes(rule.subject.toLowerCase()) && msg.plainBody.toLowerCase().includes(rule.plainBody.toLowerCase()))) {
      row.push(1)
    }
    sheet.appendRow(row);
    sheet.getRange(sheet.getLastRow(), 1).setFontWeight("normal").setFontColor("black");
  });

  const endRow = sheet.getLastRow();

  if (title === "Manual") {
    const sumFormula = `=SUM(D${startRow}:D${endRow})`;
    const sumRow = ["", "", "Total:", sumFormula];
    sheet.appendRow(sumRow);
    sheet.getRange(sheet.getLastRow(), 3, 1, 2).setFontWeight("bold");
  }
  sheet.appendRow([" "]); // Spacer
}

function isWeekend(dateStr) {
  const [month, day, year] = dateStr.split("/").map(Number);
  const date = new Date(year, month - 1, day); // month is 0-based in JS
  const dayOfWeek = date.getDay();
  return dayOfWeek === 0 || dayOfWeek === 6; // 0 = Sunday, 6 = Saturday
}

function checkEmailsAndNotifySlack() {
  const now = new Date();
  const time2 = Math.floor(now.getTime() / 1000);
  const offset = 1;
  const time1 = time2 - 60 - offset;
  const query = `newer:${time1} older:${time2} category:primary in:inbox is:unread`;

  const slackBanList = getPublicSheetData("slack");
  const mergedBanList = getMergedBanList(slackBanList);

  const threads = GmailApp.search(query);
  if (!threads.length) return Logger.log("No new emails found.");

  let messages = removeDuplicateMessages(fetchMessagesFromThreads(threads));

  messages.forEach(message => {
    if (!isBannedEmail(message.from, message.subject, message.plainBody, mergedBanList)) {
      const payload = JSON.stringify({
        text: `To: ${Session.getActiveUser().getEmail()}\nFrom: ${message.from}\nSubject: ${message.subject}`
      });
      sendMsgToSlack(payload);
    } else {
      message.gmailMessage.markRead();
    }
  });
}

function fetchEmailsByQuery(query) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const spreadsheetBanList = getPublicSheetData("spreadsheet")
  const mergedBanList = getMergedBanList(spreadsheetBanList);

  const easyApplyRules = getPublicSheetData("easy");

  const threads = GmailApp.search(query);
  if (!threads.length) return Logger.log("No new emails found.");

  let messages = removeDuplicateMessages(fetchMessagesFromThreads(threads));

  const { easyApply, remaining } = categorizeEasyApplyMessages(messages, easyApplyRules);

  logMessagesToSheet(sheet, "Easy", easyApply, true);
  logMessagesToSheet(sheet, "Manual", remaining.filter(msg => !isBannedEmail(msg.from, msg.subject, msg.plainBody, mergedBanList)));

  Logger.log("Emails successfully processed and recorded.");
}

function fetchEmailsDaily() {
  const today = new Date();
  const targetDate = `${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`;

  if (isWeekend(targetDate)) {
    Logger.log("It's a weekend!");
    return;
  }

  const [month, day, year] = targetDate.split("/").map(Number);

  const endET = new Date(`${month}/${day}/${year} ${today.getHours()}:${today.getMinutes()}:${today.getSeconds()} GMT-0400`);
  const startET = new Date(endET.getTime() - 20 * 60 * 60 * 1000);

  const time1 = Math.floor(startET.getTime() / 1000);
  const time2 = Math.floor(endET.getTime() / 1000);
  const query = `newer:${time1} older:${time2} category:primary in:inbox`;
  fetchEmailsByQuery(query);
}

function fetchEmailsForCertainDay() {
  const today = new Date();
  const targetDate = "04/30/2025"; // or dynamically generate

  if (isWeekend(targetDate)) {
    Logger.log("It's a weekend!");
    return;
  }

  const [month, day, year] = targetDate.split("/").map(Number);

  const endET = new Date(`${month}/${day}/${year} 18:${today.getMinutes()}:${today.getSeconds()} GMT-0400`);
  const startET = new Date(endET.getTime() - 20 * 60 * 60 * 1000);

  const time1 = Math.floor(startET.getTime() / 1000);
  const time2 = Math.floor(endET.getTime() / 1000);
  const query = `newer:${time1} older:${time2} category:primary in:inbox`;
  fetchEmailsByQuery(query);
}
