/**
 * Centralized constants and reusable utilities for email filtering and processing
 */
const GLOBAL_BAN_LIST = [
  { from: "jobs-noreply@linkedin.com", subject: "your application was sent to" },
  { from: "jobs-noreply@linkedin.com", subject: "Your application to" },
  { from: "LinkedIn <jobs-noreply@linkedin.com>", subject: "Your application was viewed by" },
  { from: "LinkedIn Job Alerts <jobalerts-noreply@linkedin.com>", subject: "" },
  { from: "applyonline@dice.com", subject: "Application for Dice Job" },
  { from: "Indeed Apply <indeedapply@indeed.com>", subject: "Indeed Application:" },
  { from: "Discord <noreply@discord.com>", subject: "" },
  { from: "Google <no-reply@accounts.google.com>", subject: "Security alert" },
  { from: "", subject: "be the first to apply!" },
  { from: "", subject: "Your job alert for" },
  { from: "JobLeads <mailer@jobleads.com>", subject: "new jobs match your job search" },
  { from: "Amy at Adzuna <no-reply@adzuna.com>", subject: "You could be a great fit with" },
  { from: "Amy at Adzuna <no-reply@adzuna.com>", subject: "is hiring and more new" },
  { from: "UKG Notifications <noreply@notifications.ultipro.com>", subject: "You have a new password" },
  { from: "Jooble <subscribe@jooble.org>", subject: "more new jobs" },
  { from: "Glassdoor Jobs <noreply@glassdoor.com>", subject: "Apply Now." },
  { from: "Glassdoor Jobs <noreply@glassdoor.com>", subject: "you would be a great fit!" },
  { from: "Glassdoor Jobs <noreply@glassdoor.com>", subject: "job search going?" },
  { from: "Glassdoor <noreply@glassdoor.com>", subject: "This week's employee reviews and more" },
  { from: "Jooble <subscribe@jooble.org>", subject: "more new jobs" },
  { from: "<updates@pmail.jobcase.com>", subject: "has open roles that may interest you" },
  { from: "Lensa Aggregated <aggregated@lensa.com>", subject: "jobs open" },
  { from: "JobLeads <mailer@jobleads.com>", subject: "new jobs match your job search" },
  // { from: "", subject: "" },
  { from: `Proofreader (via JH)" <jobs@pmail.jobhat.com>`, subject: "Apply now." },
  { from: "<updates@pmail.jobcase.com>", subject: "has open roles that may interest you" },
  { from: "ZipRecruiter <alerts@ziprecruiter.com>", subject: "Today's jobs chosen for you" },
  { from: "", subject: "Your account has been created" },
  { from: Session.getActiveUser().getEmail(), subject: "" },
];

function getMergedBanList(localList = []) {
  return [...GLOBAL_BAN_LIST, ...localList];
}

function isBannedEmail(from, subject, mergedList) {
  return mergedList.some(ban => from.includes(ban.from) && subject.includes(ban.subject));
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
    const identifier = `${msg.sender}|${msg.subject}`;
    if (seen.has(identifier)) return false;
    seen.add(identifier);
    return true;
  });
}

function fetchMessagesFromThreads(threads) {
  return threads.flatMap(thread => thread.getMessages()).map(message => ({
    date: message.getDate(),
    sender: message.getFrom(),
    subject: message.getSubject(),
    gmailMessage: message
  }));
}

function categorizeEasyApplyMessages(messages, rules) {
  const easyApply = [];
  const remaining = [];

  for (const msg of messages) {
    const matched = rules.some(rule =>
      msg.subject.toLowerCase().includes(rule.subject) &&
      msg.sender.toLowerCase().includes(rule.sender)
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

  messages.forEach(msg => {
    const row = [msg.date, msg.sender, msg.subject];
    if (countColumn) row.push(messages.length);
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
  const time1 = time2 - 60;
  const query = `newer:${time1} older:${time2} category:primary in:inbox is:unread`;

  const slackBanList = [
    { from: "", subject: "Thank you for applying" },
    { from: "", subject: "Thanks for applying" },
    { from: "", subject: "Thank you for your application to" },
  ];
  const mergedBanList = getMergedBanList(slackBanList);

  const threads = GmailApp.search(query);
  if (!threads.length) return Logger.log("No new emails found.");

  let messages = removeDuplicateMessages(fetchMessagesFromThreads(threads));

  messages.forEach(message => {
    if (!isBannedEmail(message.sender, message.subject, mergedBanList)) {
      const payload = JSON.stringify({
        text: `To: ${Session.getActiveUser().getEmail()}\nFrom: ${message.sender}\nSubject: ${message.subject}`
      });
      sendMsgToSlack(payload);
    } else {
      message.gmailMessage.markRead();
    }
  });
}

function fetchEmailsDaily() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const today = new Date();
  // const targetDate = "03/27/2025"; // or dynamically generate
  const targetDate = `${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`;

  if (isWeekend(targetDate)) {
    Logger.log("It's a weekend!");
    return;
  }

  const [month, day, year] = targetDate.split("/").map(Number);

  const startET = new Date(`${month}/${day}/${year} 12:00:00 GMT-0400`);
  const endET = new Date(`${month}/${day}/${year} 19:00:00 GMT-0400`);
  const time1 = Math.floor(startET.getTime() / 1000);
  const time2 = Math.floor(endET.getTime() / 1000);
  const query = `newer:${time1} older:${time2} category:primary in:inbox`;

  const spreadsheetBanList = [];
  const mergedBanList = getMergedBanList(spreadsheetBanList);

  const easyApplyRules = [
    { subject: "your application was sent to", sender: "linkedin" },
    { subject: "application for dice job", sender: "applyonline@dice.com" },
    { subject: "indeed application:", sender: "indeed apply <indeedapply@indeed.com>" }
  ];

  const threads = GmailApp.search(query);
  if (!threads.length) return Logger.log("No new emails found.");

  let messages = removeDuplicateMessages(fetchMessagesFromThreads(threads));

  const { easyApply, remaining } = categorizeEasyApplyMessages(messages, easyApplyRules);

  logMessagesToSheet(sheet, "Easy", easyApply, true);
  logMessagesToSheet(sheet, "Manual", remaining.filter(msg => !isBannedEmail(msg.sender, msg.subject, mergedBanList)));

  Logger.log("Emails successfully processed and recorded.");
}
