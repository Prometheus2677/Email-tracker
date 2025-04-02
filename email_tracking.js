/**
 * Centralized constants and reusable utilities for email filtering and processing
 */
const GLOBAL_BAN_LIST = [
  { from: "jobs-noreply@linkedin.com", subject: "your application was sent to", plainBody: "" },
  { from: "jobs-noreply@linkedin.com", subject: "Your application to", plainBody: "" },
  { from: "LinkedIn <jobs-noreply@linkedin.com>", subject: "Your application was viewed by", plainBody: "" },
  { from: "LinkedIn Job Alerts <jobalerts-noreply@linkedin.com>", subject: "", plainBody: "" },
  { from: "LinkedIn <jobs-noreply@linkedin.com>", subject: "you have new application updates this week", plainBody: "" },
  { from: "applyonline@dice.com", subject: "Application for Dice Job", plainBody: "" },
  { from: "Dice <dice@connect.dice.com>", subject: "and other open positions!", plainBody: "" },
  { from: "Dice <dice@connect.dice.com>", subject: "Your IntelliSearch Alert:", plainBody: "" },
  { from: "Indeed Apply <indeedapply@indeed.com>", subject: "Indeed Application:", plainBody: "" },
  { from: "Indeed <donotreply@match.indeed.com>", subject: "more new job", plainBody: "" },
  { from: "Discord <noreply@discord.com>", subject: "", plainBody: "" },
  { from: "Google <no-reply@accounts.google.com>", subject: "Security alert", plainBody: "" },
  { from: "", subject: "be the first to apply!", plainBody: "" },
  { from: "<newsletters-noreply@linkedin.com>", subject: "", plainBody: "" },
  { from: "", subject: "Your job alert for", plainBody: "" },
  { from: "JobLeads <mailer@jobleads.com>", subject: "new jobs match your job search", plainBody: "" },
  { from: "Amy at Adzuna <no-reply@adzuna.com>", subject: "You could be a great fit with", plainBody: "" },
  { from: "Amy at Adzuna <no-reply@adzuna.com>", subject: "is hiring and more new", plainBody: "" },
  { from: "Amy at Adzuna <no-reply@adzuna.com>", subject: "is looking for", plainBody: "" },
  { from: "Amy at Adzuna <no-reply@adzuna.com>", subject: "are looking for", plainBody: "" },
  { from: "Amy at Adzuna <no-reply@adzuna.com>", subject: "vacancies for you", plainBody: "" },
  { from: "UKG Notifications <noreply@notifications.ultipro.com>", subject: "You have a new password", plainBody: "" },
  { from: "Jooble <subscribe@jooble.org>", subject: "more new jobs", plainBody: "" },
  { from: "Glassdoor Jobs <noreply@glassdoor.com>", subject: "Apply Now.", plainBody: "" },
  { from: "Glassdoor Jobs <noreply@glassdoor.com>", subject: "you would be a great fit!", plainBody: "" },
  { from: "Glassdoor Jobs <noreply@glassdoor.com>", subject: "job search going?", plainBody: "" },
  { from: "Glassdoor <noreply@glassdoor.com>", subject: "This week's employee reviews and more", plainBody: "" },
  { from: "Jooble <subscribe@jooble.org>", subject: "more new jobs", plainBody: "" },
  { from: "<updates@pmail.jobcase.com>", subject: "has open roles that may interest you", plainBody: "" },
  { from: "Lensa Aggregated <aggregated@lensa.com>", subject: "jobs open", plainBody: "" },
  { from: "<lensa24@lensa.com>", subject: "Be the first to apply to", plainBody: "" },
  { from: "<lensa24@lensa.com>", subject: "jobs posted in the last 24 hours", plainBody: "" },
  { from: "JobLeads <mailer@jobleads.com>", subject: "new jobs match your job search", plainBody: "" },
  { from: "", subject: "has been successfully submitted", plainBody: "" },
  { from: "<team@hi.wellfound.com>", subject: "new jobs you'd be a great fit for", plainBody: "" },
  { from: "<updates@pmail.jobcase.com>", subject: "and other job searches for you", plainBody: "" },
  { from: "<eric.beck@email.jobleads.com>", subject: "matches your personal job search", plainBody: "" },
  { from: "<noreply@message.get.it>", subject: "Job Matches @ Get.It", plainBody: "" },
  { from: "<jobs@umail.texasjobdepartment.com>", subject: "Apply now.", plainBody: "" },
  { from: "<yashvant.t@highbrow-tech.com>", subject: "Job Opportunity :", plainBody: "" },
  { from: "", subject: "", plainBody: "unfortunately we have decided not to consider you further for this position." },
  { from: " <alerts@ziprecruiter.com>", subject: "has an open position", plainBody: "" },
  { from: "", subject: "Job Opportunity", plainBody: "" },
  { from: "episerveri-jobnotification@noreply55.jobs2web.com", subject: "New jobs posted from careers.optimizely.com", plainBody: "" },
  { from: "", subject: "(ONSITE)", plainBody: "" },
  { from: "<donotreply@job-announcements.com>", subject: "New positions at", plainBody: "" },
  { from: "", subject: "", plainBody: "we have decided to move forward with candidates whose qualifications more closely meet our needs at this time." },
  // { from: "", subject: "", plainBody: "" },
  { from: `Proofreader (via JH)" <jobs@pmail.jobhat.com>`, subject: "Apply now.", plainBody: "" },
  { from: "<updates@pmail.jobcase.com>", subject: "has open roles that may interest you", plainBody: "" },
  { from: "ZipRecruiter <alerts@ziprecruiter.com>", subject: "Today's jobs chosen for you", plainBody: "" },
  { from: "", subject: "Your account has been created", plainBody: "" },
  { from: "", subject: "", plainBody: "After careful consideration, we have decided not to move forward with your application." },
  { from: Session.getActiveUser().getEmail(), subject: "", plainBody: "" },
];

function getMergedBanList(localList = []) {
  return [...GLOBAL_BAN_LIST, ...localList];
}

function isBannedEmail(from, subject, plainBody, mergedList) {
  return mergedList.some(ban => from.includes(ban.from) && subject.includes(ban.subject) && plainBody.includes(ban.plainBody));
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
    plainBody: message.getPlainBody(),
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

  const manualList = [
    { from: "", subject: "Thank you for applying", plainBody: "" },
    { from: "", subject: "Thanks for applying", plainBody: "" },
    { from: "", subject: "Thank you for your application to", plainBody: "" },
  ];

  messages.forEach(msg => {
    const row = [msg.date, msg.sender, msg.subject];
    if (countColumn) row.push(messages.length);
    if (title === "Manual" && manualList.some(ban => msg.sender.includes(ban.from) && msg.subject.includes(ban.subject) && msg.plainBody.includes(ban.plainBody))) {
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
  const time1 = time2 - 62;
  const query = `newer:${time1} older:${time2} category:primary in:inbox is:unread`;

  const slackBanList = [
    { from: "", subject: "Thank you for applying", plainBody: "" },
    { from: "", subject: "Thanks for applying", plainBody: "" },
    { from: "", subject: "Thank you for your application to", plainBody: "" },
  ];
  const mergedBanList = getMergedBanList(slackBanList);

  const threads = GmailApp.search(query);
  if (!threads.length) return Logger.log("No new emails found.");

  let messages = removeDuplicateMessages(fetchMessagesFromThreads(threads));

  messages.forEach(message => {
    if (!isBannedEmail(message.sender, message.subject, message.plainBody, mergedBanList)) {
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

  const startET = new Date(`${month}/${day}/${year} 00:00:00 GMT-0400`);
  const endET = new Date(`${month}/${day}/${year} 23:59:59 GMT-0400`);
  const time1 = Math.floor(startET.getTime() / 1000);
  const time2 = Math.floor(endET.getTime() / 1000);
  const query = `newer:${time1} older:${time2} category:primary in:inbox`;

  const spreadsheetBanList = [
    { from: "<messaging-digest-noreply@linkedin.com>", subject: "just messaged you", plainBody: "" },
    { from: "", subject: "Invitation from an unknown sender: Interview with Prompt", plainBody: "" },
    // { from: "", subject: "", plainBody: "" },
  ];
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
  logMessagesToSheet(sheet, "Manual", remaining.filter(msg => !isBannedEmail(msg.sender, msg.subject, msg.plainBody, mergedBanList)));

  Logger.log("Emails successfully processed and recorded.");
}
