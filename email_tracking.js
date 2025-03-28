/**
 * Global ban list shared across all use cases
 */
const GLOBAL_BAN_LIST = [
    { from: "jobs-noreply@linkedin.com", subject: "your application was sent to" },
    { from: "jobs-noreply@linkedin.com", subject: "Your application to" },
    { from: "LinkedIn <jobs-noreply@linkedin.com>", subject: "Your application was viewed by" },
    { from: "applyonline@dice.com", subject: "Application for Dice Job" },
    { from: "Indeed Apply <indeedapply@indeed.com>", subject: "Indeed Application:" },
    { from: "Discord <noreply@discord.com>", subject: "" },
    { from: "Google <no-reply@accounts.google.com>", subject: "Security alert" },
    { from: "", subject: "be the first to apply!" },
    { from: "", subject: "Your job alert for" },
    { from: "LinkedIn Job Alerts <jobalerts-noreply@linkedin.com>", subject: "" },
    { from: "Glassdoor Jobs <noreply@glassdoor.com>", subject: "Apply Now." },
    { from: "Glassdoor Jobs <noreply@glassdoor.com>", subject: "you would be a great fit!" },
    { from: "ZipRecruiter <alerts@ziprecruiter.com>", subject: "Today's jobs chosen for you" },
    { from: "", subject: "Your account has been created" },
  ];
  
  /**
   * Merges global and local ban lists
   * @param {Array} localList - Additional ban list entries for the specific context
   * @returns {Array}
   */
  function getMergedBanList(localList = []) {
    return [...GLOBAL_BAN_LIST, ...localList];
  }
  
  /**
   * Checks if an email is banned based on from/subject and merged list
   * @param {string} from 
   * @param {string} subject 
   * @param {Array} mergedList 
   * @returns {boolean}
   */
  function isBannedEmail(from, subject, mergedList) {
    return mergedList.some(ban =>
      from.includes(ban.from) && subject.includes(ban.subject)
    );
  }
  
  /**
   * Retrieves environment variables from script properties
   * @param {string} key 
   * @returns {string|null}
   */
  function getEnv(key) {
    return PropertiesService.getScriptProperties().getProperty(key);
  }
  
  /**
   * Sends a payload to the configured Slack webhook
   * @param {string} payload - A JSON-formatted string
   */
  function sendMsgToSlack(payload) {
    const slackWebhookUrl = getEnv('SLACK_WEBHOOK');
    UrlFetchApp.fetch(slackWebhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: payload
    });
  }
  
  /**
   * Checks for unread emails in the last minute and notifies Slack if not banned
   */
  function checkEmailsAndNotifySlack() {
    const now = new Date();
    const time2 = Math.floor(now.getTime() / 1000);
    const time1 = time2 - 60;
  
    const query = `newer:${time1} older:${time2} category:primary in:inbox is:unread`;
    const threads = GmailApp.search(query);
  
    const slackBanList = [
        { from: "", subject: "Thank you for applying" },
        { from: "", subject: "Thanks for applying" },
    ];
  
    const mergedBanList = getMergedBanList(slackBanList);
  
    let messages = threads.flatMap(thread => thread.getMessages());
  
    if (messages.length === 0) {
      Logger.log("No new emails found for the specified time range.");
      return;
    }
  
    messages = messages.map(msg => ({
      date: msg.getDate(),
      sender: msg.getFrom(),
      subject: msg.getSubject()
    }));
  
    messages = removeDuplicateMessages(messages);

    messages.forEach(message => {
        if (!isBannedEmail(message.sender, message.subject, mergedBanList)) {
          const payload = JSON.stringify({
            text: `To: ${Session.getActiveUser().getEmail()}\nFrom: ${message.sender}\nSubject: ${message.subject}`
          });
          sendMsgToSlack(payload);
        }
    });
  }
  
  /**
   * Removes duplicate messages based on sender and subject
   * @param {Array} messages 
   * @returns {Array}
   */
  function removeDuplicateMessages(messages) {
    const seen = new Set();
    return messages.filter(msg => {
      const identifier = `${msg.sender}|${msg.subject}`;
      if (seen.has(identifier)) return false;
      seen.add(identifier);
      return true;
    });
  }
  
  /**
   * Fetches emails for a specific day and logs them to a spreadsheet,
   * separating "Easy Apply" and "Manual" emails with custom and global filtering
   */
  function fetchEmailsDaily() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    // const targetDate = "03/27/2025"; // or dynamically generate
    const targetDate = (today.getMonth() + 1) + "/" + today.getDate() + "/" + today.getFullYear();
  
    const [month, day, year] = targetDate.split("/").map(Number);
    const startET = new Date(`${month}/${day}/${year} 12:00:00 GMT-0400`);
    const endET = new Date(`${month}/${day}/${year} 19:00:00 GMT-0400`);
  
    const time1 = Math.floor(startET.getTime() / 1000);
    const time2 = Math.floor(endET.getTime() / 1000);
  
    const query = `newer:${time1} older:${time2} category:primary in:inbox`;
    const threads = GmailApp.search(query);
    let messages = threads.flatMap(thread => thread.getMessages());
  
    if (messages.length === 0) {
      Logger.log("No new emails found for the specified time range.");
      return;
    }
  
    messages = messages.map(msg => ({
      date: msg.getDate(),
      sender: msg.getFrom(),
      subject: msg.getSubject()
    }));
  
    messages = removeDuplicateMessages(messages);
  
    const spreadsheetBanList = [
    ];
  
    const mergedBanList = getMergedBanList(spreadsheetBanList);
  
    const easyApplyRules = [
      { subject: "your application was sent to", sender: "linkedin" },
      { subject: "application for dice job", sender: "applyonline@dice.com" },
      { subject: "indeed application:", sender: "indeed apply <indeedapply@indeed.com>" }
    ];
  
    const easyApply = [];
  
    for (let i = 0; i < messages.length; i++) {
      const msg = messages[i];
      const matched = easyApplyRules.some(rule =>
        msg.subject.toLowerCase().includes(rule.subject) &&
        msg.sender.toLowerCase().includes(rule.sender)
      );
  
      if (matched) {
        easyApply.push(msg);
        messages.splice(i, 1);
        i--;
      }
    }
  
    easyApply.sort((a, b) =>
      a.sender.localeCompare(b.sender) || a.subject.localeCompare(b.subject)
    );
  
    // Easy Apply Section
    sheet.appendRow(["Easy"]);
    sheet.getRange(sheet.getLastRow(), 1).setFontWeight("bold").setFontColor("blue");
    easyApply.forEach(msg => {
      sheet.appendRow([msg.date, msg.sender, msg.subject, easyApply.length]);
    });
  
    sheet.appendRow([""]); // Spacer
  
    // Manual Review Section
    sheet.appendRow(["Manual"]);
    sheet.getRange(sheet.getLastRow(), 1).setFontWeight("bold").setFontColor("blue");
  
    messages.forEach(msg => {
      if (!isBannedEmail(msg.sender, msg.subject, mergedBanList)) {
        sheet.appendRow([msg.date, msg.sender, msg.subject]);
      }
    });
  
    sheet.appendRow([""]); // Final spacer
    Logger.log("Emails successfully processed and recorded.");
  }
  