function getEnv(key) {
    return PropertiesService.getScriptProperties().getProperty(key);
}

function sendMsgToSlack(payload) {
    var slackWebhookUrl = getEnv('SLACK_WEBHOOK');
    UrlFetchApp.fetch(slackWebhookUrl, {
        method: 'post',
        contentType: 'application/json',
        payload: payload
    });
}

function checkEmailsAndNotifySlack() {
    var now = new Date();
    var time2 = Math.floor(now.getTime() / 1000); // current time in seconds
    var time1 = time2 - (15 * 60); // 15 minutes ago

    var query = `newer:${time1} older:${time2} category:primary in:inbox is:unread`;
    var threads = GmailApp.search(query);
    for (var i = 0; i < threads.length; i++) {
      var messages = threads[i].getMessages();
      for (var j = 0; j < messages.length; j++) {
        // var body = messages[j].getPlainBody();
        var subject = messages[j].getSubject();
        var from = messages[j].getFrom();

        var banList = [
            {from: "jobs-noreply@linkedin.com", subject: "Your application was sent to"},
            {from: "jobs-noreply@linkedin.com", subject: "Your application to"},
            {from: "applyonline@dice.com", subject: "application for dice job"},
        ]
        var isBanned = banList.some(function(banItem) {
            return from.includes(banItem.from) && subject.includes(banItem.subject);
        });
    
        if (!isBanned) {
            var payload = JSON.stringify({
                text: `From: ${from}\nSubject: ${subject}`
            });
    
            sendMsgToSlack(payload)
        }
      }
    }
}

function removeDuplicateMessages(messages) {
    let uniqueMessages = [];
    let seen = new Set();

    messages.forEach(message => {
        let identifier = message.sender + '|' + message.subject;
        if (!seen.has(identifier)) {
            seen.add(identifier);
            uniqueMessages.push(message);
        }
    });

    return uniqueMessages;
}

function fetchEmailsDaily() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    let today = new Date();
    // var targetDate = "03/21/2025";
    var targetDate = (today.getMonth() + 1) + "/" + today.getDate() + "/" + today.getFullYear();
    var [month, day, year] = targetDate.split("/").map(Number);

    var startET = new Date(`${month}/${day}/${year} 12:00:00 GMT-0400`);
    var endET = new Date(`${month}/${day}/${year} 19:00:00 GMT-0400`);

    var time1 = Math.floor(startET.getTime() / 1000);
    var time2 = Math.floor(endET.getTime() / 1000);

    var query = `newer:${time1} older:${time2} category:primary in:inbox`;
    var threads = GmailApp.search(query);
    var messages = threads.flatMap(thread => thread.getMessages());

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

    messages = messages.filter(msg => {
        var subject = msg.subject.toLowerCase();
        var sender = msg.sender.toLowerCase();
        
        var isNotApplication = !subject.includes("your application to");
        var isFromLinkedIn = sender.includes("jobs-noreply@linkedin.com");
        
        return (isNotApplication && isFromLinkedIn) || !isFromLinkedIn;
    });

    var easyApply = [];
    
    for (let i = 0; i < messages.length; i++) {
        let msg = messages[i];
        
        if ((msg.subject.toLowerCase().includes("your application was sent to") && 
            msg.sender.toLowerCase().includes("linkedin")) || 
            (msg.subject.toLowerCase().includes("application for dice job") && 
            msg.sender.toLowerCase().includes("applyonline@dice.com"))) {
            easyApply.push(msg);
            messages.splice(i, 1);
            i--;
        }
    }

    easyApply.sort((a, b) => a.sender.localeCompare(b.sender) || a.subject.localeCompare(b.subject));

    // Append "Easy" section
    sheet.appendRow(["Easy"]);
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1).setFontWeight("bold").setFontColor("blue");

    easyApply.forEach(msg => {
        sheet.appendRow([
            msg.date,
            msg.sender,
            msg.subject,
            easyApply.length
        ]);
    });

    sheet.appendRow([" "]); // Blank line fix

    // Append "Manual" section with bold and blue formatting
    sheet.appendRow(["Manual"]);
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1).setFontWeight("bold").setFontColor("blue");

    messages.forEach(msg => {
        sheet.appendRow([
            msg.date,
            msg.sender,
            msg.subject
        ]);
    });

    sheet.appendRow([" "]); // Blank line fix
    Logger.log("Emails inserted successfully.");
}