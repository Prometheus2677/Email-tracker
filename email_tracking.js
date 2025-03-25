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

    // Define the target date (MM/DD/YYYY format)
    // var targetDate = "03/21/2025"; // Change this to the desired date
    var targetDate = (today.getMonth() + 1) + "/" + today.getDate() + "/" + today.getFullYear();
    
    var [month, day, year] = targetDate.split("/").map(Number);
    
    // Create Date objects for 12 PM and 7 PM Eastern Time (ET)
    var startET = new Date(`${month}/${day}/${year} 12:00:00 GMT-0400`);
    var endET = new Date(`${month}/${day}/${year} 19:00:00 GMT-0400`);

    // Convert to timestamps in seconds for Gmail search
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
        
        if (msg.subject.toLowerCase().includes("your application was sent to") && 
            msg.sender.toLowerCase().includes("linkedin")) {
            easyApply.push(msg);
            messages.splice(i, 1);  // Removes the message from the array
            i--;  // Decrement the index to account for the shift in array after removal
        }
    }

    // Sort messages by sender alphabetically
    easyApply.sort((a, b) => {
        let senderComparison = a.sender.localeCompare(b.sender);
        return senderComparison !== 0 ? senderComparison : a.subject.localeCompare(b.subject);
    });

    sheet.appendRow([
        "Easy"
    ]);
    // Append email details to Google Sheet
    easyApply.forEach(msg => {
        sheet.appendRow([
            msg.date,
            msg.sender,
            msg.subject,
            easyApply.length
        ]);
    });

    sheet.appendRow([
        ""
    ]);
    sheet.appendRow([
        "Manual"
    ]);
    // Append email details to Google Sheet
    messages.forEach(msg => {
        sheet.appendRow([
            msg.date,
            msg.sender,
            msg.subject,
        ]);
    });

    sheet.appendRow([
        ""
    ]);
    Logger.log("Emails inserted successfully.");
}