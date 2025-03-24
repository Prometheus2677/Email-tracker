function fetchEmailsDaily() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Define the target date (MM/DD/YYYY format)
    var targetDate = "02/12/2025"; // Change this to the desired date
    
    var [month, day, year] = targetDate.split("/").map(Number);
    
    // Create Date objects for 12 PM and 7 PM Eastern Time (ET)
    var startET = new Date(`${month}/${day}/${year} 12:00:00 GMT-0400`);
    var endET = new Date(`${month}/${day}/${year} 19:00:00 GMT-0400`);

    // Convert to timestamps in seconds for Gmail search
    var time1 = Math.floor(startET.getTime() / 1000);
    var time2 = Math.floor(endET.getTime() / 1000);

    var query = `newer:${time1} older:${time2} in:inbox`;

    var threads = GmailApp.search(query);
    var messages = threads.flatMap(thread => thread.getMessages());

    if (messages.length === 0) {
        Logger.log("No new emails found for the specified time range.");
        return;
    }

    var linkedinMsg = messages.filter(msg => 
        msg.getSubject().toLowerCase().includes("your application was sent to") && 
        msg.getFrom().toLowerCase().includes("linkedin")
    );

    // Append email details to Google Sheet
    messages.forEach(msg => {
        sheet.appendRow([
            msg.getDate(),
            msg.getFrom(),
            msg.getSubject(),
            messages.length,
            linkedinMsg.length
        ]);
    });

    Logger.log("Emails inserted successfully.");
}
