function fetchEmailsDaily() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var month = new Date().getMonth();
    var date = new Date().getDate();
    var year = new Date().getFullYear();
    var time1 = new Date(year, month, date, 0, 0, 0).getTime();
    var time2 = time1 - 86400000;
    var query = "newer:" + time2/1000 + " older:"  + time1/1000 + " in:inbox";
    
    var threads = GmailApp.search(query);
    var messages = threads.flatMap(thread => thread.getMessages());

    if (messages.length === 0) {
        Logger.log("No new emails found for today.");
        return;
    }

    // Append email details to Google Sheet
    messages.forEach(msg => {
        sheet.appendRow([
            msg.getDate(),
            msg.getFrom(),
            msg.getSubject(),
            msg.getPlainBody().substring(0, 500) // Limiting body size to 500 chars
        ]);
    });

    Logger.log("Emails inserted successfully.");
}
