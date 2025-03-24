var getFormattedDate = function (timestamp) {
    var date = new Date(timestamp);

    // Format the date to "HH:mm:ss DD/MM/YYYY"
    var formattedDate = date.toLocaleString("en-GB", { 
        hour: '2-digit', 
        minute: '2-digit', 
        second: '2-digit', 
        day: '2-digit', 
        month: '2-digit', 
        year: 'numeric',
        hour12: false // Use 24-hour format
    });

    return formattedDate;
}

function fetchEmailsDaily() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var now = new Date();
    // var estOffset = 5 * 60 * 60 * 1000; // EST is UTC-5 hours
    // var time1 = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime() - estOffset;
    var now = new Date();
    var year = now.getFullYear();
    var month = now.getMonth();
    var date = now.getDate();

    // Create a Date object at 12 AM UTC
    var utcMidnight = new Date(year, month, date, 0, 0, 0);

    // Convert to Eastern Time (ET)
    var estTime = new Date(utcMidnight.toLocaleString("en-US", { timeZone: "America/New_York" }));

    // Get timestamp for 12 AM ET
    var time1 = estTime.getTime();
    var time2 = now.getTime();

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
            // msg.getPlainBody().substring(0, 500) // Limiting body size to 500 chars
            time1,
            time2,
            getFormattedDate(time1),
            getFormattedDate(time2),
            messages.length
        ]);
    });

    Logger.log("Emails inserted successfully.");
}
