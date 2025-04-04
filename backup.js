function getEnv(key) {
    return PropertiesService.getScriptProperties().getProperty(key);
  }
  
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
  