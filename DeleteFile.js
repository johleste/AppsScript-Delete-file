function deleteFileFromID() {
    // Get the active spreadsheet and the value in cell A1 (File ID)
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var fileId = sheet.getRange('A1').getValue();
    
    // If the file ID is not entered, alert the user
    if (!fileId) {
      SpreadsheetApp.getUi().alert('Please enter a valid file ID in cell A1.');
      return;
    }
    
    try {
      // Get the file by its ID
      var file = DriveApp.getFileById(fileId);
      
      // Delete the file
      file.setTrashed(true);
      
      // Notify the user that the file has been deleted
      SpreadsheetApp.getUi().alert('File has been moved to trash.');
      
    } catch (e) {
      // If there's an error (like invalid file ID), notify the user
      SpreadsheetApp.getUi().alert('Error: ' + e.message);
    }
  }
  