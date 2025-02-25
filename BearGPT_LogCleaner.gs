// Configuration object - replace these values with your actual folder IDs
const CONFIG = {
  inputFolderId: '1ul32v7n4ajCGQHLXYqPfmBmPFee6zHkq',  // BearGPT-Logs/Input
  outputFolderId: '1Az6y1WrzXeXx3lKLuH6LlffLcfSldY5L'  // BearGPT-Logs/Output
};

function createTrigger() {
  // First, check if we have existing triggers and delete them
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Create new time-driven trigger that runs every minute
  ScriptApp.newTrigger('checkForNewFiles')
    .timeBased()
    .everyMinutes(1)
    .create();
}

function checkForNewFiles() {
  try {
    const inputFolder = DriveApp.getFolderById(CONFIG.inputFolderId);
    const files = inputFolder.getFiles();
    
    while (files.hasNext()) {
      const file = files.next();
      
      // Check if file has been processed (using properties)
      const processedFiles = PropertiesService.getScriptProperties().getProperty('processedFiles') || '[]';
      const processedFileIds = JSON.parse(processedFiles);
      
      if (!processedFileIds.includes(file.getId())) {
        // Process only Excel files
        if (file.getMimeType() === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
            file.getMimeType() === 'application/vnd.ms-excel') {
          processExcelFile(file);
          
          // Add file ID to processed files
          processedFileIds.push(file.getId());
          PropertiesService.getScriptProperties().setProperty('processedFiles', JSON.stringify(processedFileIds));
        }
      }
    }
  } catch (error) {
    Logger.log('Error in checkForNewFiles: ' + error.toString());
  }
}

function processExcelFile(file) {
  let finalSpreadsheet = null;
  
  try {
    // Get original filename without extension
    const originalName = file.getName().replace('.xlsx', '').replace('.xls', '');
    Logger.log('Processing file: ' + originalName);

    // Convert Excel to Google Sheets using Drive API
    const blob = file.getBlob();
    
    Logger.log('Converting Excel file...');
    const insertResponse = Drive.Files.insert({
      title: originalName,
      parents: [{ id: CONFIG.outputFolderId }],
      mimeType: 'application/vnd.google-apps.spreadsheet'
    }, blob, {
      convert: true,
      ocr: false
    });
    
    Logger.log('Conversion complete. New file ID: ' + insertResponse.id);
    
    
    // Open the converted spreadsheet
    finalSpreadsheet = SpreadsheetApp.openById(insertResponse.id);
    const sourceData = finalSpreadsheet.getSheets()[0].getDataRange().getValues();
    
    Logger.log('Data retrieved. Rows: ' + sourceData.length);
    if (sourceData.length > 0) {
      Logger.log('Columns: ' + sourceData[0].length);
      Logger.log('Headers: ' + sourceData[0].join(', '));
    }

    // Find the user_email column index
    const headerRow = sourceData[0];
    const emailColIndex = headerRow.indexOf('user_email');
    
    Logger.log('Email column index: ' + emailColIndex);
    
    if (emailColIndex !== -1) {
      // Create new data array without the user_email column
      const newData = sourceData.map(row => {
        return row.filter((cell, index) => index !== emailColIndex);
      });
      
      // Clear and write the filtered data
      const sheet = finalSpreadsheet.getActiveSheet();
      sheet.clear();
      sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
      
      Logger.log(`Successfully processed file: ${originalName}`);
      Logger.log(`Final spreadsheet ID: ${finalSpreadsheet.getId()}`);
      Logger.log(`Final spreadsheet URL: ${finalSpreadsheet.getUrl()}`);
    } else {
      Logger.log('Available columns: ' + headerRow.join(', '));
      throw new Error(`No user_email column found in file: ${originalName}`);
    }
    
  } catch (error) {
    Logger.log('Error in processExcelFile: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    
    // Clean up on error
    try {
      if (finalSpreadsheet) {
        DriveApp.getFileById(finalSpreadsheet.getId()).setTrashed(true);
      }
    } catch (cleanupError) {
      Logger.log('Error during cleanup: ' + cleanupError.toString());
    }
    
    throw error;
  }
}

function testFullProcess() {
  try {
    const inputFolder = DriveApp.getFolderById(CONFIG.inputFolderId);
    const files = inputFolder.getFiles();
    
    if (files.hasNext()) {
      const file = files.next();
      Logger.log('Testing full process for file: ' + file.getName());
      
      // Process the file using our main function
      processExcelFile(file);
      
      Logger.log('Full process test completed');
    } else {
      Logger.log('No files found in input folder');
    }
  } catch (error) {
    Logger.log('Test failed: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
  }
}


function testSingleFile() {
  try {
    const inputFolder = DriveApp.getFolderById(CONFIG.inputFolderId);
    const files = inputFolder.getFiles();
    
    if (files.hasNext()) {
      const file = files.next();
      Logger.log('Testing conversion of: ' + file.getName());
      
      // Create new spreadsheet first
      const tempSpreadsheet = SpreadsheetApp.create('Temp_Conversion_' + new Date().getTime());
      const tempFile = DriveApp.getFileById(tempSpreadsheet.getId());
      Logger.log('Created temporary spreadsheet with ID: ' + tempFile.getId());
      
      // Copy the Excel content
      const blob = file.getBlob();
      const url = "https://docs.google.com/spreadsheets/d/" + tempSpreadsheet.getId() + "/import";
      const options = {
        method: "post",
        headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        payload: blob.getBytes(),
        muteHttpExceptions: true
      };
      
      UrlFetchApp.fetch(url, options);
      
      // Wait a moment for the import to complete
      Utilities.sleep(2000);
      
      // Verify the data
      const data = tempSpreadsheet.getSheets()[0].getDataRange().getValues();
      Logger.log('Conversion successful');
      Logger.log('Rows: ' + data.length);
      if (data.length > 0) {
        Logger.log('Columns: ' + data[0].length);
        Logger.log('Headers: ' + data[0].join(', '));
      }
      
      // Clean up
      tempFile.setTrashed(true);
      Logger.log('Temporary file cleaned up');
      
    } else {
      Logger.log('No files found in input folder');
    }
  } catch (error) {
    Logger.log('Test failed: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    
    // Additional error details
    if (error.details) {
      Logger.log('Error details: ' + JSON.stringify(error.details));
    }
  }
}

function validateFolders() {
  try {
    const inputFolder = DriveApp.getFolderById(CONFIG.inputFolderId);
    const outputFolder = DriveApp.getFolderById(CONFIG.outputFolderId);
    Logger.log('Input folder name: ' + inputFolder.getName());
    Logger.log('Output folder name: ' + outputFolder.getName());
    Logger.log('Folder validation successful');
    return true;
  } catch (error) {
    Logger.log('Folder validation failed: ' + error.toString());
    return false;
  }
}

function clearProcessedFiles() {
  PropertiesService.getScriptProperties().deleteProperty('processedFiles');
  Logger.log('Processed files history cleared');
}
