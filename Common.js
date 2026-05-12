// ==========================================
// COMMON - Shared Utilities & Constants
// ==========================================

/**
 * Handles incoming POST requests from the spreadsheet.
 * Always executes as the owner of the script (orp05@doj.gov.ph).
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    
    MailApp.sendEmail({
      to: data.recipient,
      subject: data.subject,
      body: data.body,
      name: 'Department of Justice V - HR',
      replyTo: 'orp05@doj.gov.ph'
    });
    
    return ContentService.createTextOutput("Success");
  } catch (err) {
    // If this returns anything other than "Success", the spreadsheet throws the Proxy Error
    return ContentService.createTextOutput("Error: " + err.message);
  }
}

// Global constants
const SETTINGS_KEYS = {
  examFolderLink: 'examFolderLink',
  hasDisqualified: 'hasDisqualified',
  dqFolderLink: 'dqFolderLink',
  lastSuccessfulSignature: 'lastSuccessfulSignature'
};

const CONFIG = {
  SHEET_NAME: "SELECT POSITION",
  DROPDOWN_CELL: "B2",
  FOLDER_ID: "1EHQGn1eo19pyeGQaKk3_XFjhMm6YkPGa",
  SOURCE_DATA_SHEET: "SELECT POSITION",
  SOURCE_DATA_RANGE: "B6:B20"
};

/**
 * Build email message based on letter type
 */
function buildEmailMessage(label, fileLink) {
  if (label === 'DQ') {
    return {
      subject: 'Update regarding your application status',
      body: 'Dear Applicant,\n\n' +
        'Good day!\n\n' +
        'Thank you for your interest in applying for a vacant position at our office. We have received your application and appreciate the time you took to apply.\n\n' +
        'After a careful review of all applications and the specific requirements for the role, we have decided to move forward with other candidates at this time.\n\n' +
        'Please see the file in the link below for the complete details.\n\n' +
        'Link to file: ' + fileLink + '\n\n' +
        'We wish you the very best of luck with your future applications and all your professional endeavors!'
    };
  }

  return {
    subject: 'Notice of Written Examination',
    body: 'Dear Applicant,\n\n' +
      'Good day!\n\n' +
      'Thank you for your interest in the vacant position at our office. We have received your application and appreciate the time you took to apply.\n\n' +
      'Please see the file in the link below for the details of your written examination.\n\n' +
      'Kindly arrive at the examination site 5 to 10 minutes before the scheduled time. Please note that late examinees without a valid reason will not be allowed to take the exam.\n\n' +
      'Link: ' + fileLink
  };
}

/**
 * Get letter type label
 */
function getLetterTypeLabel(letterType) {
  const typeMap = {
    'forExam': 'FOR EXAM',
    'unqualified': 'UNQUALIFIED',
    'forInterview': 'FOR INTERVIEW',
    'failed': 'FAILED',
    'backup': 'BACKUP'
  };
  return typeMap[letterType] || 'FOR EXAM';
}

/**
 * Get sidebar state
 */
function getSidebarState() {
  const properties = PropertiesService.getDocumentProperties();
  return {
    folderLink: properties.getProperty(SETTINGS_KEYS.examFolderLink) || '',
    hasDisqualified: properties.getProperty(SETTINGS_KEYS.hasDisqualified) || 'no',
    dqFolderLink: properties.getProperty(SETTINGS_KEYS.dqFolderLink) || '',
    lastSuccessfulSignature: properties.getProperty(SETTINGS_KEYS.lastSuccessfulSignature) || ''
  };
}

/**
 * Save sidebar state
 */
function saveSidebarState(normalizedPayload, markSuccessful) {
  const properties = PropertiesService.getDocumentProperties();
  properties.setProperty(SETTINGS_KEYS.examFolderLink, normalizedPayload.folderLink);
  properties.setProperty(SETTINGS_KEYS.hasDisqualified, normalizedPayload.hasDisqualified);
  properties.setProperty(SETTINGS_KEYS.dqFolderLink, normalizedPayload.dqFolderLink);

  if (markSuccessful) {
    properties.setProperty(SETTINGS_KEYS.lastSuccessfulSignature, buildPayloadSignature(normalizedPayload));
  }
}

/**
 * Build payload signature (helper)
 */
function buildPayloadSignature(payload) {
  return JSON.stringify(payload);
}

/**
 * Mark process as done
 */
function markProcessAsDone() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SELECT POSITION");
  if (sheet) {
    sheet.getRange("B3").setValue("Last Completed: " + new Date().toLocaleString());
  }
  return true;
}

/**
 * Clear working sheet ranges when a new position is selected
 */
function clearAllWorkingSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetExam = ss.getSheetByName('LETTER - EXAM SCHED');
  if (sheetExam && sheetExam.getLastRow() >= 2) {
    sheetExam.getRange(2, 15, sheetExam.getLastRow() - 1, 6).clearContent();
  }

  const sheetDq = ss.getSheetByName('LETTER - DQ');
  if (sheetDq && sheetDq.getLastRow() >= 2) {
    sheetDq.getRange(2, 16, sheetDq.getLastRow() - 1, 3).clearContent();
  }

  const sheetInterview = ss.getSheetByName('LETTER - FOR INTERVIEW');
  if (sheetInterview && sheetInterview.getLastRow() >= 2) {
    sheetInterview.getRange(2, 15, sheetInterview.getLastRow() - 1, 6).clearContent();
  }

  const sheetFailed = ss.getSheetByName('LETTER - FAILED');
  if (sheetFailed && sheetFailed.getLastRow() >= 2) {
    sheetFailed.getRange(2, 12, sheetFailed.getLastRow() - 1, 3).clearContent();
  }
}

/**
 * Backup all sheets to folder
 */
function backupSheetToFolder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const cell = mainSheet.getRange(CONFIG.DROPDOWN_CELL);
  const cellValue = cell.getValue();

  const sheetsToBackup = [
    'LETTER - EXAM SCHED', 
    'LETTER - DQ', 
    'LETTER - FOR INTERVIEW', 
    'LETTER - FAILED'
  ];

  if (!cellValue) {
    SpreadsheetApp.getUi().alert("Please select a position first.");
    return;
  }

  try {
    const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const newFileName = cellValue + " - " + currentDate;
    const destFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);

    const backupFile = SpreadsheetApp.create(newFileName);
    const backupId = backupFile.getId();
    DriveApp.getFileById(backupId).moveTo(destFolder);

    const backupSS = SpreadsheetApp.openById(backupId);

    let sheetsCopied = 0;
    sheetsToBackup.forEach(name => {
      const sourceSheet = ss.getSheetByName(name);
      if (sourceSheet) {
        const sourceRange = sourceSheet.getDataRange();
        const displayValues = sourceRange.getDisplayValues();
        
        const backupSheet = backupSS.insertSheet(name);
        backupSheet.getRange(1, 1, displayValues.length, displayValues[0].length).setValues(displayValues);
        
        for (let col = 1; col <= displayValues[0].length; col++) {
          backupSheet.setColumnWidth(col, sourceSheet.getColumnWidth(col));
        }

        backupSheet.getRange(1, 1, 1, displayValues[0].length)
                   .setBackground('#b6d7a8')
                   .setFontWeight('bold');

        sheetsCopied++;
      }
    });

    const defaultSheet = backupSS.getSheetByName('Sheet1');
    if (defaultSheet) { backupSS.deleteSheet(defaultSheet); }

    // Clear data in original sheets
    clearAllWorkingSheets();

    // Unlock dropdown
    const sourceListSheet = ss.getSheetByName(CONFIG.SOURCE_DATA_SHEET);
    if (sourceListSheet) {
      const sourceListRange = sourceListSheet.getRange(CONFIG.SOURCE_DATA_RANGE);
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(sourceListRange)
        .setAllowInvalid(false)
        .build();
      cell.setDataValidation(rule);
    }

    SpreadsheetApp.getUi().alert("Backup Successful! Original sheets have been cleared for the next position.");

  } catch (err) {
    SpreadsheetApp.getUi().alert("Setup Error: " + err.message);
  }
}

/**
 * Handle edit events - Warns user before changing position
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();

  // Only trigger if editing the specific Dropdown Cell in the SELECT POSITION sheet
  if (sheet.getName() === CONFIG.SHEET_NAME && range.getA1Notation() === CONFIG.DROPDOWN_CELL) {
    const newValue = range.getValue();
    const oldValue = e.oldValue;

    // If the cell was cleared manually, do nothing
    if (!newValue) return;

    // Prompt the user for confirmation
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Changing Position',
      'Warning: Changing the position will clear all entered data in all working sheets (Exam, DQ, Interview, Failed). \n\nDo you want to proceed?',
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      // 1. User clicked "Accept & Clear" (YES)
      clearAllWorkingSheets();

      // 2. Mark that position was changed so sidebar resets on next load
      const properties = PropertiesService.getDocumentProperties();
      properties.setProperty('positionChanged', 'true');

      ui.alert("Sheets cleared. You can now proceed with the new position.");
    } else {
      // 3. User clicked "Cancel" (NO)
      // Revert the cell to the previous value without triggering a loop
      range.setValue(oldValue);
    }
  }
}
