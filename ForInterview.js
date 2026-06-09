// ==========================================
// FOR INTERVIEW - Letter Generation & Email Workflow
// ==========================================

// Constants for For Interview workflow
const FOR_INTERVIEW = {
  SHEET_NAME: 'LETTER - FOR INTERVIEW',
  TEMPLATE_ID: '1u9gWRR9UV5_ENJ8pEjueBWVIh3dHr-eNINItRZMDAGs',
  COL_LAST_NAME: 1,
  COL_FIRST_NAME: 2,
  COL_EMAIL: 7,
  COL_TEMPLATE_COLS_START: 15, // Column O
  COL_TEMPLATE_COLS_END: 18,   // Column R
  COL_INTERVIEW_LINK: 19,      // Column S
  COL_INTERVIEW_PROGRESS: 20,  // Column T
  COL_REGENERATE: 21,          // Column U
  START_ROW: 2
};

/**
 * Check if columns O-R have data in the FOR INTERVIEW tab
 */
function checkColumnsOtoRInSheet(sheetName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      return { hasData: false, message: 'Sheet not found' };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < FOR_INTERVIEW.START_ROW) {
      return { hasData: false };
    }

    const dataRange = sheet.getRange(
      FOR_INTERVIEW.START_ROW,
      1,
      lastRow - FOR_INTERVIEW.START_ROW + 1,
      FOR_INTERVIEW.COL_TEMPLATE_COLS_END
    ).getValues();

    for (let i = 0; i < dataRange.length; i++) {
      const rowData = dataRange[i];
      const valA = rowData[0]; // Column A

      if (valA && valA.toString().trim() !== '') {
        for (let colIdx = FOR_INTERVIEW.COL_TEMPLATE_COLS_START - 1; colIdx <= FOR_INTERVIEW.COL_TEMPLATE_COLS_END - 1; colIdx++) {
          const cellValue = rowData[colIdx];
          if (!cellValue || cellValue.toString().trim() === '') {
            const rowNum = FOR_INTERVIEW.START_ROW + i;
            return {
              hasData: false,
              message: 'Missing data in columns O-R for applicant at row ' + rowNum
            };
          }
        }
      }
    }

    return { hasData: true };
  } catch (e) {
    return { hasData: false, message: e.message };
  }
}


/**
 * Create main folder and For Interview subfolder
 */
function forInterviewCreateFolders() {
  try {
    const parentFolderId = "16Os72EpQfNxY6mFLd78qWnqlKMB5ZS03";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("SELECT POSITION");

    if (!sheet) throw new Error("Sheet 'SELECT POSITION' not found.");

    const position = sheet.getRange("B2").getValue();
    const assignedOffice = sheet.getRange("C2").getValue();

    if (!position || position.toString().trim() === "") {
      throw new Error("Position (Cell B2) is empty.");
    }
    if (!assignedOffice || assignedOffice.toString().trim() === "") {
      throw new Error("Assigned Office (Cell C2) is empty.");
    }

    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM");
    const folderName = position + " - " + assignedOffice + " (" + dateStr + ")";

    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const props = PropertiesService.getDocumentProperties();

    // INTEGRATED FIX: Check specific property and validate against current folder name
    let mainFolderId = props.getProperty('forInterviewMainFolderId');
    let mainFolder = null;

    if (mainFolderId) {
      try {
        const tempFolder = DriveApp.getFolderById(mainFolderId);
        // Verify if the stored ID actually matches our new target position folder name
        if (tempFolder.getName() === folderName) {
          mainFolder = tempFolder;
        }
      } catch (e) {
        mainFolder = null;
      }
    }

    if (!mainFolder) {
      const existingFolders = parentFolder.getFoldersByName(folderName);
      if (existingFolders.hasNext()) {
        mainFolder = existingFolders.next();
      } else {
        mainFolder = parentFolder.createFolder(folderName);
      }
      mainFolderId = mainFolder.getId();
      props.setProperty('forInterviewMainFolderId', mainFolderId);
    }

    // INTEGRATED FIX: Ensure the subfolder's parent matches the current main folder
    let forInterviewSubFolderId = props.getProperty('forInterviewSubFolderId');
    let forInterviewSubFolder = null;

    if (forInterviewSubFolderId) {
      try {
        const tempSub = DriveApp.getFolderById(forInterviewSubFolderId);
        if (tempSub.getParents().hasNext() && tempSub.getParents().next().getId() === mainFolderId) {
          forInterviewSubFolder = tempSub;
        }
      } catch (e) {
        forInterviewSubFolder = null;
      }
    }

    if (!forInterviewSubFolder) {
      const existingSubFolders = mainFolder.getFoldersByName('For Interview');
      if (existingSubFolders.hasNext()) {
        forInterviewSubFolder = existingSubFolders.next();
      } else {
        forInterviewSubFolder = mainFolder.createFolder('For Interview');
      }
      forInterviewSubFolderId = forInterviewSubFolder.getId();
      props.setProperty('forInterviewSubFolderId', forInterviewSubFolderId);
    }

    return {
      mainFolderId: mainFolderId,
      forInterviewSubFolderId: forInterviewSubFolderId,
      folderUrl: forInterviewSubFolder.getUrl()
    };
  } catch (e) {
    throw new Error('Error creating folders: ' + e.message);
  }
}

/**
 * Generate PDFs from sheet and return list of processed applicants
 */
function forInterviewGeneratePDFs(targetFolderId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FOR_INTERVIEW.SHEET_NAME);

    if (!sheet) throw new Error("Sheet '" + FOR_INTERVIEW.SHEET_NAME + "' not found!");

    const data = sheet.getDataRange().getDisplayValues();
    const header = data[0];
    const rows = data.slice(1).filter(row => row[FOR_INTERVIEW.COL_TEMPLATE_COLS_START - 1] && row[FOR_INTERVIEW.COL_TEMPLATE_COLS_START - 1].toString().trim() !== "");

    rows.sort((a, b) => {
      const lastNameA = String(a[0] || "").trim().toLowerCase();
      const lastNameB = String(b[0] || "").trim().toLowerCase();
      if (lastNameA !== lastNameB) return lastNameA.localeCompare(lastNameB);
      return String(a[1] || "").trim().toLowerCase().localeCompare(String(b[1] || "").trim().toLowerCase());
    });

    const templateFile = DriveApp.getFileById(FOR_INTERVIEW.TEMPLATE_ID);
    const destinationFolder = DriveApp.getFolderById(targetFolderId);
    
    // Use batch processing with key for For Interview
    const batchKey = 'forInterview_pdf_generation_' + SpreadsheetApp.getActiveSpreadsheet().getId();
    const batchResult = processPDFBatch(batchKey, rows, header, templateFile, destinationFolder, 20); // Process 20 at a time

    let returnMessage = batchResult.message;
    
    // If batch is still processing, suggest running again
    if (!batchResult.completed) {
      returnMessage += '\n\nTo continue processing remaining applicants (total: ' + batchResult.totalRows + '), run this step again.';
    } else {
      // Batch is complete, clear the state
      clearBatchState(batchKey);
      returnMessage = 'PDF generation completed! ' + batchResult.totalProcessed + ' PDFs generated.';
    }

    return {
      success: true,
      count: batchResult.totalProcessed,
      applicants: batchResult.allApplicants,
      completed: batchResult.completed,
      message: returnMessage
    };
  } catch (e) {
    throw new Error('Error generating PDFs: ' + e.message);
  }
}

/**
 * Generate Google Drive links and insert into spreadsheet
 */
function forInterviewGenerateLinks() {
  try {
    const settings = PropertiesService.getDocumentProperties();
    const folderId = settings.getProperty('forInterviewSubFolderId');

    if (!folderId) {
      throw new Error('For Interview subfolder not found. Please run Step 2 first.');
    }

    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByType(MimeType.PDF);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FOR_INTERVIEW.SHEET_NAME);

    if (!sheet) {
      throw new Error('Sheet "' + FOR_INTERVIEW.SHEET_NAME + '" not found.');
    }

    const fileData = [];
    while (files.hasNext()) {
      const file = files.next();
      fileData.push({
        name: file.getName(),
        url: file.getUrl()
      });
    }

    fileData.sort((a, b) => {
      return a.name.toLowerCase().localeCompare(b.name.toLowerCase());
    });

    const links = fileData.map(item => [item.url]);

    if (links.length > 0) {
      sheet.getRange(FOR_INTERVIEW.START_ROW, FOR_INTERVIEW.COL_INTERVIEW_LINK, links.length, 1).setValues(links);
    }

    return 'Successfully generated and inserted ' + links.length + ' Google Drive links into Column S.';
  } catch (e) {
    throw new Error('Error generating Drive links: ' + e.message);
  }
}

/**
 * Get the For Interview folder URL
 */
function forInterviewGetFolderUrl() {
  try {
    const settings = PropertiesService.getDocumentProperties();
    const folderId = settings.getProperty('forInterviewSubFolderId');

    if (!folderId) {
      throw new Error('For Interview subfolder not found. Please run the process first.');
    }

    const folder = DriveApp.getFolderById(folderId);
    return folder.getUrl();
  } catch (e) {
    throw new Error('Error retrieving folder URL: ' + e.message);
  }
}

/**
 * Backup the Letter - For Interview sheet to the For Interview folder
 */
function forInterviewBackupSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(FOR_INTERVIEW.SHEET_NAME);

    if (!sourceSheet) {
      throw new Error('Sheet "' + FOR_INTERVIEW.SHEET_NAME + '" not found.');
    }

    const settings = PropertiesService.getDocumentProperties();
    const forInterviewSubFolderId = settings.getProperty('forInterviewSubFolderId');

    if (!forInterviewSubFolderId) {
      throw new Error('For Interview subfolder not found. Please run Step 2 first.');
    }

    const forInterviewFolder = DriveApp.getFolderById(forInterviewSubFolderId);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    const backupFileName = 'LETTER - FOR INTERVIEW_' + timestamp + '.csv';

    const data = sourceSheet.getDataRange().getValues();

    let csvContent = '';
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const csvRow = row.map(cell => {
        if (typeof cell === 'string' && (cell.includes(',') || cell.includes('"') || cell.includes('\n'))) {
          return '"' + cell.replace(/"/g, '""') + '"';
        }
        return cell;
      }).join(',');
      csvContent += csvRow + '\n';
    }

    const backupBlob = Utilities.newBlob(csvContent, MimeType.CSV, backupFileName);
    forInterviewFolder.createFile(backupBlob);

    return {
      message: 'Backup successful! LETTER - FOR INTERVIEW has been saved to the For Interview folder.',
      folderUrl: forInterviewFolder.getUrl()
    };
  } catch (e) {
    throw new Error('Error backing up sheet: ' + e.message);
  }
}

/**
 * Send emails to interview applicants
 */
function forInterviewSendEmails() {
  // PASTE YOUR DEPLOYED WEB APP URL HERE
  const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxJpyg6KPFUMxeHSOdOVnVe4WyN6JssT9DhoufEn2pE7vIp02joOQ6jZVD-FwZCLKW7FQ/exec";

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FOR_INTERVIEW.SHEET_NAME);

    if (!sheet) throw new Error('Sheet "' + FOR_INTERVIEW.SHEET_NAME + '" not found.');

    const lastRow = sheet.getLastRow();
    if (lastRow < FOR_INTERVIEW.START_ROW) {
      return { status: 'No applicants found', count: 0 };
    }

    const data = sheet.getRange(FOR_INTERVIEW.START_ROW, 1, lastRow - FOR_INTERVIEW.START_ROW + 1, FOR_INTERVIEW.COL_INTERVIEW_PROGRESS).getValues();
    let emailCount = 0;
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const applicantName = row[0];
      const applicantLName = row[12]; // Column M
      const salutation = row[11]; // Column L
      const email = row[FOR_INTERVIEW.COL_EMAIL - 1];
      const driveLink = row[FOR_INTERVIEW.COL_INTERVIEW_LINK - 1];
      const position = row[7]; // Column H
      const office = row[8]; // Column I
      const statusCell = sheet.getRange(FOR_INTERVIEW.START_ROW + i, FOR_INTERVIEW.COL_INTERVIEW_PROGRESS);

      if (!applicantName || applicantName.toString().trim() === '') continue;

      if (!email || email.toString().trim() === '' || !driveLink || driveLink.toString().trim() === '') {
        statusCell.setValue('Not sent - missing email or link (' + now + ')');
        continue;
      }

      const subject = 'Job Application Update - Notice of Interview ' + '[' + position + ']';
      const body = 'Dear ' + salutation + ' ' + applicantLName + ',\n\n' +
        'Good day!\n\n' +
        'Congratulations! You have passed the written examination and have been selected to proceed to the interview stage.\n\n' +
        'Link: ' + driveLink + '\n\n' +
        'Please arrive at the interview site 5-10 minutes early. We look forward to meeting you!\n\n' +
        'Kindly acknowledge receipt of this email. If you have any questions, please do not hesitate to contact us.\n\n' +
        'Best regards,\nDOJ RPO V - Human Resource Unit';

      // --- INTEGRATED PROXY CALL ---
      const payload = {
        recipient: email.toString().trim(),
        cc: 'orp05.hiring@gmail.com',
        replyTo: 'orp05.hiring@gmail.com',
        subject: subject,
        body: body
      };

      const options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(WEB_APP_URL, options);

      if (response.getContentText() === "Success") {
        statusCell.setValue('Sent (' + now + ')');
        // Log the sent letter
        logSentLetter('LETTER - FOR INTERVIEW', position || '', office || '', applicantName || '');
        emailCount++;
      } else {
        statusCell.setValue('Error: Proxy failed (' + now + ')');
      }
    }

    return { status: 'Emails sent', count: emailCount };
  } catch (e) {
    throw new Error('Error sending email notifications: ' + e.message);
  }
}

function forInterviewGenerateIndividualPDFs() {
  try {
    const props = PropertiesService.getDocumentProperties();
    let folderId = props.getProperty('forInterviewSubFolderId');
    let destinationFolder = null;

    if (folderId) {
      try {
        destinationFolder = DriveApp.getFolderById(folderId);
      } catch (folderError) {
        const created = forInterviewCreateFolders();
        folderId = created.forInterviewSubFolderId;
        destinationFolder = DriveApp.getFolderById(folderId);
      }
    }

    if (!destinationFolder) {
      const created = forInterviewCreateFolders();
      folderId = created.forInterviewSubFolderId;
      destinationFolder = DriveApp.getFolderById(folderId);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FOR_INTERVIEW.SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + FOR_INTERVIEW.SHEET_NAME + '" not found.');

    const data = sheet.getDataRange().getDisplayValues();
    const header = data[0] || [];
    const rowsToProcess = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowIndex = FOR_INTERVIEW.START_ROW + i - 1;
      const regenerateVal = row[FOR_INTERVIEW.COL_REGENERATE - 1];
      const shouldProcess = regenerateVal === true || String(regenerateVal).toLowerCase() === 'true';
      if (!shouldProcess) continue;
      if (!row[FOR_INTERVIEW.COL_TEMPLATE_COLS_START - 1] || row[FOR_INTERVIEW.COL_TEMPLATE_COLS_START - 1].toString().trim() === '') continue;
      rowsToProcess.push({ row: row, rowIndex: rowIndex });
    }

    if (rowsToProcess.length === 0) {
      throw new Error('No items checked in REGENERATE column (U). Please check at least one checkbox to proceed.');
    }

    const templateFile = DriveApp.getFileById(FOR_INTERVIEW.TEMPLATE_ID);
    const processed = [];

    rowsToProcess.forEach((rowObj) => {
      const row = rowObj.row;
      const rowIndex = rowObj.rowIndex;
      const lastName = String(row[FOR_INTERVIEW.COL_LAST_NAME - 1] || '').trim();
      const firstName = String(row[FOR_INTERVIEW.COL_FIRST_NAME - 1] || '').trim();
      const fileName = (lastName || 'Applicant') + (firstName ? (', ' + firstName) : '');

      try {
        const copy = templateFile.makeCopy(fileName, destinationFolder);
        const doc = DocumentApp.openById(copy.getId());
        const body = doc.getBody();

        header.forEach((label, j) => {
          body.replaceText('{{' + label + '}}', row[j]);
        });

        doc.saveAndClose();
        const pdfBlob = copy.getAs(MimeType.PDF);
        const pdfFile = destinationFolder.createFile(pdfBlob).setName(fileName + '.pdf');
        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        copy.setTrashed(true);

        const pdfUrl = pdfFile.getUrl();
        sheet.getRange(rowIndex, FOR_INTERVIEW.COL_INTERVIEW_LINK).setValue(pdfUrl);
        processed.push(fileName);
      } catch (itemError) {
        console.log('Error generating interview PDF for row ' + rowIndex + ': ' + itemError.message);
      }
    });

    return { success: true, count: processed.length, applicants: processed };
  } catch (e) {
    throw new Error('Error generating individual interview PDFs: ' + e.message);
  }
}

function forInterviewSendSelectedEmails() {
  const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxJpyg6KPFUMxeHSOdOVnVe4WyN6JssT9DhoufEn2pE7vIp02joOQ6jZVD-FwZCLKW7FQ/exec";

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FOR_INTERVIEW.SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + FOR_INTERVIEW.SHEET_NAME + '" not found.');

    const lastRow = sheet.getLastRow();
    if (lastRow < FOR_INTERVIEW.START_ROW) {
      return { status: 'No applicants found', count: 0 };
    }

    const data = sheet.getRange(FOR_INTERVIEW.START_ROW, 1, lastRow - FOR_INTERVIEW.START_ROW + 1, FOR_INTERVIEW.COL_REGENERATE).getValues();

    const hasSelected = data.some(row => {
      const val = row[FOR_INTERVIEW.COL_REGENERATE - 1];
      return val === true || String(val).toLowerCase() === 'true';
    });
    if (!hasSelected) {
      throw new Error('No items checked in REGENERATE column. Please check at least one checkbox to proceed.');
    }

    let emailCount = 0;
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const regenerateVal = row[FOR_INTERVIEW.COL_REGENERATE - 1];
      const shouldSend = regenerateVal === true || String(regenerateVal).toLowerCase() === 'true';
      if (!shouldSend) continue;

      const applicantName = row[0];
      const applicantLName = row[12]; // Column M
      const salutation = row[11]; // Column L
      const email = row[FOR_INTERVIEW.COL_EMAIL - 1];
      const driveLink = row[FOR_INTERVIEW.COL_INTERVIEW_LINK - 1];
      const position = row[7];
      const office = row[8];
      const statusCell = sheet.getRange(FOR_INTERVIEW.START_ROW + i, FOR_INTERVIEW.COL_INTERVIEW_PROGRESS);

      if (!applicantName || applicantName.toString().trim() === '') {
        statusCell.setValue('Not sent - missing name (' + now + ')');
        sheet.getRange(FOR_INTERVIEW.START_ROW + i, FOR_INTERVIEW.COL_REGENERATE).setValue(false);
        continue;
      }

      if (!email || email.toString().trim() === '' || !driveLink || driveLink.toString().trim() === '') {
        statusCell.setValue('Not sent - missing email or link (' + now + ')');
        sheet.getRange(FOR_INTERVIEW.START_ROW + i, FOR_INTERVIEW.COL_REGENERATE).setValue(false);
        continue;
      }

      const subject = 'Job Application Update - Notice of Interview ' + '[' + position + ']';
      const body = 'Dear ' + salutation + ' ' + applicantLName + ',\n\n' +
        'Good day!\n\n' +
        'Congratulations! You have passed the written examination and have been selected to proceed to the interview stage.\n\n' +
        'Link: ' + driveLink + '\n\n' +
        'Please arrive at the interview site 5-10 minutes early. We look forward to meeting you!\n\n' +
        'Kindly acknowledge receipt of this email. If you have any questions, please do not hesitate to contact us.\n\n' +
        'Best regards,\nDOJ RPO V - Human Resource Unit';

      const payload = {
        recipient: email.toString().trim(),
        cc: 'orp05.hiring@gmail.com',
        replyTo: 'orp05.hiring@gmail.com',
        subject: subject,
        body: body
      };

      const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(WEB_APP_URL, options);
      if (response.getContentText() === 'Success') {
        statusCell.setValue('Sent (re-sent) (' + now + ')');
        logSentLetter('LETTER - FOR INTERVIEW', position || '', office || '', applicantName || '');
        emailCount++;
        sheet.getRange(FOR_INTERVIEW.START_ROW + i, FOR_INTERVIEW.COL_REGENERATE).setValue(false);
      } else {
        statusCell.setValue('Error: Proxy failed (' + now + ')');
      }
    }

    return { status: 'Selected emails processed', count: emailCount };
  } catch (e) {
    throw new Error('Error sending selected interview emails: ' + e.message);
  }
}

/**
 * Master function for For Interview workflow
 */
function runInterviewCompleteProcess() {
  try {
    const folderIds = forInterviewCreateFolders();
    const pdfResult = forInterviewGeneratePDFs(folderIds.forInterviewSubFolderId);
    
    let message = pdfResult.message || ('Generated ' + pdfResult.count + ' PDFs');
    if (!pdfResult.completed) {
      message = pdfResult.message + '\n\nStep 2 is still running. Click "Step 2 - Generate PDFs" again to continue processing remaining applicants.';
    }
    
    return {
      success: true,
      folders: folderIds,
      pdfGeneration: {
        ...pdfResult,
        message: message
      }
    };
  } catch (e) {
    throw new Error('Error in complete process: ' + e.message);
  }
}