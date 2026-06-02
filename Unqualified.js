// ==========================================
// UNQUALIFIED - Letter Generation & Email Workflow
// ==========================================

const UNQUALIFIED = {
  SHEET_NAME: 'LETTER - DQ',
  TEMPLATE_ID: '1PgZESY5toAEzZrUZnczuu3KP_NNWqFG6kSMgskFehgs',
  COL_LAST_NAME: 1,
  COL_FIRST_NAME: 2,
  COL_EMAIL: 7,
  COL_REASON: 10,
  COL_LINK: 16, // Column P
  COL_STATUS: 17, // Column Q
  COL_REGENERATE: 19, // Column S
  START_ROW: 2
};

function getUnqualifiedPositionFolder() {
  const parentFolderId = '16Os72EpQfNxY6mFLd78qWnqlKMB5ZS03';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('SELECT POSITION');
  if (!sheet) throw new Error("Sheet 'SELECT POSITION' not found.");

  const position = sheet.getRange('B2').getValue();
  const assignedOffice = sheet.getRange('C2').getValue();

  if (!position || position.toString().trim() === '') {
    throw new Error('Position (Cell B2) is empty.');
  }
  if (!assignedOffice || assignedOffice.toString().trim() === '') {
    throw new Error('Assigned Office (Cell C2) is empty.');
  }

  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  const folderName = position + ' - ' + assignedOffice + ' (' + dateStr + ')';

  const parentFolder = DriveApp.getFolderById(parentFolderId);
  const props = PropertiesService.getDocumentProperties();

  let mainFolderId = props.getProperty('unqualifiedMainFolderId');
  let mainFolder = null;
  
  if (mainFolderId) {
    try {
      const tempFolder = DriveApp.getFolderById(mainFolderId);
      // FIX: Only reuse if it matches our current target position folder name
      if (tempFolder.getName() === folderName) {
        mainFolder = tempFolder;
      }
    } catch (e) {
      mainFolder = null;
    }
  }

  // If no cached folder exists OR it didn't match the new position name
  if (!mainFolder) {
    const existingFolders = parentFolder.getFoldersByName(folderName);
    if (existingFolders.hasNext()) {
      mainFolder = existingFolders.next();
    } else {
      mainFolder = parentFolder.createFolder(folderName);
    }
    mainFolderId = mainFolder.getId();
    props.setProperty('unqualifiedMainFolderId', mainFolderId);
  }

  // Handle the 'Unqualified' specific subfolder
  let unqualifiedFolderId = props.getProperty('unqualifiedSubFolderId');
  let unqualifiedFolder = null;
  
  if (unqualifiedFolderId) {
    try {
      const tempSub = DriveApp.getFolderById(unqualifiedFolderId);
      // Ensure the subfolder's parent is actually our current main position folder
      if (tempSub.getParents().hasNext() && tempSub.getParents().next().getId() === mainFolderId) {
        unqualifiedFolder = tempSub;
      }
    } catch (e) {
      unqualifiedFolder = null;
    }
  }

  if (!unqualifiedFolder) {
    const existingUnqualFolders = mainFolder.getFoldersByName('Unqualified');
    if (existingUnqualFolders.hasNext()) {
      unqualifiedFolder = existingUnqualFolders.next();
    } else {
      unqualifiedFolder = mainFolder.createFolder('Unqualified');
    }
    unqualifiedFolderId = unqualifiedFolder.getId();
    props.setProperty('unqualifiedSubFolderId', unqualifiedFolderId);
  }

  return {
    mainFolderId: mainFolderId,
    unqualifiedSubFolderId: unqualifiedFolderId,
    folderUrl: unqualifiedFolder.getUrl()
  };
}

/**
 * Check if column R has data in the specified sheet (for unqualified validation)
 */
function checkColumnRInSheet(sheetName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      return { hasData: false, message: 'Sheet not found' };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { hasData: false };
    }

    const dataRange = sheet.getRange(2, 1, lastRow - 1, 18).getValues(); // Check up to column R

    for (let i = 0; i < dataRange.length; i++) {
      const rowData = dataRange[i];
      const valA = rowData[0]; // Column A

      if (valA && valA.toString().trim() !== '') {
        const colRValue = rowData[17]; // Column R (0-indexed)
        if (!colRValue || colRValue.toString().trim() === '') {
          const rowNum = 2 + i;
        throw new Error('Missing data in column R for applicant at row ' + rowNum);
        }
      }
    }

    return { hasData: true };
  } catch (e) {
    return { hasData: false, message: e.message };
  }
}


function unqualifiedGeneratePDFs(targetFolderId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(UNQUALIFIED.SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + UNQUALIFIED.SHEET_NAME + '" not found!');

    const data = sheet.getDataRange().getDisplayValues();
    const header = data[0];
    const rows = data.slice(1).filter(row => row[0] && row[0].toString().trim() !== '');

    rows.sort((a, b) => {
      const lastNameA = String(a[0] || '').trim().toLowerCase();
      const lastNameB = String(b[0] || '').trim().toLowerCase();
      if (lastNameA !== lastNameB) return lastNameA.localeCompare(lastNameB);
      return String(a[1] || '').trim().toLowerCase().localeCompare(String(b[1] || '').trim().toLowerCase());
    });

    const templateFile = DriveApp.getFileById(UNQUALIFIED.TEMPLATE_ID);
    const destinationFolder = DriveApp.getFolderById(targetFolderId);
    
    // Use batch processing with key for Unqualified
    const batchKey = 'unqualified_pdf_generation_' + SpreadsheetApp.getActiveSpreadsheet().getId();
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
    throw new Error('Error generating unqualified PDFs: ' + e.message);
  }
}

function unqualifiedGenerateLinks() {
  try {
    const settings = PropertiesService.getDocumentProperties();
    const folderId = settings.getProperty('unqualifiedSubFolderId');
    if (!folderId) {
      throw new Error('Unqualified subfolder not found. Please run Step 2 first.');
    }

    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByType(MimeType.PDF);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(UNQUALIFIED.SHEET_NAME);
    if (!sheet) {
      throw new Error('Sheet "' + UNQUALIFIED.SHEET_NAME + '" not found.');
    }

    const fileData = [];
    while (files.hasNext()) {
      const file = files.next();
      fileData.push({ name: file.getName(), url: file.getUrl() });
    }

    fileData.sort((a, b) => a.name.toLowerCase().localeCompare(b.name.toLowerCase()));

    const links = fileData.map(item => [item.url]);
    if (links.length > 0) {
      sheet.getRange(UNQUALIFIED.START_ROW, UNQUALIFIED.COL_LINK, links.length, 1).setValues(links);
    }

    return 'Successfully generated and inserted ' + links.length + ' Google Drive links into Column P.';
  } catch (e) {
    throw new Error('Error generating Unqualified Drive links: ' + e.message);
  }
}

function unqualifiedGetFolderUrl() {
  try {
    const settings = PropertiesService.getDocumentProperties();
    const folderId = settings.getProperty('unqualifiedSubFolderId');
    if (!folderId) {
      throw new Error('Unqualified subfolder not found. Please run the process first.');
    }
    return DriveApp.getFolderById(folderId).getUrl();
  } catch (e) {
    throw new Error('Error retrieving unqualified folder URL: ' + e.message);
  }
}

function unqualifiedBackupSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(UNQUALIFIED.SHEET_NAME);
    if (!sourceSheet) {
      throw new Error('Sheet "' + UNQUALIFIED.SHEET_NAME + '" not found.');
    }

    const settings = PropertiesService.getDocumentProperties();
    const folderId = settings.getProperty('unqualifiedSubFolderId');
    if (!folderId) {
      throw new Error('Unqualified subfolder not found. Please run Step 2 first.');
    }

    const folder = DriveApp.getFolderById(folderId);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    const backupFileName = 'LETTER - DQ_' + timestamp + '.csv';

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
    folder.createFile(backupBlob);
    return {
      message: 'Backup successful! LETTER - DQ has been saved to the Unqualified folder.',
      folderUrl: folder.getUrl()
    };
  } catch (e) {
    throw new Error('Error backing up DQ sheet: ' + e.message);
  }
}

function unqualifiedSendEmails() {
  // PASTE YOUR DEPLOYED WEB APP URL HERE
  const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyFPxd3UelHmFuh4fqQC7YPLpVk44rorubWx_My_0S2OV7Il4GlJC1wd7rq8aVKJKpKNg/exec";

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(UNQUALIFIED.SHEET_NAME);

    if (!sheet) throw new Error('Sheet "' + UNQUALIFIED.SHEET_NAME + '" not found.');

    const lastRow = sheet.getLastRow();
    if (lastRow < UNQUALIFIED.START_ROW) {
      return { status: 'No applicants found', count: 0 };
    }

    const data = sheet.getRange(UNQUALIFIED.START_ROW, 1, lastRow - UNQUALIFIED.START_ROW + 1, UNQUALIFIED.COL_STATUS).getValues();
    let emailCount = 0;
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const applicantName = row[0];
      const email = row[UNQUALIFIED.COL_EMAIL - 1];
      const driveLink = row[UNQUALIFIED.COL_LINK - 1];
      const position = row[7]; // Column H
      const office = row[8]; // Column I
      const statusCell = sheet.getRange(UNQUALIFIED.START_ROW + i, UNQUALIFIED.COL_STATUS);

      if (!applicantName || applicantName.toString().trim() === '') continue;

      if (!email || email.toString().trim() === '' || !driveLink || driveLink.toString().trim() === '') {
        statusCell.setValue('Not sent - missing email or link (' + now + ')');
        continue;
      }

      const subject = 'JOB APPLICATION UPDATE';
      const body = 'Dear Applicant,\n\n' +
        'Good day!\n\n' +
        'Please see attached file regarding your application.\n\n' +
        'Link: ' + driveLink;

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
        logSentLetter('LETTER - DQ', position || '', office || '', applicantName || '');
        emailCount++;
      } else {
        statusCell.setValue('Error: Proxy failed (' + now + ')');
      }

/**
      GmailApp.sendEmail(email.toString().trim(), subject, body, { replyTo: 'orp05.hiring@gmail.com' });
      statusCell.setValue('Sent (' + now + ')');
      emailCount++;
      */
    }

    return { status: 'Emails sent', count: emailCount };
  } catch (e) {
    throw new Error('Error sending email notifications: ' + e.message);
  }
}

function unqualifiedSendSelectedEmails() {
  const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyFPxd3UelHmFuh4fqQC7YPLpVk44rorubWx_My_0S2OV7Il4GlJC1wd7rq8aVKJKpKNg/exec";

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(UNQUALIFIED.SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + UNQUALIFIED.SHEET_NAME + '" not found.');

    const lastRow = sheet.getLastRow();
    if (lastRow < UNQUALIFIED.START_ROW) {
      return { status: 'No applicants found', count: 0 };
    }

    const data = sheet.getRange(UNQUALIFIED.START_ROW, 1, lastRow - UNQUALIFIED.START_ROW + 1, UNQUALIFIED.COL_REGENERATE).getValues();
    let emailCount = 0;
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const regenerateVal = row[UNQUALIFIED.COL_REGENERATE - 1];
      const shouldSend = regenerateVal === true || String(regenerateVal).toLowerCase() === 'true';
      if (!shouldSend) continue;

      const applicantName = row[0];
      const email = row[UNQUALIFIED.COL_EMAIL - 1];
      const driveLink = row[UNQUALIFIED.COL_LINK - 1];
      const position = row[7];
      const office = row[8];
      const statusCell = sheet.getRange(UNQUALIFIED.START_ROW + i, UNQUALIFIED.COL_STATUS);

      if (!applicantName || applicantName.toString().trim() === '') {
        statusCell.setValue('Not sent - missing name (' + now + ')');
        sheet.getRange(UNQUALIFIED.START_ROW + i, UNQUALIFIED.COL_REGENERATE).setValue(false);
        continue;
      }

      if (!email || email.toString().trim() === '' || !driveLink || driveLink.toString().trim() === '') {
        statusCell.setValue('Not sent - missing email or link (' + now + ')');
        sheet.getRange(UNQUALIFIED.START_ROW + i, UNQUALIFIED.COL_REGENERATE).setValue(false);
        continue;
      }

      const subject = 'JOB APPLICATION UPDATE';
      const body = 'Dear Applicant,\n\n' +
        'Good day!\n\n' +
        'Please see attached file regarding your application.\n\n' +
        'Link: ' + driveLink;

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
        logSentLetter('LETTER - DQ', position || '', office || '', applicantName || '');
        emailCount++;
        sheet.getRange(UNQUALIFIED.START_ROW + i, UNQUALIFIED.COL_REGENERATE).setValue(false);
      } else {
        statusCell.setValue('Error: Proxy failed (' + now + ')');
      }
    }

    return { status: 'Selected emails processed', count: emailCount };
  } catch (e) {
    throw new Error('Error sending selected emails: ' + e.message);
  }
}

function unqualifiedRunCompleteProcess() {
  } catch (e) {
    throw new Error('Error sending email notifications: ' + e.message);
  }
}

function unqualifiedRunCompleteProcess() {
  try {
    const folderIds = getUnqualifiedPositionFolder();
    const pdfResult = unqualifiedGeneratePDFs(folderIds.unqualifiedSubFolderId);
    
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
    throw new Error('Error in unqualified complete process: ' + e.message);
  }
}
