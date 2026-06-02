// ==========================================
// FAILED - Letter Generation & Backup Workflow
// ==========================================

const FAILED = {
  SHEET_NAME: 'LETTER - FAILED',
  TEMPLATE_ID: '1-08ufjwnG0bCV9nrZ6LbOwA_DDBlr8nfUNZYncwFGP8',
  COL_LAST_NAME: 1,
  COL_FIRST_NAME: 2,
  COL_EMAIL: 6,              // Column F
  // --- INTEGRATED NOTES FOR LETTER RECIPIENT FIELDS ---
  COL_UPPER_SALUTATION: 9,   // Column I
  COL_UPPER_FULLNAME: 10,    // Column J
  COL_PROPER_SALUTATION: 11, // Column K
  COL_PROPER_LASTNAME: 12,   // Column L
  // --------------------------------------------------
  COL_EMAIL_DATE: 13,        // Column M
  COL_LINK: 14,              // Column N
  COL_STATUS: 15,            // Column O
  COL_REGENERATE: 16,         // Column P
  START_ROW: 2
};

function getFailedPositionFolder() {
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

  let mainFolderId = props.getProperty('failedMainFolderId');
  let mainFolder = null;
  
  if (mainFolderId) {
    try {
      const tempFolder = DriveApp.getFolderById(mainFolderId);
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
    props.setProperty('failedMainFolderId', mainFolderId);
  }

  let failedFolderId = props.getProperty('failedSubFolderId');
  let failedFolder = null;
  
  if (failedFolderId) {
    try {
      const tempSub = DriveApp.getFolderById(failedFolderId);
      if (tempSub.getParents().hasNext() && tempSub.getParents().next().getId() === mainFolderId) {
        failedFolder = tempSub;
      }
    } catch (e) {
      failedFolder = null;
    }
  }

  if (!failedFolder) {
    const existingFailedFolders = mainFolder.getFoldersByName('Failed');
    if (existingFailedFolders.hasNext()) {
      failedFolder = existingFailedFolders.next();
    } else {
      failedFolder = mainFolder.createFolder('Failed');
    }
    failedFolderId = failedFolder.getId();
    props.setProperty('failedSubFolderId', failedFolderId);
  }

  return {
    mainFolderId: mainFolderId,
    failedSubFolderId: failedFolderId,
    folderUrl: failedFolder.getUrl()
  };
}

function checkColumnNInSheet(sheetName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      return { hasData: false, message: 'Sheet not found' };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < FAILED.START_ROW) {
      return { hasData: false };
    }

    const dataRange = sheet.getRange(FAILED.START_ROW, 1, lastRow - FAILED.START_ROW + 1, FAILED.COL_STATUS).getValues();

    for (let i = 0; i < dataRange.length; i++) {
      const rowData = dataRange[i];
      const valA = rowData[0];

      if (valA && valA.toString().trim() !== '') {
        const statusValue = rowData[FAILED.COL_STATUS - 1];
        if (!statusValue || statusValue.toString().trim() === '') {
          const rowNum = FAILED.START_ROW + i;
          throw new Error('Missing data in column N for applicant at row ' + rowNum);
        }
      }
    }

    return { hasData: true };
  } catch (e) {
    return { hasData: false, message: e.message };
  }
}

function checkFailedColumnM() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(FAILED.SHEET_NAME);
  if (!sheet) {
    return { passed: false, message: 'Sheet "' + FAILED.SHEET_NAME + '" not found.' };
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < FAILED.START_ROW) {
    return { passed: false, message: 'No data rows found in LETTER - FAILED.' };
  }

  const names = sheet.getRange(FAILED.START_ROW, 1, lastRow - FAILED.START_ROW + 1, 2).getValues();
  const emailDates = sheet.getRange(FAILED.START_ROW, FAILED.COL_EMAIL_DATE, lastRow - FAILED.START_ROW + 1, 1).getValues();

  for (let i = 0; i < names.length; i++) {
    const nameValue = names[i][0];
    if (nameValue && nameValue.toString().trim() !== '') {
      const dateValue = emailDates[i][0];
      if (dateValue === '' || dateValue === null || dateValue.toString().trim() === '') {
        return { passed: false, message: 'Missing Column M value on row ' + (FAILED.START_ROW + i) };
      }
    }
  }

  return { passed: true };
}

function failedVerifyAlignment() {
  const settings = PropertiesService.getDocumentProperties();
  const folderId = settings.getProperty('failedSubFolderId');
  if (!folderId) {
    throw new Error('Failed subfolder not found. Please run Step 2 first.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(FAILED.SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet "' + FAILED.SHEET_NAME + '" not found.');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < FAILED.START_ROW) {
    return { passed: false, message: 'No data rows found to verify.' };
  }

  const data = sheet.getDataRange().getDisplayValues();
  const rows = data.slice(FAILED.START_ROW - 1).filter(row => row[0] && row[0].toString().trim() !== '');
  const expectedNames = [];

  for (let i = 0; i < rows.length; i++) {
    const lastName = String(rows[i][0] || '').trim(); // Column A
    const firstName = String(rows[i][1] || '').trim(); // Column B
    if (lastName && lastName !== '') {
      expectedNames.push(lastName + ', ' + firstName);
    }
  }

  expectedNames.sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));

  const folder = DriveApp.getFolderById(folderId);
  const fileIterator = folder.getFilesByType(MimeType.PDF);
  const fileNames = [];
  while (fileIterator.hasNext()) {
    fileNames.push(fileIterator.next().getName());
  }

  fileNames.sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));

  if (expectedNames.length !== fileNames.length) {
    return { passed: false, message: 'Expected ' + expectedNames.length + ' PDF(s), but found ' + fileNames.length + '.' };
  }

  for (let index = 0; index < expectedNames.length; index++) {
    const normalizedFileName = fileNames[index].replace(/\.pdf$/i, '');
    if (expectedNames[index].toLowerCase() !== normalizedFileName.toLowerCase()) {
      return {
        passed: false,
        message: 'Alignment mismatch discovered at entry indexes.'
      };
    }
  }

  return { passed: true, message: 'Alignment verified' };
}

function failedSendEmails() {
  const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbyFPxd3UelHmFuh4fqQC7YPLpVk44rorubWx_My_0S2OV7Il4GlJC1wd7rq8aVKJKpKNg/exec';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(FAILED.SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet "' + FAILED.SHEET_NAME + '" not found.');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < FAILED.START_ROW) {
    return { status: 'No applicants found', count: 0 };
  }

  const data = sheet.getRange(FAILED.START_ROW, 1, lastRow - FAILED.START_ROW + 1, FAILED.COL_STATUS).getValues();
  let emailCount = 0;
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const applicantName = row[0]; // Column A
    const email = row[FAILED.COL_EMAIL - 1]; // Column F
    const driveLink = row[FAILED.COL_LINK - 1]; // Column N
    
    // --- MATCHES YOUR SHEET LAYOUT VISUALS ---
    const position = row[6]; // Index 6 is Column G (POSITION EXT)
    const office = row[7];   // Index 7 is Column H (ASSIGNED OFFICE)
    // ----------------------------------------
    
    const statusCell = sheet.getRange(FAILED.START_ROW + i, FAILED.COL_STATUS);

    if (!applicantName || applicantName.toString().trim() === '') continue;

    if (!email || email.toString().trim() === '' || !driveLink || driveLink.toString().trim() === '') {
      statusCell.setValue('Not sent - missing email or link (' + now + ')');
      continue;
    }

    const subject = 'JOB APPLICATION UPDATE';
    const body = 'Dear Applicant,\n\n' +
      'Good day!\n\n' +
      'Please see attached file regarding your application.\n\n' +
      'Link: ' + driveLink + '\n\n' +
      'Best regards,\n' +
      'DOJ RPO V - Human Resource Unit';

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
      statusCell.setValue('Sent (' + now + ')');
      // Log the sent letter using the reference from Column G and H
      logSentLetter('LETTER - FAILED', position || '', office || '', applicantName || '');
      emailCount++;
    } else {
      statusCell.setValue('Error: Proxy failed (' + now + ')');
    }
  }

  return { status: 'Emails sent', count: emailCount };
}

function failedSendSelectedEmails() {
  const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbyFPxd3UelHmFuh4fqQC7YPLpVk44rorubWx_My_0S2OV7Il4GlJC1wd7rq8aVKJKpKNg/exec';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(FAILED.SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet "' + FAILED.SHEET_NAME + '" not found.');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < FAILED.START_ROW) {
    return { status: 'No applicants found', count: 0 };
  }

  const data = sheet.getRange(FAILED.START_ROW, 1, lastRow - FAILED.START_ROW + 1, FAILED.COL_REGENERATE).getValues();
  let emailCount = 0;
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const regenerateVal = row[FAILED.COL_REGENERATE - 1];
    const shouldSend = regenerateVal === true || String(regenerateVal).toLowerCase() === 'true';
    if (!shouldSend) continue;

    const applicantName = row[0];
    const email = row[FAILED.COL_EMAIL - 1];
    const driveLink = row[FAILED.COL_LINK - 1];
    const position = row[6];
    const office = row[7];
    const statusCell = sheet.getRange(FAILED.START_ROW + i, FAILED.COL_STATUS);

    if (!applicantName || applicantName.toString().trim() === '') {
      statusCell.setValue('Not sent - missing name (' + now + ')');
      sheet.getRange(FAILED.START_ROW + i, FAILED.COL_REGENERATE).setValue(false);
      continue;
    }

    if (!email || email.toString().trim() === '' || !driveLink || driveLink.toString().trim() === '') {
      statusCell.setValue('Not sent - missing email or link (' + now + ')');
      sheet.getRange(FAILED.START_ROW + i, FAILED.COL_REGENERATE).setValue(false);
      continue;
    }

    const subject = 'JOB APPLICATION UPDATE';
    const body = 'Dear Applicant,\n\n' +
      'Good day!\n\n' +
      'Please see attached file regarding your application.\n\n' +
      'Link: ' + driveLink + '\n\n' +
      'Best regards,\n' +
      'DOJ RPO V - Human Resource Unit';

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
      // Log the sent letter using the reference from Column G and H
      logSentLetter('LETTER - FAILED', position || '', office || '', applicantName || '');
      emailCount++;
      sheet.getRange(FAILED.START_ROW + i, FAILED.COL_REGENERATE).setValue(false);
    } else {
      statusCell.setValue('Error: Proxy failed (' + now + ')');
    }
  }

  return { status: 'Selected emails processed', count: emailCount };
}

function failedGeneratePDFs(targetFolderId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FAILED.SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + FAILED.SHEET_NAME + '" not found!');

    const data = sheet.getDataRange().getDisplayValues();
    
    // Create a copy of the header row and add our two new combined fields
    const header = [...data[0]];
    header.push("RECIPIENT_BLOCK", "DEAR_BLOCK");

    const rawRows = data.slice(1).filter(row => row[0] && row[0].toString().trim() !== '');

    // --- INTEGRATED NOTE: Combine the requested columns per row ---
    const rows = rawRows.map(row => {
      const newRow = [...row];
      
      // Pull values using our mapped constants (subtracting 1 for 0-indexed arrays)
      const upperSalutation = newRow[FAILED.COL_UPPER_SALUTATION - 1] || "";
      const upperFullName = newRow[FAILED.COL_UPPER_FULLNAME - 1] || "";
      const properSalutation = newRow[FAILED.COL_PROPER_SALUTATION - 1] || "";
      const properLastName = newRow[FAILED.COL_PROPER_LASTNAME - 1] || "";

      // Combine Column I and J for the "Recipient" part
      const recipientBlock = (upperSalutation + " " + upperFullName).trim();
      
      // Combine Column K and L for the "Dear" part
      const dearBlock = (properSalutation + " " + properLastName).trim();

      // Push them to the end of the array so processPDFBatch can map them to the template tags
      newRow.push(recipientBlock, dearBlock);
      return newRow;
    });

    // Sort rows alphabetically matching the layout configuration
    rows.sort((a, b) => {
      const lastNameA = String(a[0] || '').trim().toLowerCase();
      const lastNameB = String(b[0] || '').trim().toLowerCase();
      if (lastNameA !== lastNameB) return lastNameA.localeCompare(lastNameB);
      return String(a[1] || '').trim().toLowerCase().localeCompare(String(b[1] || '').trim().toLowerCase());
    });

    const templateFile = DriveApp.getFileById(FAILED.TEMPLATE_ID);
    const destinationFolder = DriveApp.getFolderById(targetFolderId);
    
    // Use batch processing with key for Failed
    const batchKey = 'failed_pdf_generation_' + SpreadsheetApp.getActiveSpreadsheet().getId();
    
    // Since we appended the combined blocks to 'rows' and 'header', processPDFBatch will handle them automatically
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
    throw new Error('Error generating failed PDFs: ' + e.message);
  }
}

function failedGenerateLinks() {
  try {
    const settings = PropertiesService.getDocumentProperties();
    const folderId = settings.getProperty('failedSubFolderId');
    if (!folderId) {
      throw new Error('Failed subfolder not found. Please run Step 2 first.');
    }

    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByType(MimeType.PDF);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FAILED.SHEET_NAME);
    if (!sheet) {
      throw new Error('Sheet "' + FAILED.SHEET_NAME + '" not found.');
    }

    const fileData = [];
    while (files.hasNext()) {
      const file = files.next();
      fileData.push({ name: file.getName(), url: file.getUrl() });
    }

    fileData.sort((a, b) => a.name.toLowerCase().localeCompare(b.name.toLowerCase()));

    const links = fileData.map(item => [item.url]);
    if (links.length > 0) {
      sheet.getRange(FAILED.START_ROW, FAILED.COL_LINK, links.length, 1).setValues(links);
    }

    return 'Successfully generated and inserted ' + links.length + ' Google Drive links into N.';
  } catch (e) {
    throw new Error('Error generating Failed Drive links: ' + e.message);
  }
}

function failedGetFolderUrl() {
  try {
    const settings = PropertiesService.getDocumentProperties();
    const folderId = settings.getProperty('failedSubFolderId');
    if (!folderId) {
      throw new Error('Failed subfolder not found. Please run the process first.');
    }
    return DriveApp.getFolderById(folderId).getUrl();
  } catch (e) {
    throw new Error('Error retrieving failed folder URL: ' + e.message);
  }
}

function failedBackupSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(FAILED.SHEET_NAME);
    if (!sourceSheet) {
      throw new Error('Sheet "' + FAILED.SHEET_NAME + '" not found.');
    }

    const settings = PropertiesService.getDocumentProperties();
    const folderId = settings.getProperty('failedSubFolderId');
    if (!folderId) {
      throw new Error('Failed subfolder not found. Please run Step 2 first.');
    }

    const folder = DriveApp.getFolderById(folderId);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    const backupFileName = 'LETTER - FAILED_' + timestamp + '.csv';

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
      message: 'Backup successful! LETTER - FAILED has been saved to the Failed folder.',
      folderUrl: folder.getUrl()
    };
  } catch (e) {
    throw new Error('Error backing up Failed sheet: ' + e.message);
  }
}

function failedRunCompleteProcess() {
  try {
    const folderIds = getFailedPositionFolder();
    const pdfResult = failedGeneratePDFs(folderIds.failedSubFolderId);
    
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
    throw new Error('Error in failed complete process: ' + e.message);
  }
}