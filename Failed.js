// ==========================================
// FAILED - Letter Generation & Backup Workflow
// ==========================================

const FAILED = {
  SHEET_NAME: 'LETTER - FAILED',
  TEMPLATE_ID: '1-08ufjwnG0bCV9nrZ6LbOwA_DDBlr8nfUNZYncwFGP8',
  COL_LAST_NAME: 1,
  COL_FIRST_NAME: 2,
  COL_EMAIL: 6,              // Column F
  COL_UPPER_SALUTATION: 9,   // Column I
  COL_UPPER_FULLNAME: 10,    // Column J
  COL_PROPER_SALUTATION: 11, // Column K
  COL_PROPER_LASTNAME: 12,   // Column L
  COL_EMAIL_DATE: 13,        // Column M
  COL_LINK: 14,              // Column N
  COL_STATUS: 15,            // Column O
  COL_REGENERATE: 16,        // Column P
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
  const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxJpyg6KPFUMxeHSOdOVnVe4WyN6JssT9DhoufEn2pE7vIp02joOQ6jZVD-FwZCLKW7FQ/exec";

  try {
    PropertiesService.getDocumentProperties().deleteProperty('cancel_failed_send');
    CacheService.getDocumentCache().remove('cancel_failed_send');
    
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
    let cancelled = false;

    for (let i = 0; i < data.length; i++) {
      if (shouldCancelFailedSend()) {
        console.log('Cancellation requested for failed email sending');
        cancelled = true;
        break;
      }

      const row = data[i];
      const applicantName = row[0]; // Column A
      const applicantLName = row[11]; // Column L
      const salutation = row[10]; // Column K
      const email = row[FAILED.COL_EMAIL - 1]; // Column F
      const driveLink = row[FAILED.COL_LINK - 1]; // Column N
      
      const position = row[6]; // Index 6 is Column G (POSITION EXT)
      const office = row[7];   // Index 7 is Column H (ASSIGNED OFFICE)
      
      const statusCell = sheet.getRange(FAILED.START_ROW + i, FAILED.COL_STATUS);

      if (!applicantName || applicantName.toString().trim() === '') break;

      const currentStatus = row[FAILED.COL_STATUS - 1];
      if (currentStatus && String(currentStatus).startsWith('Sent')) continue;

      if (!email || email.toString().trim() === '' || !driveLink || driveLink.toString().trim() === '') {
        statusCell.setValue('Not sent - missing email or link (' + now + ')');
        continue;
      }

      const subject = 'Job Application Update ' + '[' + position + ']';
      const body = 'Dear ' + salutation + ' ' + applicantLName + ',\n\n' +
        'Good day!\n\n' +
        'Please see attached file regarding your application.\n\n' +
        'Link: ' + driveLink + '\n\n' +
        'Kindly acknowledge receipt of this email.\n\n' +
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
        logSentLetter('LETTER - FAILED', position || '', office || '', applicantName || '');
        emailCount++;
      } else {
        statusCell.setValue('Error: Proxy failed (' + now + ')');
      }

      if (shouldCancelFailedSend()) {
        console.log('Cancellation requested for failed email sending before sleep');
        cancelled = true;
        break;
      }

      Utilities.sleep(1500);
    }

    return { status: cancelled ? 'Cancelled' : 'Emails sent', count: emailCount, cancelled: cancelled };
  } catch (e) {
    throw new Error('Error sending email notifications: ' + e.message);
  }
}

function failedGenerateIndividualPDFs() {
  try {
    PropertiesService.getDocumentProperties().deleteProperty('cancel_failed_run');
    CacheService.getDocumentCache().remove('cancel_failed_run');
    
    const folderIds = getFailedPositionFolder();
    const targetFolderId = folderIds.failedSubFolderId;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FAILED.SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + FAILED.SHEET_NAME + '" not found.');

    const data = sheet.getDataRange().getDisplayValues();
    const header = [...data[0]];
    header.push('RECIPIENT_BLOCK', 'DEAR_BLOCK');

    const selectedRows = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const regenerateVal = row[FAILED.COL_REGENERATE - 1];
      const shouldGenerate = regenerateVal === true || String(regenerateVal).toLowerCase() === 'true';
      if (!shouldGenerate) continue;
      if (!row[0] || row[0].toString().trim() === '') continue;

      const upperSalutation = row[FAILED.COL_UPPER_SALUTATION - 1] || '';
      const upperFullName = row[FAILED.COL_UPPER_FULLNAME - 1] || '';
      const properSalutation = row[FAILED.COL_PROPER_SALUTATION - 1] || '';
      const properLastName = row[FAILED.COL_PROPER_LASTNAME - 1] || '';
      const recipientBlock = (upperSalutation + ' ' + upperFullName).trim();
      const dearBlock = (properSalutation + ' ' + properLastName).trim();

      const rowCopy = [...row];
      rowCopy.push(recipientBlock, dearBlock);
      selectedRows.push({ row: rowCopy, rowIndex: FAILED.START_ROW + i - 1 });
    }

    if (selectedRows.length === 0) {
      throw new Error('No items checked in REGENERATE column (P). Please check at least one checkbox to proceed.');
    }

    const templateFile = DriveApp.getFileById(FAILED.TEMPLATE_ID);
    const destinationFolder = DriveApp.getFolderById(targetFolderId);
    const processed = [];
    let cancelled = false;

    for (let k = 0; k < selectedRows.length; k++) {
      if (shouldCancelFailedRun()) {
        console.log('Cancellation requested for individual failed PDF generation');
        cancelled = true;
        break;
      }

      const { row, rowIndex } = selectedRows[k];
      const lastName = String(row[FAILED.COL_LAST_NAME - 1] || '').trim();
      const firstName = String(row[FAILED.COL_FIRST_NAME - 1] || '').trim();
      const fileName = (lastName || 'Applicant') + (firstName ? (', ' + firstName) : '');

      try {
        const copy = templateFile.makeCopy(fileName, destinationFolder);

        if (shouldCancelFailedRun()) {
          copy.setTrashed(true);
          cancelled = true;
          break;
        }

        const doc = DocumentApp.openById(copy.getId());
        const body = doc.getBody();

        header.forEach((label, j) => {
          body.replaceText('{{' + label + '}}', row[j]);
        });

        doc.saveAndClose();

        if (shouldCancelFailedRun()) {
          copy.setTrashed(true);
          cancelled = true;
          break;
        }

        const pdfBlob = copy.getAs(MimeType.PDF);
        const pdfFile = destinationFolder.createFile(pdfBlob).setName(fileName + '.pdf');
        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        copy.setTrashed(true);

        const pdfUrl = pdfFile.getUrl();
        sheet.getRange(rowIndex, FAILED.COL_LINK).setValue(pdfUrl);
        sheet.getRange(rowIndex, FAILED.COL_REGENERATE).setValue(false);
        processed.push(fileName);
      } catch (itemError) {
        console.log('Error generating failed PDF for row ' + rowIndex + ': ' + itemError.message);
      }
    }

    return { success: true, count: processed.length, applicants: processed, cancelled: cancelled };
  } catch (e) {
    throw new Error('Error generating selected failed PDFs: ' + e.message);
  }
}

function failedSendSelectedEmails() {
  const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxJpyg6KPFUMxeHSOdOVnVe4WyN6JssT9DhoufEn2pE7vIp02joOQ6jZVD-FwZCLKW7FQ/exec";

  try {
    PropertiesService.getDocumentProperties().deleteProperty('cancel_failed_send');
    CacheService.getDocumentCache().remove('cancel_failed_send');
    
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

    const hasSelected = data.some(row => {
      const val = row[FAILED.COL_REGENERATE - 1];
      return val === true || String(val).toLowerCase() === 'true';
    });
    if (!hasSelected) {
      throw new Error('No items checked in REGENERATE column. Please check at least one checkbox to proceed.');
    }

    let emailCount = 0;
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    let cancelled = false;

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const regenerateVal = row[FAILED.COL_REGENERATE - 1];
      const shouldSend = regenerateVal === true || String(regenerateVal).toLowerCase() === 'true';
      if (!shouldSend) continue;

      if (shouldCancelFailedSend()) {
        console.log('Cancellation requested for selected failed email sending');
        cancelled = true;
        break;
      }

      const applicantName = row[0];
      const applicantLName = row[11]; // Column L
      const salutation = row[10]; // Column K
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

      const subject = 'Job Application Update ' + '[' + position + ']';
      const body = 'Dear ' + salutation + ' ' + applicantLName + ',\n\n' +
        'Good day!\n\n' +
        'Please see attached file regarding your application.\n\n' +
        'Link: ' + driveLink + '\n\n' +
        'Kindly acknowledge receipt of this email.\n\n' +
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
        logSentLetter('LETTER - FAILED', position || '', office || '', applicantName || '');
        emailCount++;
        sheet.getRange(FAILED.START_ROW + i, FAILED.COL_REGENERATE).setValue(false);
      } else {
        statusCell.setValue('Error: Proxy failed (' + now + ')');
      }

      if (shouldCancelFailedSend()) {
        console.log('Cancellation requested for selected failed email sending before sleep');
        cancelled = true;
        break;
      }

      Utilities.sleep(1500);
    }

    return { status: cancelled ? 'Cancelled' : 'Selected emails processed', count: emailCount, cancelled: cancelled };
  } catch (e) {
    throw new Error('Error sending selected emails: ' + e.message);
  }
}

function failedGeneratePDFs(targetFolderId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FAILED.SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + FAILED.SHEET_NAME + '" not found!');

    const data = sheet.getDataRange().getDisplayValues();
    const header = [...data[0]];
    header.push("RECIPIENT_BLOCK", "DEAR_BLOCK");

    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowIndex = FAILED.START_ROW + i - 1;
      if (row[0] && row[0].toString().trim() !== "") {
        const upperSalutation = row[FAILED.COL_UPPER_SALUTATION - 1] || "";
        const upperFullName = row[FAILED.COL_UPPER_FULLNAME - 1] || "";
        const properSalutation = row[FAILED.COL_PROPER_SALUTATION - 1] || "";
        const properLastName = row[FAILED.COL_PROPER_LASTNAME - 1] || "";
        const recipientBlock = (upperSalutation + " " + upperFullName).trim();
        const dearBlock = (properSalutation + " " + properLastName).trim();

        const rowCopy = [...row];
        rowCopy.push(recipientBlock, dearBlock);
        rows.push({ row: rowCopy, rowIndex: rowIndex });
      }
    }

    if (rows.length === 0) {
      return {
        success: true,
        count: 0,
        applicants: [],
        completed: true,
        message: 'No eligible rows found for PDF generation.'
      };
    }

    rows.sort((a, b) => {
      const lastNameA = String(a.row[0] || '').trim().toLowerCase();
      const lastNameB = String(b.row[0] || '').trim().toLowerCase();
      if (lastNameA !== lastNameB) return lastNameA.localeCompare(lastNameB);
      return String(a.row[1] || '').trim().toLowerCase().localeCompare(String(b.row[1] || '').trim().toLowerCase());
    });

    const templateFile = DriveApp.getFileById(FAILED.TEMPLATE_ID);
    const destinationFolder = DriveApp.getFolderById(targetFolderId);
    
    const batchKey = 'failed_pdf_generation_' + SpreadsheetApp.getActiveSpreadsheet().getId();
    const batchResult = processFailedPDFBatch(batchKey, rows, header, templateFile, destinationFolder, 20, sheet);

    let returnMessage = batchResult.message;
    
    if (batchResult.status === 'cancelled') {
      returnMessage = 'Process cancelled. ' + batchResult.totalProcessed + ' PDFs generated.';
    } else if (!batchResult.completed) {
      returnMessage += '\n\nTo continue processing remaining applicants (total: ' + batchResult.totalRows + '), run this step again.';
    } else {
      clearBatchState(batchKey);
      returnMessage = 'PDF generation completed! ' + batchResult.totalProcessed + ' PDFs generated.';
    }

    return {
      success: true,
      count: batchResult.totalProcessed,
      applicants: batchResult.allApplicants,
      completed: batchResult.completed || batchResult.status === 'cancelled',
      cancelled: batchResult.status === 'cancelled',
      message: returnMessage
    };
  } catch (e) {
    throw new Error('Error generating failed PDFs: ' + e.message);
  }
}

function processFailedPDFBatch(batchKey, rows, header, templateFile, destinationFolder, batchSize, sheet) {
  let state = getBatchState(batchKey);

  if (!state) {
    PropertiesService.getDocumentProperties().deleteProperty('cancel_failed_run');
    CacheService.getDocumentCache().remove('cancel_failed_run');
    state = initializeBatchProcessing(batchKey, rows.length);
  }

  const startIndex = state.currentIndex;
  const endIndex = Math.min(startIndex + batchSize, rows.length);
  const startTime = new Date().getTime();
  const timeLimit = 5 * 60 * 1000;

  let processedInThisBatch = 0;
  const newApplicants = [];

  try {
    for (let i = startIndex; i < endIndex; i++) {
      const elapsedTime = new Date().getTime() - startTime;
      if (elapsedTime > timeLimit) {
        console.log('Time limit approaching, saving progress...');
        break;
      }

      if (shouldCancelFailedRun()) {
        console.log('Cancellation requested for failed PDF generation');
        state.status = 'cancelled';
        break;
      }

      const rowObj = rows[i];
      const row = rowObj.row;
      const rowIndex = rowObj.rowIndex;
      const lastName = String(row[0] || "").trim();
      const firstName = String(row[1] || "").trim();
      const fileName = lastName + ", " + firstName;

      try {
        const copy = templateFile.makeCopy(fileName, destinationFolder);

        if (shouldCancelFailedRun()) {
          copy.setTrashed(true);
          state.status = 'cancelled';
          break;
        }

        const doc = DocumentApp.openById(copy.getId());
        const body = doc.getBody();

        header.forEach((label, j) => {
          body.replaceText('{{' + label + '}}', row[j]);
        });

        doc.saveAndClose();

        if (shouldCancelFailedRun()) {
          copy.setTrashed(true);
          state.status = 'cancelled';
          break;
        }

        const pdfBlob = copy.getAs(MimeType.PDF);
        const pdfFile = destinationFolder.createFile(pdfBlob).setName(fileName + ".pdf");
        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        copy.setTrashed(true);

        const pdfUrl = pdfFile.getUrl();
        sheet.getRange(rowIndex, FAILED.COL_LINK).setValue(pdfUrl);

        state.currentIndex = i + 1;
        state.completedCount++;
        state.processedApplicants.push(lastName + ', ' + firstName);
        newApplicants.push(lastName + ', ' + firstName);
        processedInThisBatch++;
      } catch (itemError) {
        console.log('Error processing ' + fileName + ': ' + itemError.message);
        state.currentIndex = i + 1;
      }
    }
  } catch (e) {
    console.log('Batch processing error: ' + e.message);
  }

  const isCompleted = state.currentIndex >= rows.length || state.status === 'cancelled';
  if (state.status === 'cancelled') {
    clearBatchState(batchKey);
  } else if (isCompleted) {
    state.status = 'completed';
    clearBatchState(batchKey);
  } else {
    updateBatchState(batchKey, state);
  }

  return {
    completed: isCompleted,
    processed: processedInThisBatch,
    totalProcessed: state.completedCount,
    totalRows: rows.length,
    applicants: newApplicants,
    allApplicants: state.processedApplicants,
    message: state.status === 'cancelled'
      ? 'Process cancelled. Total: ' + state.completedCount + ' / ' + rows.length
      : processedInThisBatch + ' applicants processed. Total: ' + state.completedCount + ' / ' + rows.length,
    status: state.status
  };
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

// ==========================================
// CANCELLATION HELPERS FOR FAILED
// ==========================================

function shouldCancelFailedRun() {
  const cache = CacheService.getDocumentCache();
  if (cache.get('cancel_failed_run') === 'true') return true;
  return PropertiesService.getDocumentProperties().getProperty('cancel_failed_run') === 'true';
}

function shouldCancelFailedSend() {
  const cache = CacheService.getDocumentCache();
  if (cache.get('cancel_failed_send') === 'true') return true;
  return PropertiesService.getDocumentProperties().getProperty('cancel_failed_send') === 'true';
}

function cancelFailedRun() {
  CacheService.getDocumentCache().put('cancel_failed_run', 'true', 21600);
  PropertiesService.getDocumentProperties().setProperty('cancel_failed_run', 'true');
}

function cancelFailedSend() {
  CacheService.getDocumentCache().put('cancel_failed_send', 'true', 21600);
  PropertiesService.getDocumentProperties().setProperty('cancel_failed_send', 'true');
}

function clearCancelFailedFlags() {
  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty('cancel_failed_run');
  props.deleteProperty('cancel_failed_send');
  const cache = CacheService.getDocumentCache();
  cache.remove('cancel_failed_run');
  cache.remove('cancel_failed_send');
}