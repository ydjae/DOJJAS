// ==========================================
// FINAL NOTICE - Letter Generation & Email Workflow
// ==========================================

const FINAL_NOTICE = {
  SHEET_NAME: 'LETTER - FINAL NOTICE',
  TEMPLATE_ID: '1MMPN0LssUniSQy86Q53df4SM7KT0bEtFoBOL_tm2pSY',
  COL_LAST_NAME: 1,       // Column A
  COL_FIRST_NAME: 2,      // Column B
  COL_ADDRESS: 5,         // Column E
  COL_EMAIL: 7,           // Column G
  COL_POSITION_EXTRACTED: 8, // Column H
  COL_ASSIGNED_OFFICE: 9, // Column I
  COL_SALUTATION: 10,     // Column J
  COL_UPPERCASE_NAME: 11, // Column K
  COL_PROPER_SALUTATION: 12, // Column L
  COL_PROPER_LN: 13,      // Column M
  COL_EMAIL_DATE: 15,     // Column O
  COL_LINK: 16,           // Column P
  COL_STATUS: 17,         // Column Q
  COL_REGENERATE: 18,     // Column R
  START_ROW: 2
};

/**
 * Check if required input columns have data in the FINAL NOTICE tab
 */
function finalNoticeCheckColumns() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FINAL_NOTICE.SHEET_NAME);
    if (!sheet) {
      return { hasData: false, message: 'Sheet not found' };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < FINAL_NOTICE.START_ROW) {
      return { hasData: false };
    }

    const dataRange = sheet.getRange(
      FINAL_NOTICE.START_ROW,
      1,
      lastRow - FINAL_NOTICE.START_ROW + 1,
      FINAL_NOTICE.COL_EMAIL_DATE // Column O is 15
    ).getValues();

    const colIndexesToCheck = [15]; // Column numbers: O
    const colNames = {
      15: 'O (Email Date)'
    };

    for (let i = 0; i < dataRange.length; i++) {
      const rowData = dataRange[i];
      const valA = rowData[0]; // Column A

      if (valA && valA.toString().trim() !== '') {
        for (let j = 0; j < colIndexesToCheck.length; j++) {
          const colNum = colIndexesToCheck[j];
          const cellValue = rowData[colNum - 1]; // 0-indexed
          if (!cellValue || cellValue.toString().trim() === '') {
            const rowNum = FINAL_NOTICE.START_ROW + i;
            return {
              hasData: false,
              message: 'Missing data in column ' + colNames[colNum] + ' for applicant at row ' + rowNum
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
 * Create main folder and Final Notice subfolder
 */
function finalNoticeCreateFolders() {
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

    let mainFolderId = props.getProperty('finalNoticeMainFolderId');
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
      props.setProperty('finalNoticeMainFolderId', mainFolderId);
    }

    let finalNoticeSubFolderId = props.getProperty('finalNoticeSubFolderId');
    let finalNoticeSubFolder = null;

    if (finalNoticeSubFolderId) {
      try {
        const tempSubFolder = DriveApp.getFolderById(finalNoticeSubFolderId);
        if (tempSubFolder.getParents().hasNext() && tempSubFolder.getParents().next().getId() === mainFolderId) {
          finalNoticeSubFolder = tempSubFolder;
        }
      } catch (e) {
        finalNoticeSubFolder = null;
      }
    }

    if (!finalNoticeSubFolder) {
      const existingSubFolders = mainFolder.getFoldersByName('Final Notice');
      if (existingSubFolders.hasNext()) {
        finalNoticeSubFolder = existingSubFolders.next();
      } else {
        finalNoticeSubFolder = mainFolder.createFolder('Final Notice');
      }
      finalNoticeSubFolderId = finalNoticeSubFolder.getId();
      props.setProperty('finalNoticeSubFolderId', finalNoticeSubFolderId);
    }

    return {
      mainFolderId: mainFolderId,
      finalNoticeSubFolderId: finalNoticeSubFolderId,
      folderUrl: finalNoticeSubFolder.getUrl()
    };
  } catch (e) {
    throw new Error('Error creating folders: ' + e.message);
  }
}

/**
 * Generate PDFs from sheet and return list of processed applicants
 */
function finalNoticeGeneratePDFs(targetFolderId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FINAL_NOTICE.SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + FINAL_NOTICE.SHEET_NAME + '" not found!');

    const data = sheet.getDataRange().getDisplayValues();
    const header = data[0];

    // Sort rows alphabetically matching layout
    const rawRows = data.slice(1).filter(row => row[0] && row[0].toString().trim() !== '');

    const rows = rawRows.map((row, index) => {
      return {
        row: row,
        rowIndex: FINAL_NOTICE.START_ROW + index
      };
    });

    rows.sort((a, b) => {
      const lastNameA = String(a.row[0] || '').trim().toLowerCase();
      const lastNameB = String(b.row[0] || '').trim().toLowerCase();
      if (lastNameA !== lastNameB) return lastNameA.localeCompare(lastNameB);
      return String(a.row[1] || '').trim().toLowerCase().localeCompare(String(b.row[1] || '').trim().toLowerCase());
    });

    const templateFile = DriveApp.getFileById(FINAL_NOTICE.TEMPLATE_ID);
    const destinationFolder = DriveApp.getFolderById(targetFolderId);

    const batchKey = 'finalNotice_pdf_generation_' + ss.getId();
    const batchResult = processFinalNoticePDFBatch(batchKey, rows, header, templateFile, destinationFolder, 20, sheet);

    let returnMessage = batchResult.message;
    if (!batchResult.completed) {
      returnMessage += '\n\nTo continue processing remaining applicants (total: ' + batchResult.totalRows + '), run this step again.';
    } else {
      clearBatchState(batchKey);
      returnMessage = 'PDF generation completed! ' + batchResult.totalProcessed + ' PDFs generated.';
    }

    return {
      success: true,
      count: batchResult.totalProcessed,
      applicants: batchResult.allApplicants,
      completed: batchResult.completed,
      cancelled: batchResult.status === 'cancelled',
      message: returnMessage
    };
  } catch (e) {
    throw new Error('Error generating PDFs: ' + e.message);
  }
}

function processFinalNoticePDFBatch(batchKey, rows, header, templateFile, destinationFolder, batchSize, sheet) {
  let state = getBatchState(batchKey);

  if (!state) {
    PropertiesService.getDocumentProperties().deleteProperty('cancel_finalNotice_run');
    CacheService.getDocumentCache().remove('cancel_finalNotice_run');
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

      if (shouldCancelFinalNoticeRun()) {
        console.log('Cancellation requested for PDF generation');
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

        if (shouldCancelFinalNoticeRun()) {
          copy.setTrashed(true);
          state.status = 'cancelled';
          break;
        }

        const doc = DocumentApp.openById(copy.getId());
        const body = doc.getBody();

        header.forEach((label, j) => {
          body.replaceText('{{' + label + '}}', row[j]);
        });

        // Explicit placeholder overrides matching requested columns (E, H, I, G)
        body.replaceText('{{ADDRESS}}', row[FINAL_NOTICE.COL_ADDRESS - 1] || '');
        body.replaceText('{{POSITION EXTRACTED}}', row[FINAL_NOTICE.COL_POSITION_EXTRACTED - 1] || '');
        body.replaceText('{{ASSIGNED OFFICE}}', row[FINAL_NOTICE.COL_ASSIGNED_OFFICE - 1] || '');
        body.replaceText('{{EMAIL}}', row[FINAL_NOTICE.COL_EMAIL - 1] || '');

        doc.saveAndClose();

        if (shouldCancelFinalNoticeRun()) {
          copy.setTrashed(true);
          state.status = 'cancelled';
          break;
        }

        const pdfBlob = copy.getAs(MimeType.PDF);
        const pdfFile = destinationFolder.createFile(pdfBlob).setName(fileName + ".pdf");
        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        copy.setTrashed(true);

        const pdfUrl = pdfFile.getUrl();
        sheet.getRange(rowIndex, FINAL_NOTICE.COL_LINK).setValue(pdfUrl);

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

/**
 * Generate PDFs only for rows with the REGENERATE checkbox checked (Column R).
 */
function finalNoticeGenerateIndividualPDFs() {
  try {
    PropertiesService.getDocumentProperties().deleteProperty('cancel_finalNotice_run');
    CacheService.getDocumentCache().remove('cancel_finalNotice_run');
    const props = PropertiesService.getDocumentProperties();
    let folderId = props.getProperty('finalNoticeSubFolderId');
    let destinationFolder = null;

    if (folderId) {
      try {
        destinationFolder = DriveApp.getFolderById(folderId);
      } catch (folderError) {
        console.log('Stored Final Notice subfolder ID invalid. Recreating folder: ' + folderError.message);
        const created = finalNoticeCreateFolders();
        folderId = created.finalNoticeSubFolderId;
        destinationFolder = DriveApp.getFolderById(folderId);
      }
    }

    if (!destinationFolder) {
      const created = finalNoticeCreateFolders();
      folderId = created.finalNoticeSubFolderId;
      destinationFolder = DriveApp.getFolderById(folderId);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FINAL_NOTICE.SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + FINAL_NOTICE.SHEET_NAME + '" not found.');

    const data = sheet.getDataRange().getDisplayValues();
    const header = data[0] || [];
    const rowsToProcess = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowIndex = FINAL_NOTICE.START_ROW + i - 1;
      const regenerateVal = row[FINAL_NOTICE.COL_REGENERATE - 1];
      const shouldProcess = regenerateVal === true || String(regenerateVal).toLowerCase() === 'true';
      if (!shouldProcess) continue;
      if (!row[0] || row[0].toString().trim() === '') continue;
      rowsToProcess.push({ row: row, rowIndex: rowIndex });
    }

    if (rowsToProcess.length === 0) {
      throw new Error('No items checked in REGENERATE column (R). Please check at least one checkbox to proceed.');
    }

    let templateFile = DriveApp.getFileById(FINAL_NOTICE.TEMPLATE_ID);
    const processed = [];
    let cancelled = false;

    for (let k = 0; k < rowsToProcess.length; k++) {
      if (shouldCancelFinalNoticeRun()) {
        console.log('Cancellation requested for individual PDF generation');
        cancelled = true;
        break;
      }

      const rowObj = rowsToProcess[k];
      const row = rowObj.row;
      const rowIndex = rowObj.rowIndex;
      const lastName = String(row[0] || "").trim();
      const firstName = String(row[1] || "").trim();
      const fileName = (lastName || 'Applicant') + (firstName ? (', ' + firstName) : '');

      try {
        const copy = templateFile.makeCopy(fileName, destinationFolder);

        if (shouldCancelFinalNoticeRun()) {
          copy.setTrashed(true);
          cancelled = true;
          break;
        }

        const doc = DocumentApp.openById(copy.getId());
        const body = doc.getBody();

        header.forEach((label, j) => {
          body.replaceText('{{' + label + '}}', row[j]);
        });

        // Explicit placeholder overrides matching requested columns (E, H, I, G)
        body.replaceText('{{ADDRESS}}', row[FINAL_NOTICE.COL_ADDRESS - 1] || '');
        body.replaceText('{{POSITION EXTRACTED}}', row[FINAL_NOTICE.COL_POSITION_EXTRACTED - 1] || '');
        body.replaceText('{{ASSIGNED OFFICE}}', row[FINAL_NOTICE.COL_ASSIGNED_OFFICE - 1] || '');
        body.replaceText('{{EMAIL}}', row[FINAL_NOTICE.COL_EMAIL - 1] || '');

        doc.saveAndClose();

        if (shouldCancelFinalNoticeRun()) {
          copy.setTrashed(true);
          cancelled = true;
          break;
        }

        const pdfBlob = copy.getAs(MimeType.PDF);
        const pdfFile = destinationFolder.createFile(pdfBlob).setName(fileName + ".pdf");
        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        copy.setTrashed(true);

        const pdfUrl = pdfFile.getUrl();
        sheet.getRange(rowIndex, FINAL_NOTICE.COL_LINK).setValue(pdfUrl);

        processed.push(lastName + (firstName ? (', ' + firstName) : ''));
      } catch (itemError) {
        console.log('Error generating individual PDF for row ' + rowIndex + ': ' + itemError.message);
      }
    }

    return { success: true, count: processed.length, applicants: processed, cancelled: cancelled };
  } catch (e) {
    throw new Error('Error generating individual PDFs: ' + e.message);
  }
}

/**
 * Generate Google Drive links and insert into spreadsheet
 */
function finalNoticeGenerateLinks() {
  try {
    const settings = PropertiesService.getDocumentProperties();
    const folderId = settings.getProperty('finalNoticeSubFolderId');

    if (!folderId) {
      throw new Error('Final Notice subfolder not found. Please run Step 2 first.');
    }

    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByType(MimeType.PDF);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FINAL_NOTICE.SHEET_NAME);

    if (!sheet) {
      throw new Error('Sheet "' + FINAL_NOTICE.SHEET_NAME + '" not found.');
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
      sheet.getRange(FINAL_NOTICE.START_ROW, FINAL_NOTICE.COL_LINK, links.length, 1).setValues(links);
    }

    return 'Successfully generated and inserted ' + links.length + ' Google Drive links into Column P.';
  } catch (e) {
    throw new Error('Error generating Drive links: ' + e.message);
  }
}

/**
 * Get the Final Notice folder URL
 */
function finalNoticeGetFolderUrl() {
  try {
    const settings = PropertiesService.getDocumentProperties();
    const folderId = settings.getProperty('finalNoticeSubFolderId');

    if (!folderId) {
      throw new Error('Final Notice subfolder not found. Please run the process first.');
    }

    const folder = DriveApp.getFolderById(folderId);
    return folder.getUrl();
  } catch (e) {
    throw new Error('Error retrieving folder URL: ' + e.message);
  }
}

/**
 * Backup the Letter - Final Notice sheet to the Final Notice folder
 */
function finalNoticeBackupSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(FINAL_NOTICE.SHEET_NAME);

    if (!sourceSheet) {
      throw new Error('Sheet "' + FINAL_NOTICE.SHEET_NAME + '" not found.');
    }

    const settings = PropertiesService.getDocumentProperties();
    const finalNoticeSubFolderId = settings.getProperty('finalNoticeSubFolderId');

    if (!finalNoticeSubFolderId) {
      throw new Error('Final Notice subfolder not found. Please run Step 2 first.');
    }

    const finalNoticeFolder = DriveApp.getFolderById(finalNoticeSubFolderId);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    const backupFileName = 'LETTER - FINAL NOTICE_' + timestamp + '.csv';

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
    finalNoticeFolder.createFile(backupBlob);

    return {
      message: 'Backup successful! LETTER - FINAL NOTICE has been saved to the Final Notice folder.',
      folderUrl: finalNoticeFolder.getUrl()
    };
  } catch (e) {
    throw new Error('Error backing up sheet: ' + e.message);
  }
}

function finalNoticeSendEmails() {
  const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxJpyg6KPFUMxeHSOdOVnVe4WyN6JssT9DhoufEn2pE7vIp02joOQ6jZVD-FwZCLKW7FQ/exec";

  try {
    PropertiesService.getDocumentProperties().deleteProperty('cancel_finalNotice_send');
    CacheService.getDocumentCache().remove('cancel_finalNotice_send');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FINAL_NOTICE.SHEET_NAME);

    if (!sheet) throw new Error('Sheet "' + FINAL_NOTICE.SHEET_NAME + '" not found.');

    const lastRow = sheet.getLastRow();
    if (lastRow < FINAL_NOTICE.START_ROW) {
      return { status: 'No applicants found', count: 0 };
    }

    const data = sheet.getRange(FINAL_NOTICE.START_ROW, 1, lastRow - FINAL_NOTICE.START_ROW + 1, FINAL_NOTICE.COL_STATUS).getValues();
    let emailCount = 0;
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    let cancelled = false;

    for (let i = 0; i < data.length; i++) {
      if (shouldCancelFinalNoticeSend()) {
        console.log('Cancellation requested for email sending');
        cancelled = true;
        break;
      }

      const row = data[i];
      const applicantName = row[0];
      const applicantLName = row[FINAL_NOTICE.COL_PROPER_LN - 1];
      const salutation = row[FINAL_NOTICE.COL_PROPER_SALUTATION - 1];
      const email = row[FINAL_NOTICE.COL_EMAIL - 1];
      const driveLink = row[FINAL_NOTICE.COL_LINK - 1];
      const position = row[FINAL_NOTICE.COL_POSITION_EXTRACTED - 1];
      const office = row[FINAL_NOTICE.COL_ASSIGNED_OFFICE - 1];
      const statusCell = sheet.getRange(FINAL_NOTICE.START_ROW + i, FINAL_NOTICE.COL_STATUS);

      if (!applicantName || applicantName.toString().trim() === '') continue;

      if (!email || email.toString().trim() === '' || !driveLink || driveLink.toString().trim() === '') {
        statusCell.setValue('Not sent - missing email or link (' + now + ')');
        continue;
      }

      const subject = 'Job Application Update ' + '[' + position + ']';
      const body = 'Dear ' + salutation + ' ' + applicantLName + ',\n\n' +
        'Good day!\n\n' +
        'Thank you for your interest in the vacant position at our office and for participating in the interview.\n\n' +
        'Please see the file in the link below for more details regarding your application status:\n\n' +
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
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(WEB_APP_URL, options);

      if (response.getContentText() === "Success") {
        statusCell.setValue('Sent (' + now + ')');
        logSentLetter('LETTER - FINAL NOTICE', position || '', office || '', applicantName || '');
        emailCount++;
      } else {
        statusCell.setValue('Error: Proxy failed (' + now + ')');
      }

      if (shouldCancelFinalNoticeSend()) {
        console.log('Cancellation requested for email sending before sleep');
        cancelled = true;
        break;
      }

      Utilities.sleep(1500); // 1.5 second delay
    }

    return { status: cancelled ? 'Cancelled' : 'Emails sent', count: emailCount, cancelled: cancelled };
  } catch (e) {
    throw new Error('Error sending email notifications: ' + e.message);
  }
}

function finalNoticeSendIndividualEmails() {
  const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxJpyg6KPFUMxeHSOdOVnVe4WyN6JssT9DhoufEn2pE7vIp02joOQ6jZVD-FwZCLKW7FQ/exec";

  try {
    PropertiesService.getDocumentProperties().deleteProperty('cancel_finalNotice_send');
    CacheService.getDocumentCache().remove('cancel_finalNotice_send');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FINAL_NOTICE.SHEET_NAME);

    if (!sheet) throw new Error('Sheet "' + FINAL_NOTICE.SHEET_NAME + '" not found.');

    const lastRow = sheet.getLastRow();
    if (lastRow < FINAL_NOTICE.START_ROW) {
      return { status: 'No applicants found', count: 0 };
    }

    const data = sheet.getRange(FINAL_NOTICE.START_ROW, 1, lastRow - FINAL_NOTICE.START_ROW + 1, FINAL_NOTICE.COL_REGENERATE).getValues();

    const hasSelected = data.some(row => {
      const val = row[FINAL_NOTICE.COL_REGENERATE - 1];
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
      const regenerateVal = row[FINAL_NOTICE.COL_REGENERATE - 1];
      const shouldProcess = regenerateVal === true || String(regenerateVal).toLowerCase() === 'true';
      if (!shouldProcess) continue;

      if (shouldCancelFinalNoticeSend()) {
        console.log('Cancellation requested for individual email sending');
        cancelled = true;
        break;
      }

      const applicantName = row[0];
      const applicantLName = row[FINAL_NOTICE.COL_PROPER_LN - 1];
      const salutation = row[FINAL_NOTICE.COL_PROPER_SALUTATION - 1];
      const email = row[FINAL_NOTICE.COL_EMAIL - 1];
      const driveLink = row[FINAL_NOTICE.COL_LINK - 1];
      const position = row[FINAL_NOTICE.COL_POSITION_EXTRACTED - 1];
      const office = row[FINAL_NOTICE.COL_ASSIGNED_OFFICE - 1];
      const statusCell = sheet.getRange(FINAL_NOTICE.START_ROW + i, FINAL_NOTICE.COL_STATUS);

      if (!applicantName || applicantName.toString().trim() === '') {
        statusCell.setValue('Not sent - missing name (' + now + ')');
        sheet.getRange(FINAL_NOTICE.START_ROW + i, FINAL_NOTICE.COL_REGENERATE).setValue(false);
        continue;
      }

      if (!email || email.toString().trim() === '' || !driveLink || driveLink.toString().trim() === '') {
        statusCell.setValue('Not sent - missing email or link (' + now + ')');
        sheet.getRange(FINAL_NOTICE.START_ROW + i, FINAL_NOTICE.COL_REGENERATE).setValue(false);
        continue;
      }

      const subject = 'Job Application Update ' + '[' + position + ']';
      const body = 'Dear ' + salutation + ' ' + applicantLName + ',\n\n' +
        'Good day!\n\n' +
        'Thank you for your interest in the vacant position at our office and for participating in the interview.\n\n' +
        'Please see the file in the link below for more details regarding your application status:\n\n' +
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
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(WEB_APP_URL, options);
      if (response.getContentText() === "Success") {
        statusCell.setValue('Sent (re-sent) (' + now + ')');
        logSentLetter('LETTER - FINAL NOTICE', position || '', office || '', applicantName || '');
        emailCount++;
        sheet.getRange(FINAL_NOTICE.START_ROW + i, FINAL_NOTICE.COL_REGENERATE).setValue(false);
      } else {
        statusCell.setValue('Error: Proxy failed (' + now + ')');
      }

      if (shouldCancelFinalNoticeSend()) {
        console.log('Cancellation requested for individual email sending before sleep');
        cancelled = true;
        break;
      }

      Utilities.sleep(1500); // 1.5 second delay
    }

    return { status: cancelled ? 'Cancelled' : 'Individual emails processed', count: emailCount, cancelled: cancelled };
  } catch (e) {
    throw new Error('Error sending individual emails: ' + e.message);
  }
}

/**
 * Master function for Final Notice workflow
 */
function finalNoticeRunCompleteProcess() {
  try {
    const folderIds = finalNoticeCreateFolders();
    const pdfResult = finalNoticeGeneratePDFs(folderIds.finalNoticeSubFolderId);

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

function shouldCancelFinalNoticeRun() {
  const cache = CacheService.getDocumentCache();
  if (cache.get('cancel_finalNotice_run') === 'true') return true;
  return PropertiesService.getDocumentProperties().getProperty('cancel_finalNotice_run') === 'true';
}

function shouldCancelFinalNoticeSend() {
  const cache = CacheService.getDocumentCache();
  if (cache.get('cancel_finalNotice_send') === 'true') return true;
  return PropertiesService.getDocumentProperties().getProperty('cancel_finalNotice_send') === 'true';
}

function cancelFinalNoticeRun() {
  CacheService.getDocumentCache().put('cancel_finalNotice_run', 'true', 21600);
  PropertiesService.getDocumentProperties().setProperty('cancel_finalNotice_run', 'true');
}

function cancelFinalNoticeSend() {
  CacheService.getDocumentCache().put('cancel_finalNotice_send', 'true', 21600);
  PropertiesService.getDocumentProperties().setProperty('cancel_finalNotice_send', 'true');
}

function clearCancelFinalNoticeFlags() {
  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty('cancel_finalNotice_run');
  props.deleteProperty('cancel_finalNotice_send');

  const cache = CacheService.getDocumentCache();
  cache.remove('cancel_finalNotice_run');
  cache.remove('cancel_finalNotice_send');
}
