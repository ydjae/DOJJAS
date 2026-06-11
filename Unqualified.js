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
    props.setProperty('unqualifiedMainFolderId', mainFolderId);
  }

  let unqualifiedFolderId = props.getProperty('unqualifiedSubFolderId');
  let unqualifiedFolder = null;
  
  if (unqualifiedFolderId) {
    try {
      const tempSub = DriveApp.getFolderById(unqualifiedFolderId);
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
function unqualifiedCheckColumnR() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNQUALIFIED.SHEET_NAME);
    if (!sheet) {
      return { hasData: false, message: 'Sheet not found' };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < UNQUALIFIED.START_ROW) {
      return { hasData: false };
    }

    const dataRange = sheet.getRange(
      UNQUALIFIED.START_ROW, 
      1, 
      lastRow - UNQUALIFIED.START_ROW + 1, 
      18 // Check up to Column R
    ).getValues();

    for (let i = 0; i < dataRange.length; i++) {
      const rowData = dataRange[i];
      const valA = rowData[0]; // Column A

      if (valA && valA.toString().trim() !== '') {
        const colRValue = rowData[17]; // Column R (0-indexed)
        if (!colRValue || colRValue.toString().trim() === '') {
          const rowNum = UNQUALIFIED.START_ROW + i;
          return { 
            hasData: false, 
            message: 'Missing data in column R for applicant at row ' + rowNum 
          };
        }
      }
    }

    return { hasData: true };
  } catch (e) {
    return { hasData: false, message: e.message };
  }
}

/**
 * Generate PDFs from sheet and return list of processed applicants
 */
function unqualifiedGeneratePDFs(targetFolderId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(UNQUALIFIED.SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + UNQUALIFIED.SHEET_NAME + '" not found!');

    const data = sheet.getDataRange().getDisplayValues();
    const header = data[0];
    const rows = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowIndex = UNQUALIFIED.START_ROW + i - 1;
      if (row[0] && row[0].toString().trim() !== "") {
        rows.push({ row: row, rowIndex: rowIndex });
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

    // Check target column for existing links
    const lastRowWithValue = getLastRowWithValueInColumn(sheet, UNQUALIFIED.COL_LINK);
    let alreadyProcessedCount = 0;
    if (lastRowWithValue > 1) {
      alreadyProcessedCount = rows.filter(r => r.rowIndex <= lastRowWithValue).length;
    }

    const unprocessedRows = rows.filter(r => r.rowIndex > lastRowWithValue);

    if (unprocessedRows.length === 0) {
      setGenerationProgress('unqualified', rows.length, rows.length, 'completed');
      return {
        success: true,
        count: 0,
        applicants: [],
        completed: true,
        message: 'All eligible rows already have generated PDFs.'
      };
    }

    unprocessedRows.sort((a, b) => {
      const lastNameA = String(a.row[0] || "").trim().toLowerCase();
      const lastNameB = String(b.row[0] || "").trim().toLowerCase();
      if (lastNameA !== lastNameB) return lastNameA.localeCompare(lastNameB);
      const firstNameA = String(a.row[1] || "").trim().toLowerCase();
      const firstNameB = String(b.row[1] || "").trim().toLowerCase();
      return firstNameA.localeCompare(firstNameB);
    });

    const templateFile = DriveApp.getFileById(UNQUALIFIED.TEMPLATE_ID);
    const destinationFolder = DriveApp.getFolderById(targetFolderId);
    
    // Use batch processing with key for Unqualified
    const batchKey = 'unqualified_pdf_generation_' + SpreadsheetApp.getActiveSpreadsheet().getId();
    const batchResult = processUnqualifiedPDFBatch(batchKey, unprocessedRows, header, templateFile, destinationFolder, 20, sheet, rows.length, alreadyProcessedCount);

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
    throw new Error('Error generating unqualified PDFs: ' + e.message);
  }
}

function processUnqualifiedPDFBatch(batchKey, rows, header, templateFile, destinationFolder, batchSize, sheet, totalRows, alreadyProcessedCount) {
  let state = getBatchState(batchKey);

  if (!state) {
    PropertiesService.getDocumentProperties().deleteProperty('cancel_unqualified_run');
    CacheService.getDocumentCache().remove('cancel_unqualified_run');
    state = initializeBatchProcessing(batchKey, totalRows, alreadyProcessedCount);
  }

  const startIndex = state.currentIndex;
  setGenerationProgress('unqualified', state.completedCount, state.totalRows, 'processing');
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

      if (shouldCancelUnqualifiedRun()) {
        console.log('Cancellation requested for unqualified PDF generation');
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

        if (shouldCancelUnqualifiedRun()) {
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

        if (shouldCancelUnqualifiedRun()) {
          copy.setTrashed(true);
          state.status = 'cancelled';
          break;
        }

        const pdfBlob = copy.getAs(MimeType.PDF);
        const pdfFile = destinationFolder.createFile(pdfBlob).setName(fileName + ".pdf");
        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        copy.setTrashed(true);

        const pdfUrl = pdfFile.getUrl();
        sheet.getRange(rowIndex, UNQUALIFIED.COL_LINK).setValue(pdfUrl);
        SpreadsheetApp.flush();

        state.currentIndex = i + 1;
        state.completedCount = state.alreadyProcessedCount + state.currentIndex;
        state.processedApplicants.push(lastName + ', ' + firstName);
        newApplicants.push(lastName + ', ' + firstName);
        processedInThisBatch++;
      } catch (itemError) {
        console.log('Error processing ' + fileName + ': ' + itemError.message);
        state.currentIndex = i + 1;
        state.completedCount = state.alreadyProcessedCount + state.currentIndex;
      }
      setGenerationProgress('unqualified', state.completedCount, state.totalRows, 'processing');
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
  setGenerationProgress('unqualified', state.completedCount, state.totalRows, state.status === 'cancelled' ? 'cancelled' : (state.currentIndex >= rows.length ? 'completed' : 'processing'));

  return {
    completed: isCompleted,
    processed: processedInThisBatch,
    totalProcessed: state.completedCount,
    totalRows: state.totalRows,
    applicants: newApplicants,
    allApplicants: state.processedApplicants,
    message: state.status === 'cancelled'
      ? 'Process cancelled. Total: ' + state.completedCount + ' / ' + state.totalRows
      : processedInThisBatch + ' applicants processed. Total: ' + state.completedCount + ' / ' + state.totalRows,
    status: state.status
  };
}

function unqualifiedGenerateIndividualPDFs() {
  try {
    PropertiesService.getDocumentProperties().deleteProperty('cancel_unqualified_run');
    CacheService.getDocumentCache().remove('cancel_unqualified_run');
    
    const folderIds = getUnqualifiedPositionFolder();
    const targetFolderId = folderIds.unqualifiedSubFolderId;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(UNQUALIFIED.SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + UNQUALIFIED.SHEET_NAME + '" not found.');

    const data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) {
      throw new Error('No items checked in REGENERATE column (S). Please check at least one checkbox to proceed.');
    }

    const header = data[0];
    const selectedRows = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const regenerateVal = row[UNQUALIFIED.COL_REGENERATE - 1];
      const shouldGenerate = regenerateVal === true || String(regenerateVal).toLowerCase() === 'true';
      if (!shouldGenerate) continue;
      if (!row[0] || row[0].toString().trim() === '') continue;
      selectedRows.push({ row, rowIndex: UNQUALIFIED.START_ROW + i - 1 });
    }

    if (selectedRows.length === 0) {
      throw new Error('No items checked in REGENERATE column (S). Please check at least one checkbox to proceed.');
    }
    setGenerationProgress('unqualified', 0, selectedRows.length, 'processing');

    const templateFile = DriveApp.getFileById(UNQUALIFIED.TEMPLATE_ID);
    const destinationFolder = DriveApp.getFolderById(targetFolderId);
    const processed = [];
    let cancelled = false;

    for (let k = 0; k < selectedRows.length; k++) {
      if (shouldCancelUnqualifiedRun()) {
        console.log('Cancellation requested for individual unqualified PDF generation');
        cancelled = true;
        break;
      }

      const { row, rowIndex } = selectedRows[k];
      const lastName = String(row[UNQUALIFIED.COL_LAST_NAME - 1] || '').trim();
      const firstName = String(row[UNQUALIFIED.COL_FIRST_NAME - 1] || '').trim();
      const fileName = (lastName || 'Applicant') + (firstName ? (', ' + firstName) : '');

      try {
        const copy = templateFile.makeCopy(fileName, destinationFolder);

        if (shouldCancelUnqualifiedRun()) {
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

        if (shouldCancelUnqualifiedRun()) {
          copy.setTrashed(true);
          cancelled = true;
          break;
        }

        const pdfBlob = copy.getAs(MimeType.PDF);
        const pdfFile = destinationFolder.createFile(pdfBlob).setName(fileName + '.pdf');
        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        copy.setTrashed(true);

        const pdfUrl = pdfFile.getUrl();
        sheet.getRange(rowIndex, UNQUALIFIED.COL_LINK).setValue(pdfUrl);
        SpreadsheetApp.flush();
        sheet.getRange(rowIndex, UNQUALIFIED.COL_REGENERATE).setValue(false);
        processed.push(fileName);
      } catch (itemError) {
        console.log('Error generating unqualified PDF for row ' + rowIndex + ': ' + itemError.message);
      }
      setGenerationProgress('unqualified', k + 1, selectedRows.length, 'processing');
    }
    setGenerationProgress('unqualified', processed.length, selectedRows.length, cancelled ? 'cancelled' : 'completed');

    return { success: true, count: processed.length, applicants: processed, cancelled: cancelled };
  } catch (e) {
    throw new Error('Error generating selected unqualified PDFs: ' + e.message);
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
  const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxJpyg6KPFUMxeHSOdOVnVe4WyN6JssT9DhoufEn2pE7vIp02joOQ6jZVD-FwZCLKW7FQ/exec";

  try {
    PropertiesService.getDocumentProperties().deleteProperty('cancel_unqualified_send');
    CacheService.getDocumentCache().remove('cancel_unqualified_send');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(UNQUALIFIED.SHEET_NAME);

    if (!sheet) throw new Error('Sheet "' + UNQUALIFIED.SHEET_NAME + '" not found.');

    const lastRow = sheet.getLastRow();
    if (lastRow < UNQUALIFIED.START_ROW) {
      return { status: 'No applicants found', count: 0 };
    }

    const data = sheet.getRange(UNQUALIFIED.START_ROW, 1, lastRow - UNQUALIFIED.START_ROW + 1, UNQUALIFIED.COL_STATUS).getValues();
    let validRowCount = 0;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() !== '') validRowCount++;
      else break;
    }
    let emailCount = 0;
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    let cancelled = false;

    setGenerationProgress('unqualifiedEmail', 0, validRowCount, 'processing');

    for (let i = 0; i < data.length; i++) {
      setGenerationProgress('unqualifiedEmail', i, validRowCount, 'processing');

      if (shouldCancelUnqualifiedSend()) {
        console.log('Cancellation requested for unqualified email sending');
        cancelled = true;
        break;
      }

      const row = data[i];
      const applicantName = row[0];
      const applicantLName = row[13]; // Column N
      const salutation = row[12]; // Column M
      const email = row[UNQUALIFIED.COL_EMAIL - 1];
      const driveLink = row[UNQUALIFIED.COL_LINK - 1];
      const position = row[7]; // Column H
      const office = row[8]; // Column I
      const statusCell = sheet.getRange(UNQUALIFIED.START_ROW + i, UNQUALIFIED.COL_STATUS);

      if (!applicantName || applicantName.toString().trim() === '') break;

      const currentStatus = row[UNQUALIFIED.COL_STATUS - 1];
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
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(WEB_APP_URL, options);

      if (response.getContentText() === "Success") {
        statusCell.setValue('Sent (' + now + ')');
        logSentLetter('LETTER - DQ', position || '', office || '', applicantName || '');
        emailCount++;
      } else {
        statusCell.setValue('Error: Proxy failed (' + now + ')');
      }

      if (shouldCancelUnqualifiedSend()) {
        console.log('Cancellation requested for unqualified email sending before sleep');
        cancelled = true;
        break;
      }

      Utilities.sleep(1500);
    }

    setGenerationProgress('unqualifiedEmail', validRowCount, validRowCount, cancelled ? 'cancelled' : 'completed');

    return { status: cancelled ? 'Cancelled' : 'Emails sent', count: emailCount, cancelled: cancelled };
  } catch (e) {
    throw new Error('Error sending email notifications: ' + e.message);
  }
}

function unqualifiedSendSelectedEmails() {
  const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxJpyg6KPFUMxeHSOdOVnVe4WyN6JssT9DhoufEn2pE7vIp02joOQ6jZVD-FwZCLKW7FQ/exec";

  try {
    PropertiesService.getDocumentProperties().deleteProperty('cancel_unqualified_send');
    CacheService.getDocumentCache().remove('cancel_unqualified_send');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(UNQUALIFIED.SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + UNQUALIFIED.SHEET_NAME + '" not found.');

    const lastRow = sheet.getLastRow();
    if (lastRow < UNQUALIFIED.START_ROW) {
      return { status: 'No applicants found', count: 0 };
    }

    const data = sheet.getRange(UNQUALIFIED.START_ROW, 1, lastRow - UNQUALIFIED.START_ROW + 1, UNQUALIFIED.COL_REGENERATE).getValues();

    let validRowCount = 0;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() !== '') validRowCount++;
      else break;
    }

    const hasSelected = data.some(row => {
      const val = row[UNQUALIFIED.COL_REGENERATE - 1];
      return val === true || String(val).toLowerCase() === 'true';
    });
    if (!hasSelected) {
      throw new Error('No items checked in REGENERATE column. Please check at least one checkbox to proceed.');
    }

    let emailCount = 0;
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    let cancelled = false;

    setGenerationProgress('unqualifiedEmail', 0, validRowCount, 'processing');

    for (let i = 0; i < data.length; i++) {
      setGenerationProgress('unqualifiedEmail', i, validRowCount, 'processing');

      const row = data[i];
      const regenerateVal = row[UNQUALIFIED.COL_REGENERATE - 1];
      const shouldSend = regenerateVal === true || String(regenerateVal).toLowerCase() === 'true';
      if (!shouldSend) continue;

      if (shouldCancelUnqualifiedSend()) {
        console.log('Cancellation requested for selected unqualified email sending');
        cancelled = true;
        break;
      }

      const applicantName = row[0];
      const applicantLName = row[13]; // Column N
      const salutation = row[12]; // Column M
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
        logSentLetter('LETTER - DQ', position || '', office || '', applicantName || '');
        emailCount++;
        sheet.getRange(UNQUALIFIED.START_ROW + i, UNQUALIFIED.COL_REGENERATE).setValue(false);
      } else {
        statusCell.setValue('Error: Proxy failed (' + now + ')');
      }

      if (shouldCancelUnqualifiedSend()) {
        console.log('Cancellation requested for selected unqualified email sending before sleep');
        cancelled = true;
        break;
      }

      Utilities.sleep(1500);
    }

    setGenerationProgress('unqualifiedEmail', validRowCount, validRowCount, cancelled ? 'cancelled' : 'completed');

    return { status: cancelled ? 'Cancelled' : 'Selected emails processed', count: emailCount, cancelled: cancelled };
  } catch (e) {
    throw new Error('Error sending selected emails: ' + e.message);
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

// ==========================================
// CANCELLATION HELPERS FOR UNQUALIFIED
// ==========================================

function shouldCancelUnqualifiedRun() {
  const cache = CacheService.getDocumentCache();
  if (cache.get('cancel_unqualified_run') === 'true') return true;
  return PropertiesService.getDocumentProperties().getProperty('cancel_unqualified_run') === 'true';
}

function shouldCancelUnqualifiedSend() {
  const cache = CacheService.getDocumentCache();
  if (cache.get('cancel_unqualified_send') === 'true') return true;
  return PropertiesService.getDocumentProperties().getProperty('cancel_unqualified_send') === 'true';
}

function cancelUnqualifiedRun() {
  CacheService.getDocumentCache().put('cancel_unqualified_run', 'true', 21600);
  PropertiesService.getDocumentProperties().setProperty('cancel_unqualified_run', 'true');
}

function cancelUnqualifiedSend() {
  CacheService.getDocumentCache().put('cancel_unqualified_send', 'true', 21600);
  PropertiesService.getDocumentProperties().setProperty('cancel_unqualified_send', 'true');
}

function clearCancelUnqualifiedFlags() {
  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty('cancel_unqualified_run');
  props.deleteProperty('cancel_unqualified_send');
  const cache = CacheService.getDocumentCache();
  cache.remove('cancel_unqualified_run');
  cache.remove('cancel_unqualified_send');
}
