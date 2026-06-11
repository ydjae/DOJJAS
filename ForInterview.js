// ==========================================
// FOR INTERVIEW - Letter Generation & Email Workflow
// ==========================================

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
function forInterviewCheckColumnsOtoR() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FOR_INTERVIEW.SHEET_NAME);
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

    let mainFolderId = props.getProperty('forInterviewMainFolderId');
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
      props.setProperty('forInterviewMainFolderId', mainFolderId);
    }

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
    const rows = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowIndex = FOR_INTERVIEW.START_ROW + i - 1;
      if (row[FOR_INTERVIEW.COL_TEMPLATE_COLS_START - 1] && row[FOR_INTERVIEW.COL_TEMPLATE_COLS_START - 1].toString().trim() !== "") {
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
    const lastRowWithValue = getLastRowWithValueInColumn(sheet, FOR_INTERVIEW.COL_INTERVIEW_LINK);
    let alreadyProcessedCount = 0;
    if (lastRowWithValue > 1) {
      alreadyProcessedCount = rows.filter(r => r.rowIndex <= lastRowWithValue).length;
    }

    const unprocessedRows = rows.filter(r => r.rowIndex > lastRowWithValue);

    if (unprocessedRows.length === 0) {
      setGenerationProgress('interview', rows.length, rows.length, 'completed');
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

    const templateFile = DriveApp.getFileById(FOR_INTERVIEW.TEMPLATE_ID);
    const destinationFolder = DriveApp.getFolderById(targetFolderId);
    
    // Use batch processing with key for For Interview
    const batchKey = 'forInterview_pdf_generation_' + SpreadsheetApp.getActiveSpreadsheet().getId();
    const batchResult = processInterviewPDFBatch(batchKey, unprocessedRows, header, templateFile, destinationFolder, 20, sheet, rows.length, alreadyProcessedCount);

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
    throw new Error('Error generating PDFs: ' + e.message);
  }
}

function processInterviewPDFBatch(batchKey, rows, header, templateFile, destinationFolder, batchSize, sheet, totalRows, alreadyProcessedCount) {
  let state = getBatchState(batchKey);

  if (!state) {
    PropertiesService.getDocumentProperties().deleteProperty('cancel_forInterview_run');
    CacheService.getDocumentCache().remove('cancel_forInterview_run');
    state = initializeBatchProcessing(batchKey, totalRows, alreadyProcessedCount);
  }

  const startIndex = state.currentIndex;
  setGenerationProgress('interview', state.completedCount, state.totalRows, 'processing');
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

      if (shouldCancelInterviewRun()) {
        console.log('Cancellation requested for interview PDF generation');
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

        if (shouldCancelInterviewRun()) {
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

        if (shouldCancelInterviewRun()) {
          copy.setTrashed(true);
          state.status = 'cancelled';
          break;
        }

        const pdfBlob = copy.getAs(MimeType.PDF);
        const pdfFile = destinationFolder.createFile(pdfBlob).setName(fileName + ".pdf");
        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        copy.setTrashed(true);

        const pdfUrl = pdfFile.getUrl();
        sheet.getRange(rowIndex, FOR_INTERVIEW.COL_INTERVIEW_LINK).setValue(pdfUrl);
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
      setGenerationProgress('interview', state.completedCount, state.totalRows, 'processing');
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
  setGenerationProgress('interview', state.completedCount, state.totalRows, state.status === 'cancelled' ? 'cancelled' : (state.currentIndex >= rows.length ? 'completed' : 'processing'));

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

function forInterviewGenerateIndividualPDFs() {
  try {
    PropertiesService.getDocumentProperties().deleteProperty('cancel_forInterview_run');
    CacheService.getDocumentCache().remove('cancel_forInterview_run');
    
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
    setGenerationProgress('interview', 0, rowsToProcess.length, 'processing');

    const templateFile = DriveApp.getFileById(FOR_INTERVIEW.TEMPLATE_ID);
    const processed = [];
    let cancelled = false;

    for (let k = 0; k < rowsToProcess.length; k++) {
      if (shouldCancelInterviewRun()) {
        console.log('Cancellation requested for individual interview PDF generation');
        cancelled = true;
        break;
      }

      const rowObj = rowsToProcess[k];
      const row = rowObj.row;
      const rowIndex = rowObj.rowIndex;
      const lastName = String(row[FOR_INTERVIEW.COL_LAST_NAME - 1] || '').trim();
      const firstName = String(row[FOR_INTERVIEW.COL_FIRST_NAME - 1] || '').trim();
      const fileName = (lastName || 'Applicant') + (firstName ? (', ' + firstName) : '');

      try {
        const copy = templateFile.makeCopy(fileName, destinationFolder);

        if (shouldCancelInterviewRun()) {
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

        if (shouldCancelInterviewRun()) {
          copy.setTrashed(true);
          cancelled = true;
          break;
        }

        const pdfBlob = copy.getAs(MimeType.PDF);
        const pdfFile = destinationFolder.createFile(pdfBlob).setName(fileName + '.pdf');
        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        copy.setTrashed(true);

        const pdfUrl = pdfFile.getUrl();
        sheet.getRange(rowIndex, FOR_INTERVIEW.COL_INTERVIEW_LINK).setValue(pdfUrl);
        SpreadsheetApp.flush();
        sheet.getRange(rowIndex, FOR_INTERVIEW.COL_REGENERATE).setValue(false);
        processed.push(fileName);
      } catch (itemError) {
        console.log('Error generating interview PDF for row ' + rowIndex + ': ' + itemError.message);
      }
      setGenerationProgress('interview', k + 1, rowsToProcess.length, 'processing');
    }
    setGenerationProgress('interview', processed.length, rowsToProcess.length, cancelled ? 'cancelled' : 'completed');

    return { success: true, count: processed.length, applicants: processed, cancelled: cancelled };
  } catch (e) {
    throw new Error('Error generating individual interview PDFs: ' + e.message);
  }
}

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

function forInterviewSendEmails() {
  const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxJpyg6KPFUMxeHSOdOVnVe4WyN6JssT9DhoufEn2pE7vIp02joOQ6jZVD-FwZCLKW7FQ/exec";

  try {
    PropertiesService.getDocumentProperties().deleteProperty('cancel_forInterview_send');
    CacheService.getDocumentCache().remove('cancel_forInterview_send');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FOR_INTERVIEW.SHEET_NAME);

    if (!sheet) throw new Error('Sheet "' + FOR_INTERVIEW.SHEET_NAME + '" not found.');

    const lastRow = sheet.getLastRow();
    if (lastRow < FOR_INTERVIEW.START_ROW) {
      return { status: 'No applicants found', count: 0 };
    }

    const data = sheet.getRange(FOR_INTERVIEW.START_ROW, 1, lastRow - FOR_INTERVIEW.START_ROW + 1, FOR_INTERVIEW.COL_INTERVIEW_PROGRESS).getValues();
    let validRowCount = 0;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() !== '') validRowCount++;
      else break;
    }
    let emailCount = 0;
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    let cancelled = false;

    setGenerationProgress('interviewEmail', 0, validRowCount, 'processing');

    for (let i = 0; i < data.length; i++) {
      setGenerationProgress('interviewEmail', i, validRowCount, 'processing');

      if (shouldCancelInterviewSend()) {
        console.log('Cancellation requested for interview email sending');
        cancelled = true;
        break;
      }

      const row = data[i];
      const applicantName = row[0];
      const applicantLName = row[12]; // Column M
      const salutation = row[11]; // Column L
      const email = row[FOR_INTERVIEW.COL_EMAIL - 1];
      const driveLink = row[FOR_INTERVIEW.COL_INTERVIEW_LINK - 1];
      const position = row[7]; // Column H
      const office = row[8]; // Column I
      const statusCell = sheet.getRange(FOR_INTERVIEW.START_ROW + i, FOR_INTERVIEW.COL_INTERVIEW_PROGRESS);

      if (!applicantName || applicantName.toString().trim() === '') break;

      const currentStatus = row[FOR_INTERVIEW.COL_INTERVIEW_PROGRESS - 1];
      if (currentStatus && String(currentStatus).startsWith('Sent')) continue;

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
        logSentLetter('LETTER - FOR INTERVIEW', position || '', office || '', applicantName || '');
        emailCount++;
      } else {
        statusCell.setValue('Error: Proxy failed (' + now + ')');
      }

      if (shouldCancelInterviewSend()) {
        console.log('Cancellation requested for interview email sending before sleep');
        cancelled = true;
        break;
      }

      Utilities.sleep(1500);
    }

    setGenerationProgress('interviewEmail', validRowCount, validRowCount, cancelled ? 'cancelled' : 'completed');

    return { status: cancelled ? 'Cancelled' : 'Emails sent', count: emailCount, cancelled: cancelled };
  } catch (e) {
    throw new Error('Error sending email notifications: ' + e.message);
  }
}

function forInterviewSendSelectedEmails() {
  const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxJpyg6KPFUMxeHSOdOVnVe4WyN6JssT9DhoufEn2pE7vIp02joOQ6jZVD-FwZCLKW7FQ/exec";

  try {
    PropertiesService.getDocumentProperties().deleteProperty('cancel_forInterview_send');
    CacheService.getDocumentCache().remove('cancel_forInterview_send');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FOR_INTERVIEW.SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + FOR_INTERVIEW.SHEET_NAME + '" not found.');

    const lastRow = sheet.getLastRow();
    if (lastRow < FOR_INTERVIEW.START_ROW) {
      return { status: 'No applicants found', count: 0 };
    }

    const data = sheet.getRange(FOR_INTERVIEW.START_ROW, 1, lastRow - FOR_INTERVIEW.START_ROW + 1, FOR_INTERVIEW.COL_REGENERATE).getValues();

    let validRowCount = 0;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() !== '') validRowCount++;
      else break;
    }

    const hasSelected = data.some(row => {
      const val = row[FOR_INTERVIEW.COL_REGENERATE - 1];
      return val === true || String(val).toLowerCase() === 'true';
    });
    if (!hasSelected) {
      throw new Error('No items checked in REGENERATE column. Please check at least one checkbox to proceed.');
    }

    let emailCount = 0;
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    let cancelled = false;

    setGenerationProgress('interviewEmail', 0, validRowCount, 'processing');

    for (let i = 0; i < data.length; i++) {
      setGenerationProgress('interviewEmail', i, validRowCount, 'processing');

      const row = data[i];
      const regenerateVal = row[FOR_INTERVIEW.COL_REGENERATE - 1];
      const shouldSend = regenerateVal === true || String(regenerateVal).toLowerCase() === 'true';
      if (!shouldSend) continue;

      if (shouldCancelInterviewSend()) {
        console.log('Cancellation requested for selected interview email sending');
        cancelled = true;
        break;
      }

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

      if (shouldCancelInterviewSend()) {
        console.log('Cancellation requested for selected interview email sending before sleep');
        cancelled = true;
        break;
      }

      Utilities.sleep(1500);
    }

    setGenerationProgress('interviewEmail', validRowCount, validRowCount, cancelled ? 'cancelled' : 'completed');

    return { status: cancelled ? 'Cancelled' : 'Selected emails processed', count: emailCount, cancelled: cancelled };
  } catch (e) {
    throw new Error('Error sending selected interview emails: ' + e.message);
  }
}

function forInterviewRunCompleteProcess() {
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

// ==========================================
// CANCELLATION HELPERS FOR INTERVIEW
// ==========================================

function shouldCancelInterviewRun() {
  const cache = CacheService.getDocumentCache();
  if (cache.get('cancel_forInterview_run') === 'true') return true;
  return PropertiesService.getDocumentProperties().getProperty('cancel_forInterview_run') === 'true';
}

function shouldCancelInterviewSend() {
  const cache = CacheService.getDocumentCache();
  if (cache.get('cancel_forInterview_send') === 'true') return true;
  return PropertiesService.getDocumentProperties().getProperty('cancel_forInterview_send') === 'true';
}

function cancelInterviewRun() {
  CacheService.getDocumentCache().put('cancel_forInterview_run', 'true', 21600);
  PropertiesService.getDocumentProperties().setProperty('cancel_forInterview_run', 'true');
}

function cancelInterviewSend() {
  CacheService.getDocumentCache().put('cancel_forInterview_send', 'true', 21600);
  PropertiesService.getDocumentProperties().setProperty('cancel_forInterview_send', 'true');
}

function clearCancelInterviewFlags() {
  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty('cancel_forInterview_run');
  props.deleteProperty('cancel_forInterview_send');
  const cache = CacheService.getDocumentCache();
  cache.remove('cancel_forInterview_run');
  cache.remove('cancel_forInterview_send');
}