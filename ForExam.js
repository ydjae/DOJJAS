// ==========================================
// FOR EXAM - Letter Generation & Email Workflow
// ==========================================

// Constants for For Exam workflow
const FOR_EXAM = {
  SHEET_NAME: 'LETTER - EXAM SCHED',
  TEMPLATE_ID: '1V0icY2qWob_D2LQpe24H5Q85p5QXnfHgwlLt8sWeSNs',
  COL_LAST_NAME: 1,
  COL_FIRST_NAME: 2,
  COL_EMAIL: 7,
  COL_TEMPLATE_COLS_START: 15, // Column O
  COL_TEMPLATE_COLS_END: 18,   // Column R
  COL_EXAM_LINK: 19,           // Column S
  COL_EXAM_PROGRESS: 20,       // Column T
  COL_REGENERATE: 21,         // Column U - REGENERATE and RESEND checkbox
  START_ROW: 2
};

/**
 * Check if columns O-R have data in the EXAM tab
 */
function forExamCheckColumnsOtoR() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FOR_EXAM.SHEET_NAME);
    if (!sheet) {
      return { hasData: false, message: 'Sheet not found' };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < FOR_EXAM.START_ROW) {
      return { hasData: false };
    }

    const dataRange = sheet.getRange(
      FOR_EXAM.START_ROW, 
      1, 
      lastRow - FOR_EXAM.START_ROW + 1, 
      FOR_EXAM.COL_TEMPLATE_COLS_END
    ).getValues();

    for (let i = 0; i < dataRange.length; i++) {
      const rowData = dataRange[i];
      const valA = rowData[0]; // Column A

      if (valA && valA.toString().trim() !== '') {
        for (let colIdx = FOR_EXAM.COL_TEMPLATE_COLS_START - 1; colIdx <= FOR_EXAM.COL_TEMPLATE_COLS_END - 1; colIdx++) {
          const cellValue = rowData[colIdx];
          if (!cellValue || cellValue.toString().trim() === '') {
            const rowNum = FOR_EXAM.START_ROW + i;
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
 * Create main folder and For Exam subfolder
 */
function forExamCreateFolders() {
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

    let mainFolderId = props.getProperty('forExamMainFolderId');
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
      props.setProperty('forExamMainFolderId', mainFolderId);
    }

    let forExamSubFolderId = props.getProperty('forExamSubFolderId');
    let forExamSubFolder = null;

    if (forExamSubFolderId) {
      try {
        const tempSubFolder = DriveApp.getFolderById(forExamSubFolderId);
        if (tempSubFolder.getParents().hasNext() && tempSubFolder.getParents().next().getId() === mainFolderId) {
          forExamSubFolder = tempSubFolder;
        }
      } catch (e) {
        forExamSubFolder = null;
      }
    }

    if (!forExamSubFolder) {
      const existingSubFolders = mainFolder.getFoldersByName('For Exam');
      if (existingSubFolders.hasNext()) {
        forExamSubFolder = existingSubFolders.next();
      } else {
        forExamSubFolder = mainFolder.createFolder('For Exam');
      }
      forExamSubFolderId = forExamSubFolder.getId();
      props.setProperty('forExamSubFolderId', forExamSubFolderId);
    }

    return {
      mainFolderId: mainFolderId,
      forExamSubFolderId: forExamSubFolderId,
      folderUrl: forExamSubFolder.getUrl()
    };
  } catch (e) {
    throw new Error('Error creating folders: ' + e.message);
  }
}

/**
 * Generate PDFs from sheet and return list of processed applicants
 */
function forExamGeneratePDFs(targetFolderId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FOR_EXAM.SHEET_NAME);
    
    if (!sheet) throw new Error("Sheet '" + FOR_EXAM.SHEET_NAME + "' not found!");
    
    const data = sheet.getDataRange().getDisplayValues();
    const header = data[0];
    const rows = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowIndex = FOR_EXAM.START_ROW + i - 1;
      if (row[FOR_EXAM.COL_TEMPLATE_COLS_START - 1] && row[FOR_EXAM.COL_TEMPLATE_COLS_START - 1].toString().trim() !== "") {
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

    rows.sort((a, b) => {
      const lastNameA = String(a.row[0] || "").trim().toLowerCase();
      const lastNameB = String(b.row[0] || "").trim().toLowerCase();
      if (lastNameA !== lastNameB) return lastNameA.localeCompare(lastNameB);
      const firstNameA = String(a.row[1] || "").trim().toLowerCase();
      const firstNameB = String(b.row[1] || "").trim().toLowerCase();
      return firstNameA.localeCompare(firstNameB);
    });
    
    const templateFile = DriveApp.getFileById(FOR_EXAM.TEMPLATE_ID);
    const destinationFolder = DriveApp.getFolderById(targetFolderId);
    
    // Use batch processing with key for For Exam
    const batchKey = 'forExam_pdf_generation_' + SpreadsheetApp.getActiveSpreadsheet().getId();
    const batchResult = processForExamPDFBatch(batchKey, rows, header, templateFile, destinationFolder, 20, sheet);

    let returnMessage = batchResult.message;
    
    if (batchResult.status === 'cancelled') {
      returnMessage = 'Process cancelled. ' + batchResult.totalProcessed + ' PDFs generated.';
    } else if (!batchResult.completed) {
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
      completed: batchResult.completed || batchResult.status === 'cancelled',
      cancelled: batchResult.status === 'cancelled',
      message: returnMessage
    };
  } catch (e) {
    throw new Error('Error generating PDFs: ' + e.message);
  }
}

function processForExamPDFBatch(batchKey, rows, header, templateFile, destinationFolder, batchSize, sheet) {
  let state = getBatchState(batchKey);

  if (!state) {
    PropertiesService.getDocumentProperties().deleteProperty('cancel_forExam_run');
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

      if (PropertiesService.getDocumentProperties().getProperty('cancel_forExam_run') === 'true') {
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

        if (PropertiesService.getDocumentProperties().getProperty('cancel_forExam_run') === 'true') {
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

        if (PropertiesService.getDocumentProperties().getProperty('cancel_forExam_run') === 'true') {
          copy.setTrashed(true);
          state.status = 'cancelled';
          break;
        }

        const pdfBlob = copy.getAs(MimeType.PDF);
        const pdfFile = destinationFolder.createFile(pdfBlob).setName(fileName + ".pdf");
        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        copy.setTrashed(true);

        const pdfUrl = pdfFile.getUrl();
        sheet.getRange(rowIndex, FOR_EXAM.COL_EXAM_LINK).setValue(pdfUrl);
        SpreadsheetApp.flush();

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
 * Generate PDFs only for rows with the REGENERATE checkbox checked (Column U).
 * Writes the Drive PDF link to Column S immediately. Does NOT clear the checkbox.
 */
function forExamGenerateIndividualPDFs() {
  try {
    PropertiesService.getDocumentProperties().deleteProperty('cancel_forExam_run');
    const props = PropertiesService.getDocumentProperties();
    let folderId = props.getProperty('forExamSubFolderId');
    let destinationFolder = null;

    if (folderId) {
      try {
        destinationFolder = DriveApp.getFolderById(folderId);
      } catch (folderError) {
        console.log('Stored For Exam subfolder ID invalid. Recreating folder: ' + folderError.message);
        const created = forExamCreateFolders();
        folderId = created.forExamSubFolderId;
        destinationFolder = DriveApp.getFolderById(folderId);
      }
    }

    if (!destinationFolder) {
      const created = forExamCreateFolders();
      folderId = created.forExamSubFolderId;
      destinationFolder = DriveApp.getFolderById(folderId);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FOR_EXAM.SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + FOR_EXAM.SHEET_NAME + '" not found.');

    const data = sheet.getDataRange().getDisplayValues();
    const header = data[0] || [];
    const rowsToProcess = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowIndex = FOR_EXAM.START_ROW + i - 1;
      const regenerateVal = row[FOR_EXAM.COL_REGENERATE - 1];
      const shouldProcess = regenerateVal === true || String(regenerateVal).toLowerCase() === 'true';
      if (!shouldProcess) continue;
      if (!row[FOR_EXAM.COL_TEMPLATE_COLS_START - 1] || row[FOR_EXAM.COL_TEMPLATE_COLS_START - 1].toString().trim() === '') continue;
      rowsToProcess.push({ row: row, rowIndex: rowIndex });
    }

    if (rowsToProcess.length === 0) {
      throw new Error('No items checked in REGENERATE column (U). Please check at least one checkbox to proceed.');
    }

    let templateFile;
    try {
      templateFile = DriveApp.getFileById(FOR_EXAM.TEMPLATE_ID);
    } catch (idError) {
      throw new Error('Error loading template file: ' + idError.message + '. Please verify FOR_EXAM.TEMPLATE_ID is a valid, accessible Google Docs template.');
    }
    if (templateFile.getMimeType() !== MimeType.GOOGLE_DOCS) {
      throw new Error('Template file is not a Google Doc. Please use a Google Docs template for FOR EXAM letters.');
    }

    const processed = [];
    let cancelled = false;

    for (let k = 0; k < rowsToProcess.length; k++) {
      if (PropertiesService.getDocumentProperties().getProperty('cancel_forExam_run') === 'true') {
        console.log('Cancellation requested for individual PDF generation');
        cancelled = true;
        break;
      }

      const rowObj = rowsToProcess[k];
      const row = rowObj.row;
      const rowIndex = rowObj.rowIndex;
      const lastName = String(row[FOR_EXAM.COL_LAST_NAME - 1] || "").trim();
      const firstName = String(row[FOR_EXAM.COL_FIRST_NAME - 1] || "").trim();
      const fileName = (lastName || 'Applicant') + (firstName ? (', ' + firstName) : '');

      try {
        const copy = templateFile.makeCopy(fileName, destinationFolder);

        if (PropertiesService.getDocumentProperties().getProperty('cancel_forExam_run') === 'true') {
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

        if (PropertiesService.getDocumentProperties().getProperty('cancel_forExam_run') === 'true') {
          copy.setTrashed(true);
          cancelled = true;
          break;
        }

        const pdfBlob = copy.getAs(MimeType.PDF);
        const pdfFile = destinationFolder.createFile(pdfBlob).setName(fileName + ".pdf");
        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        copy.setTrashed(true);

        const pdfUrl = pdfFile.getUrl();
        sheet.getRange(rowIndex, FOR_EXAM.COL_EXAM_LINK).setValue(pdfUrl);
        SpreadsheetApp.flush();

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
function forExamGenerateLinks() {
  try {
    const settings = PropertiesService.getDocumentProperties();
    const folderId = settings.getProperty('forExamSubFolderId');
    
    if (!folderId) {
      throw new Error('For Exam subfolder not found. Please run Step 2 first.');
    }
    
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByType(MimeType.PDF);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FOR_EXAM.SHEET_NAME);
    
    if (!sheet) {
      throw new Error('Sheet "' + FOR_EXAM.SHEET_NAME + '" not found.');
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
      sheet.getRange(FOR_EXAM.START_ROW, FOR_EXAM.COL_EXAM_LINK, links.length, 1).setValues(links);
    }
    
    return 'Successfully generated and inserted ' + links.length + ' Google Drive links into Column S.';
  } catch (e) {
    throw new Error('Error generating Drive links: ' + e.message);
  }
}

/**
 * Get the For Exam folder URL
 */
function forExamGetFolderUrl() {
  try {
    const settings = PropertiesService.getDocumentProperties();
    const folderId = settings.getProperty('forExamSubFolderId');
    
    if (!folderId) {
      throw new Error('For Exam subfolder not found. Please run the process first.');
    }
    
    const folder = DriveApp.getFolderById(folderId);
    return folder.getUrl();
  } catch (e) {
    throw new Error('Error retrieving folder URL: ' + e.message);
  }
}

/**
 * Backup the Letter - Exam Sched sheet to the For Exam folder
 */
function forExamBackupSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(FOR_EXAM.SHEET_NAME);
    
    if (!sourceSheet) {
      throw new Error('Sheet "' + FOR_EXAM.SHEET_NAME + '" not found.');
    }
    
    const settings = PropertiesService.getDocumentProperties();
    const forExamSubFolderId = settings.getProperty('forExamSubFolderId');
    
    if (!forExamSubFolderId) {
      throw new Error('For Exam subfolder not found. Please run Step 2 first.');
    }
    
    const forExamFolder = DriveApp.getFolderById(forExamSubFolderId);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    const backupFileName = 'LETTER - EXAM SCHED_' + timestamp + '.csv';
    
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
    forExamFolder.createFile(backupBlob);
    
    return {
      message: 'Backup successful! LETTER - EXAM SCHED has been saved to the For Exam folder.',
      folderUrl: forExamFolder.getUrl()
    };
  } catch (e) {
    throw new Error('Error backing up sheet: ' + e.message);
  }
}

function forExamSendEmails() {
  // PASTE YOUR DEPLOYED WEB APP URL HERE
  const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxJpyg6KPFUMxeHSOdOVnVe4WyN6JssT9DhoufEn2pE7vIp02joOQ6jZVD-FwZCLKW7FQ/exec"; 

  try {
    PropertiesService.getDocumentProperties().deleteProperty('cancel_forExam_send');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FOR_EXAM.SHEET_NAME);
    
    if (!sheet) throw new Error('Sheet "' + FOR_EXAM.SHEET_NAME + '" not found.');

    const lastRow = sheet.getLastRow();
    if (lastRow < FOR_EXAM.START_ROW) {
      return { status: 'No applicants found', count: 0 };
    }

    const data = sheet.getRange(FOR_EXAM.START_ROW, 1, lastRow - FOR_EXAM.START_ROW + 1, FOR_EXAM.COL_EXAM_PROGRESS).getValues();
    let emailCount = 0;
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    let cancelled = false;

    for (let i = 0; i < data.length; i++) {
      if (PropertiesService.getDocumentProperties().getProperty('cancel_forExam_send') === 'true') {
        console.log('Cancellation requested for email sending');
        cancelled = true;
        break;
      }

      const row = data[i];
      const applicantName = row[0];
      const applicantLName = row[12]; // Column M
      const salutation = row[11]; // Column L
      const email = row[FOR_EXAM.COL_EMAIL - 1];
      const driveLink = row[FOR_EXAM.COL_EXAM_LINK - 1];
      const position = row[7]; // Column H
      const office = row[8]; // Column I
      const statusCell = sheet.getRange(FOR_EXAM.START_ROW + i, FOR_EXAM.COL_EXAM_PROGRESS);

      if (!applicantName || applicantName.toString().trim() === '') break;

      const currentStatus = row[FOR_EXAM.COL_EXAM_PROGRESS - 1];
      if (currentStatus && String(currentStatus).startsWith('Sent')) continue;

      if (!email || email.toString().trim() === '' || !driveLink || driveLink.toString().trim() === '') {
        statusCell.setValue('Not sent - missing email or link (' + now + ')');
        continue;
      }

      const subject = 'Job Application Update - Notice of Written Exam ' + '[' + position + ']';
      const body = 'Dear ' + salutation + ' ' + applicantLName + ',\n\n' +
        'Good day!\n\n' +
        'Thank you for your interest in the vacant position at our office. We have ' +
        'received your application and appreciate the time you took to apply.\n\n' +
        'Please see the file in the link below for your written examination details:\n\n' +
        'Link: ' + driveLink + '\n\n' +
        'Reminder: Please arrive at the site 5-10 minutes early. Late examinees ' +
        'without a valid reason will not be permitted to take the exam.\n\n' +
        'Kindly acknowledge receipt of this email. If you have any questions, please do not hesitate to contact us.\n\n' +
        'Best regards,\n' +
        'DOJ RPO V - Human Resource Unit';

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
        logSentLetter('LETTER - EXAM SCHED', position || '', office || '', applicantName || '');
        emailCount++;
      } else {
        statusCell.setValue('Error: Proxy failed (' + now + ')');
      }

      if (PropertiesService.getDocumentProperties().getProperty('cancel_forExam_send') === 'true') {
        console.log('Cancellation requested for email sending before sleep');
        cancelled = true;
        break;
      }

      // 1.5-second delay between each email to avoid rate limits
      Utilities.sleep(1500);
    }

    return { status: cancelled ? 'Cancelled' : 'Emails sent', count: emailCount, cancelled: cancelled };
  } catch (e) {
    throw new Error('Error sending email notifications: ' + e.message);
  }
}

function forExamSendIndividualEmails() {
  // PASTE YOUR DEPLOYED WEB APP URL HERE
  const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxJpyg6KPFUMxeHSOdOVnVe4WyN6JssT9DhoufEn2pE7vIp02joOQ6jZVD-FwZCLKW7FQ/exec"; 

  try {
    PropertiesService.getDocumentProperties().deleteProperty('cancel_forExam_send');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FOR_EXAM.SHEET_NAME);
    
    if (!sheet) throw new Error('Sheet "' + FOR_EXAM.SHEET_NAME + '" not found.');

    const lastRow = sheet.getLastRow();
    if (lastRow < FOR_EXAM.START_ROW) {
      return { status: 'No applicants found', count: 0 };
    }

    // Read up through the regenerate column (U)
    const data = sheet.getRange(FOR_EXAM.START_ROW, 1, lastRow - FOR_EXAM.START_ROW + 1, FOR_EXAM.COL_REGENERATE).getValues();

    const hasSelected = data.some(row => {
      const val = row[FOR_EXAM.COL_REGENERATE - 1];
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
      const regenerateVal = row[FOR_EXAM.COL_REGENERATE - 1];
      const shouldProcess = regenerateVal === true || String(regenerateVal).toLowerCase() === 'true';
      if (!shouldProcess) continue;

      if (PropertiesService.getDocumentProperties().getProperty('cancel_forExam_send') === 'true') {
        console.log('Cancellation requested for individual email sending');
        cancelled = true;
        break;
      }

      const applicantName = row[FOR_EXAM.COL_LAST_NAME - 1];
      const applicantLName = row[12]; // Column M
      const salutation = row[11]; // Column L
      const email = row[FOR_EXAM.COL_EMAIL - 1];
      const driveLink = row[FOR_EXAM.COL_EXAM_LINK - 1];
      const position = row[7]; // Column H
      const office = row[8]; // Column I
      const statusCell = sheet.getRange(FOR_EXAM.START_ROW + i, FOR_EXAM.COL_EXAM_PROGRESS);

      if (!applicantName || applicantName.toString().trim() === '') {
        statusCell.setValue('Not sent - missing name (' + now + ')');
        // clear checkbox so it won't keep attempting
        sheet.getRange(FOR_EXAM.START_ROW + i, FOR_EXAM.COL_REGENERATE).setValue(false);
        continue;
      }

      if (!email || email.toString().trim() === '' || !driveLink || driveLink.toString().trim() === '') {
        statusCell.setValue('Not sent - missing email or link (' + now + ')');
        sheet.getRange(FOR_EXAM.START_ROW + i, FOR_EXAM.COL_REGENERATE).setValue(false);
        continue;
      }

      const subject = 'Job Application Update - Notice of Written Exam ' + '[' + position + ']';
      const body = 'Dear ' + salutation + ' ' + applicantLName + ',\n\n' +
        'Good day!\n\n' +
        'Thank you for your interest in the vacant position at our office. We have ' +
        'received your application and appreciate the time you took to apply.\n\n' +
        'Please see the file in the link below for your written examination details:\n\n' +
        'Link: ' + driveLink + '\n\n' +
        'Reminder: Please arrive at the site 5-10 minutes early. Late examinees ' +
        'without a valid reason will not be permitted to take the exam.\n\n' +
        'Kindly acknowledge receipt of this email. If you have any questions, please do not hesitate to contact us.\n\n' +
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
        // Log the sent letter
        logSentLetter('LETTER - EXAM SCHED', position || '', office || '', applicantName || '');
        emailCount++;
        // clear checkbox to mark done
        sheet.getRange(FOR_EXAM.START_ROW + i, FOR_EXAM.COL_REGENERATE).setValue(false);
      } else {
        statusCell.setValue('Error: Proxy failed (' + now + ')');
      }

      if (PropertiesService.getDocumentProperties().getProperty('cancel_forExam_send') === 'true') {
        console.log('Cancellation requested for individual email sending before sleep');
        cancelled = true;
        break;
      }

      // 1.5-second delay between each email to avoid rate limits
      Utilities.sleep(1500);
    }

    return { status: cancelled ? 'Cancelled' : 'Individual emails processed', count: emailCount, cancelled: cancelled };
  } catch (e) {
    throw new Error('Error sending individual emails: ' + e.message);
  }
}

/**
 * Master function for For Exam workflow
 */
function forExamRunCompleteProcess() {
  try {
    const folderIds = forExamCreateFolders();
    const pdfResult = forExamGeneratePDFs(folderIds.forExamSubFolderId);
    
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

function cancelForExamRun() {
  PropertiesService.getDocumentProperties().setProperty('cancel_forExam_run', 'true');
}

function cancelForExamSend() {
  PropertiesService.getDocumentProperties().setProperty('cancel_forExam_send', 'true');
}

function clearCancelForExamFlags() {
  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty('cancel_forExam_run');
  props.deleteProperty('cancel_forExam_send');
}
