// ==========================================
// FAILED - Letter Generation & Backup Workflow
// ==========================================

const FAILED = {
  SHEET_NAME: 'LETTER - FAILED',
  TEMPLATE_ID: '1-08ufjwnG0bCV9nrZ6LbOwA_DDBlr8nfUNZYncwFGP8',
  COL_LAST_NAME: 1,
  COL_FIRST_NAME: 2,
  COL_EMAIL: 7,
  COL_EMAIL_DATE: 13, // Column M
  COL_LINK: 14, // Column N
  COL_STATUS: 15, // Column O
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

  let mainFolderId = props.getProperty('failedMainFolderId')
    || props.getProperty('forInterviewMainFolderId')
    || props.getProperty('unqualifiedMainFolderId')
    || props.getProperty('forExamMainFolderId');
  let mainFolder;
  if (mainFolderId) {
    try {
      mainFolder = DriveApp.getFolderById(mainFolderId);
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
    if (!props.getProperty('forExamMainFolderId')) {
      props.setProperty('forExamMainFolderId', mainFolderId);
    }
  }

  let failedFolderId = props.getProperty('failedSubFolderId');
  let failedFolder;
  if (failedFolderId) {
    try {
      failedFolder = DriveApp.getFolderById(failedFolderId);
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
    const lastName = String(rows[i][1] || '').trim();
    const firstName = String(rows[i][2] || '').trim();
    if (lastName && lastName !== '') {
      expectedNames.push(lastName + ', ' + firstName + ' - FailedLetter');
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
        message: 'Alignment verified'
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
    const applicantName = row[0];
    const email = row[FAILED.COL_EMAIL - 1];
    const driveLink = row[FAILED.COL_LINK - 1];
    const statusCell = sheet.getRange(FAILED.START_ROW + i, FAILED.COL_STATUS);

    if (!applicantName || applicantName.toString().trim() === '') continue;

    if (!email || email.toString().trim() === '' || !driveLink || driveLink.toString().trim() === '') {
      statusCell.setValue('Not sent - missing email or link (' + now + ')');
      continue;
    }

    const subject = 'JOB APPLICATION UPDATE';
    const body = 'Dear Applicant,\n\n' +
      'Good day!\n\n' +
      'Please see your failed letter document at the link below:\n\n' +
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
      emailCount++;
    } else {
      statusCell.setValue('Error: Proxy failed (' + now + ')');
    }
  }

  return { status: 'Emails sent', count: emailCount };
}

function failedGeneratePDFs(targetFolderId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FAILED.SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + FAILED.SHEET_NAME + '" not found!');

    const data = sheet.getDataRange().getDisplayValues();
    const header = data[0];
    const rows = data.slice(1).filter(row => row[0] && row[0].toString().trim() !== '');

    rows.sort((a, b) => {
      const lastNameA = String(a[1] || '').trim().toLowerCase();
      const lastNameB = String(b[1] || '').trim().toLowerCase();
      if (lastNameA !== lastNameB) return lastNameA.localeCompare(lastNameB);
      return String(a[2] || '').trim().toLowerCase().localeCompare(String(b[2] || '').trim().toLowerCase());
    });

    const templateFile = DriveApp.getFileById(FAILED.TEMPLATE_ID);
    const destinationFolder = DriveApp.getFolderById(targetFolderId);
    const processedApplicants = [];

    rows.forEach(row => {
      const lastName = String(row[1] || '').trim();
      const firstName = String(row[2] || '').trim();
      const fileName = lastName + ', ' + firstName + ' - FailedLetter';

      const copy = templateFile.makeCopy(fileName, destinationFolder);
      const doc = DocumentApp.openById(copy.getId());
      const body = doc.getBody();

      header.forEach((label, i) => {
        body.replaceText('{{' + label + '}}', row[i]);
      });

      doc.saveAndClose();
      const pdfBlob = copy.getAs(MimeType.PDF);
      const pdfFile = destinationFolder.createFile(pdfBlob).setName(fileName + '.pdf');
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      copy.setTrashed(true);

      processedApplicants.push(lastName + ', ' + firstName);
    });

    return {
      success: true,
      count: rows.length,
      applicants: processedApplicants
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
    return {
      success: true,
      folders: folderIds,
      pdfGeneration: pdfResult
    };
  } catch (e) {
    throw new Error('Error in failed complete process: ' + e.message);
  }
}
