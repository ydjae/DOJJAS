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

  let mainFolderId = props.getProperty('unqualifiedMainFolderId') || props.getProperty('forExamMainFolderId');
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
    props.setProperty('unqualifiedMainFolderId', mainFolderId);
    if (!props.getProperty('forExamMainFolderId')) {
      props.setProperty('forExamMainFolderId', mainFolderId);
    }
  }

  let unqualifiedFolderId = props.getProperty('unqualifiedSubFolderId');
  let unqualifiedFolder;
  if (unqualifiedFolderId) {
    try {
      unqualifiedFolder = DriveApp.getFolderById(unqualifiedFolderId);
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
    const processedApplicants = [];

    rows.forEach(row => {
      const lastName = String(row[0] || '').trim();
      const firstName = String(row[1] || '').trim();
      const fileName = lastName + ', ' + firstName + ' - DQletter';

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
    return 'Backup successful! LETTER - DQ has been saved to the Unqualified folder.';
  } catch (e) {
    throw new Error('Error backing up DQ sheet: ' + e.message);
  }
}

function unqualifiedSendEmails() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(UNQUALIFIED.SHEET_NAME);
    if (!sheet) {
      throw new Error('Sheet "' + UNQUALIFIED.SHEET_NAME + '" not found.');
    }

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
      const statusCell = sheet.getRange(UNQUALIFIED.START_ROW + i, UNQUALIFIED.COL_STATUS);

      if (!applicantName || applicantName.toString().trim() === '') {
        continue;
      }

      if (!email || email.toString().trim() === '' || !driveLink || driveLink.toString().trim() === '') {
        statusCell.setValue('Not sent - missing email or link (' + now + ')');
        continue;
      }

      const subject = 'UPDATE ON JOB APPLICATION';
      const body = '*** Automated Message - Please Do Not Reply ***\n' +
        'For inquiries, please email: orp05.hiring@gmail.com\n\n' +
        'Dear Applicant,\n\n' +
        'Good Day,\n' +
        'Please see attached file regarding your application.\n\n' +
        'Link: ' + driveLink;

      GmailApp.sendEmail(email.toString().trim(), subject, body, { replyTo: 'orp05@doj.gov.ph' });
      statusCell.setValue('Sent (' + now + ')');
      emailCount++;
    }

    return { status: 'Emails sent', count: emailCount };
  } catch (e) {
    throw new Error('Error sending unqualified email notifications: ' + e.message);
  }
}

function unqualifiedRunCompleteProcess() {
  try {
    const folderIds = getUnqualifiedPositionFolder();
    const pdfResult = unqualifiedGeneratePDFs(folderIds.unqualifiedSubFolderId);
    return {
      success: true,
      folders: folderIds,
      pdfGeneration: pdfResult
    };
  } catch (e) {
    throw new Error('Error in unqualified complete process: ' + e.message);
  }
}
