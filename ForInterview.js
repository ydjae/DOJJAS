// ==========================================
// FOR INTERVIEW - Letter Generation & Email Workflow
// ==========================================

// Constants for For Interview workflow
const FOR_INTERVIEW = {
  SHEET_NAME: 'LETTER - FOR INTERVIEW',
  TEMPLATE_ID: '1V0icY2qWob_D2LQpe24H5Q85p5QXnfHgwlLt8sWeSNs', // Same template as exam for now
  COL_LAST_NAME: 1,
  COL_FIRST_NAME: 2,
  COL_EMAIL: 7,
  COL_TEMPLATE_COLS_START: 15, // Column O
  COL_TEMPLATE_COLS_END: 18,   // Column R
  COL_INTERVIEW_LINK: 19,      // Column S
  COL_INTERVIEW_PROGRESS: 20,  // Column T
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
    const mainFolder = parentFolder.createFolder(folderName);
    const mainFolderId = mainFolder.getId();

    const forInterviewSubFolder = mainFolder.createFolder("For Interview");
    const forInterviewSubFolderId = forInterviewSubFolder.getId();

    const props = PropertiesService.getDocumentProperties();
    props.setProperty('forInterviewMainFolderId', mainFolderId);
    props.setProperty('forInterviewSubFolderId', forInterviewSubFolderId);

    return {
      mainFolderId: mainFolderId,
      forInterviewSubFolderId: forInterviewSubFolderId,
      folderUrl: mainFolder.getUrl()
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
    const processedApplicants = [];

    rows.forEach((row, index) => {
      const lastName = String(row[0] || "").trim();
      const firstName = String(row[1] || "").trim();
      const fileName = lastName + ", " + firstName + " - Interviewletter";

      const copy = templateFile.makeCopy(fileName, destinationFolder);
      const doc = DocumentApp.openById(copy.getId());
      const body = doc.getBody();

      header.forEach((label, i) => {
        body.replaceText('{{' + label + '}}', row[i]);
      });

      doc.saveAndClose();
      const pdfBlob = copy.getAs(MimeType.PDF);
      const pdfFile = destinationFolder.createFile(pdfBlob).setName(fileName + ".pdf");
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
    throw new Error('Error generating PDFs: ' + e.message);
  }
}

/**
 * Generate Google Drive links and insert into spreadsheet
 */
function interviewGenerateLinks() {
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
function backupInterviewSheet() {
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

    return 'Backup successful! LETTER - FOR INTERVIEW has been saved to the For Interview folder.';
  } catch (e) {
    throw new Error('Error backing up sheet: ' + e.message);
  }
}

/**
 * Send emails to interview applicants
 */
function interviewSendEmails() {
  // PASTE YOUR DEPLOYED WEB APP URL HERE
  const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwSQQYBB9G0stw13NoVgm-FaRcgsRBe6KBGbXG-RQrwfYMQP0VkMS_eObHUMbr8l0HZ/exec";

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
      const email = row[FOR_INTERVIEW.COL_EMAIL - 1];
      const driveLink = row[FOR_INTERVIEW.COL_INTERVIEW_LINK - 1];
      const statusCell = sheet.getRange(FOR_INTERVIEW.START_ROW + i, FOR_INTERVIEW.COL_INTERVIEW_PROGRESS);

      if (!applicantName || applicantName.toString().trim() === '') continue;

      if (!email || email.toString().trim() === '' || !driveLink || driveLink.toString().trim() === '') {
        statusCell.setValue('Not sent - missing email or link (' + now + ')');
        continue;
      }

      const subject = 'NOTICE OF INTERVIEW';
      const body = '*** Automated Message - Please Do Not Reply ***\n' +
        'For inquiries, please email: orp05.hiring@gmail.com\n\n' +
        'Dear Applicant,\n\n' +
        'Good day!\n\n' +
        'Congratulations! You have passed the written examination and have been selected to proceed to the interview stage.\n\n' +
        'Please see the file in the link below for your interview details:\n\n' +
        'Link: ' + driveLink + '\n\n' +
        'Please arrive at the interview site 5-10 minutes early. We look forward to meeting you!\n\n' +
        'Best regards,\n' +
        'HR Recruitment Team';

      // --- INTEGRATED PROXY CALL ---
      const payload = {
        recipient: email.toString().trim(),
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

/**
 * Master function for For Interview workflow
 */
function runInterviewCompleteProcess() {
  try {
    const folderIds = forInterviewCreateFolders();
    const pdfResult = forInterviewGeneratePDFs(folderIds.forInterviewSubFolderId);
    return {
      success: true,
      folders: folderIds,
      pdfGeneration: pdfResult
    };
  } catch (e) {
    throw new Error('Error in complete process: ' + e.message);
  }
}