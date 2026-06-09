// ==========================================
// CODE.GS - Main Entry Point & Public Functions
// ==========================================
// This file serves as the main entry point for the Google Apps Script.
// It routes calls to specific workflow modules (ForExam.gs, Unqualified.gs, etc.)
// Backend logic is organized by functionality in separate .gs files.
// ==========================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('DOJ 5 Job Application System')
    .addItem('OPEN Job Application System', 'showGeneratorSidebar')
    .addToUi();

  try {
    showGeneratorSidebar();
  } catch (e) {
    console.log('Sidebar failed to open automatically: ' + e.message);
  }
}

// ==========================================
// FOR EXAM WORKFLOW - Public API
// ==========================================

/**
 * Validate For Exam columns O-R
 */
function checkColumnsOtoR() {
  return forExamCheckColumnsOtoR();
}

/**
 * Run complete For Exam process
 */
function runCompleteProcess() {
  return forExamRunCompleteProcess();
}

/**
 * Generate Google Drive links for letter types
 */
function generateGDriveLinks(letterType) {
  if (letterType === 'forExam') {
    return forExamGenerateLinks();
  }
  if (letterType === 'unqualified') {
    return unqualifiedGenerateLinks();
  }
  if (letterType === 'failed') {
    return failedGenerateLinks();
  }
  throw new Error('Letter type not supported: ' + letterType);
}

/**
 * Get folder URL for specified letter type
 */
function getFolderUrl(letterType) {
  if (letterType === 'forExam') {
    return forExamGetFolderUrl();
  }
  if (letterType === 'unqualified') {
    return unqualifiedGetFolderUrl();
  }
  if (letterType === 'failed') {
    return failedGetFolderUrl();
  }
  if (letterType === 'interview') {
    return forInterviewGetFolderUrl();
  }
  if (letterType === 'finalNotice') {
    return finalNoticeGetFolderUrl();
  }
  throw new Error('Letter type not supported: ' + letterType);
}

/**
 * Backup Letter - Exam Sched sheet
 */
function backupExamSheet() {
  return forExamBackupSheet();
}

/**
 * Backup Letter - DQ sheet
 */
function backupDqSheet() {
  return unqualifiedBackupSheet();
}

/**
 * Backup Letter - Interview sheet
 */
function backupInterviewSheet() {
  return forInterviewBackupSheet();
}

/**
 * Backup Letter - Failed sheet
 */
function backupFailedSheet() {
  return failedBackupSheet();
}

/**
 * Backup the entire spreadsheet to Drive with timestamp
 */
function backupWholeSheet() {
  return backupWholeSheet_();
}

/**
 * Run complete For Interview process
 */
function runInterviewCompleteProcess() {
  return forInterviewRunCompleteProcess();
}

/**
 * Run complete Failed process
 */
function runFailedCompleteProcess() {
  return failedRunCompleteProcess();
}

/**
 * Verify Failed PDF alignment
 */
function verifyFailedAlignment() {
  return failedVerifyAlignment();
}

/**
 * Send emails to failed applicants
 */
function sendEmailsToFailedApplicants() {
  return failedSendEmails();
}

function sendSelectedFailedEmails() {
  return failedSendSelectedEmails();
}

function generateIndividualFailedPDFs() {
  return failedGenerateIndividualPDFs();
}

/**
 * Generate Google Drive links for For Interview
 */
function interviewGenerateLinks() {
  return forInterviewGenerateLinks();
}

/**
 * Generate selected For Interview PDFs
 */
function generateIndividualInterviewPDFs() {
  return forInterviewGenerateIndividualPDFs();
}

/**
 * Send emails to interview applicants
 */
function interviewSendEmails() {
  return forInterviewSendEmails();
}

function sendSelectedInterviewEmails() {
  return forInterviewSendSelectedEmails();
}

/**
 * Send emails to exam applicants
 */
function sendEmailsToApplicants() {
  return forExamSendEmails();
}

function sendIndividualEmailsToApplicants() {
  return forExamSendIndividualEmails();
}

function sendEmailsToUnqualified() {
  return unqualifiedSendEmails();
}

function sendSelectedEmailsToUnqualified() {
  return unqualifiedSendSelectedEmails();
}

function generateIndividualUnqualifiedPDFs() {
  return unqualifiedGenerateIndividualPDFs();
}

/**
 * Run complete unqualified process
 */
function runUnqualifiedCompleteProcess() {
  return unqualifiedRunCompleteProcess();
}

/**
 * Generate SUMMARY rows from LETTER - EXAM SCHED and LETTER - DQ
 */
function generateSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName('SUMMARY');
  const examSheet = ss.getSheetByName('LETTER - EXAM SCHED');
  const dqSheet = ss.getSheetByName('LETTER - DQ');

  if (!summarySheet) {
    throw new Error('Sheet "SUMMARY" not found.');
  }
  if (!examSheet) {
    throw new Error('Sheet "LETTER - EXAM SCHED" not found.');
  }
  if (!dqSheet) {
    throw new Error('Sheet "LETTER - DQ" not found.');
  }

  const startRow = 12;
  const writeColumns = 4; // A:D
  const lastSummaryRow = summarySheet.getLastRow();

  if (lastSummaryRow >= startRow) {
    summarySheet.getRange(startRow, 1, lastSummaryRow - startRow + 1, writeColumns).clearContent();
  }

  const examLastRow = examSheet.getLastRow();
  let summaryValues = [];

  // --- Process EXAM entries ---
  if (examLastRow >= 2) {
    // grab columns A-F so we can build name from A-D and get address (E) and exam info (F)
    const examData = examSheet.getRange(2, 1, examLastRow - 1, 6).getValues();
    const examEntries = [];
    examData.forEach(row => {
      const lastName = row[0] || '';
      const otherParts = [row[1], row[2], row[3]].filter(c => c != null && String(c).trim() !== '');
      const fullName = lastName ? (String(lastName).trim() + (otherParts.length ? ', ' + otherParts.join(' ') : '')) : otherParts.join(' ').trim();
      if (fullName !== '') {
        const address = row[4] || '';
        const examInfo = row[5] || '';
        examEntries.push({ lastName: String(lastName).toLowerCase(), row: [fullName, address, examInfo, 'FOR EXAM'] });
      }
    });

    // sort by last name (column A)
    examEntries.sort((a, b) => {
      if (a.lastName < b.lastName) return -1;
      if (a.lastName > b.lastName) return 1;
      return 0;
    });

    examEntries.forEach(e => summaryValues.push(e.row));
  }

  // --- Process DQ entries (group by reason) ---
  const dqLastRow = dqSheet.getLastRow();
  if (dqLastRow >= 2) {
    // need columns A-F and J (we'll fetch up to column 10 to include J)
    const dqData = dqSheet.getRange(2, 1, dqLastRow - 1, 10).getValues();
    const groups = {}; // reason -> array of rows

    dqData.forEach(row => {
      const lastName = row[0] || '';
      const otherParts = [row[1], row[2], row[3]].filter(c => c != null && String(c).trim() !== '');
      const fullName = lastName ? (String(lastName).trim() + (otherParts.length ? ', ' + otherParts.join(' ') : '')) : otherParts.join(' ').trim();
      if (fullName !== '') {
        const address = row[4] || '';
        const age = row[5] || '';
        const reason = row[9] || '';
        if (!groups[reason]) groups[reason] = [];
        groups[reason].push([fullName, address, age, reason]);
      }
    });

    // iterate reasons in sorted order for consistent grouping
    const reasonKeys = Object.keys(groups).sort((a, b) => {
      if (a == b) return 0;
      if (a === '') return 1; // push empty reasons to end
      if (b === '') return -1;
      return a.toLowerCase() < b.toLowerCase() ? -1 : 1;
    });

    reasonKeys.forEach(reason => {
      // optional: could add a separator/header row. For now just append grouped rows.
      groups[reason].forEach(r => summaryValues.push(r));
    });
  }

  if (summaryValues.length > 0) {
    summarySheet.getRange(startRow, 1, summaryValues.length, writeColumns).setValues(summaryValues);
  }

  return { rowsWritten: summaryValues.length };
}

/**
 * Sort SUMMARY starting at row 12 by column D priority and then by name.
 */
function sortSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('SUMMARY');
  if (!sheet) throw new Error('Sheet "SUMMARY" not found.');

  const startRow = 12;
  const lastRow = sheet.getLastRow();
  const numCols = 4;
  if (lastRow < startRow) return { rowsSorted: 0 };

  const range = sheet.getRange(startRow, 1, lastRow - startRow + 1, numCols);
  const values = range.getValues();

  const priorityMap = {
    'FOR INTERVIEW': 1,
    'FAILED': 2,
    'FOR EXAM': 3,
    'PDS NOT NOTARIZED': 4
  };

  values.sort((a, b) => {
    const va = a[3] ? String(a[3]).toUpperCase().trim() : '';
    const vb = b[3] ? String(b[3]).toUpperCase().trim() : '';
    const pa = (va === '') ? 99 : (priorityMap[va] || 5);
    const pb = (vb === '') ? 99 : (priorityMap[vb] || 5);
    if (pa !== pb) return pa - pb;
    const na = a[0] ? String(a[0]).toLowerCase() : '';
    const nb = b[0] ? String(b[0]).toLowerCase() : '';
    if (na < nb) return -1;
    if (na > nb) return 1;
    return 0;
  });

  range.setValues(values);
  return { rowsSorted: values.length };
}

/**
 * Set the active spreadsheet sheet by name.
 */
function setActiveSheetByName(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Sheet "' + sheetName + '" not found.');
  }
  ss.setActiveSheet(sheet);
  return true;
}

/**
 * Show the sidebar UI
 */
function showGeneratorSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('DOJ 5 Job Application System');
  SpreadsheetApp.getUi().showSidebar(html);
}

function checkFailedColumnL() {
  return checkFailedColumnM();
}