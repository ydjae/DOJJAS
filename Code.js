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
 * Generate Google Drive links for For Exam
 */
function generateGDriveLinks(letterType) {
  if (letterType === 'forExam') {
    return forExamGenerateLinks();
  }
  if (letterType === 'unqualified') {
    return unqualifiedGenerateLinks();
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
 * Send emails to exam applicants
 */
function sendEmailsToApplicants() {
  return forExamSendEmails();
}

/**
 * Send emails to unqualified applicants
 */
function sendEmailsToUnqualified() {
  return unqualifiedSendEmails();
}

/**
 * Run complete unqualified process
 */
function runUnqualifiedCompleteProcess() {
  return unqualifiedRunCompleteProcess();
}

// ==========================================
// FOR INTERVIEW WORKFLOW - Public API
// ==========================================

/**
 * Get folder URL for For Interview
 */
function getFolderUrl(type) {
  if (type === 'forExam') {
    return forExamGetFolderUrl();
  }
  if (type === 'unqualified') {
    return unqualifiedGetFolderUrl();
  }
  if (type === 'interview') {
    return forInterviewGetFolderUrl();
  }
  throw new Error('Letter type not supported: ' + type);
}

/**
 * Show the sidebar UI
 */
function showGeneratorSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('DOJ 5 Job Application System');
  SpreadsheetApp.getUi().showSidebar(html);
}