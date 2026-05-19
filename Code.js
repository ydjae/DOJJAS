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

/**
 * Generate Google Drive links for For Interview
 */
function interviewGenerateLinks() {
  return forInterviewGenerateLinks();
}

/**
 * Send emails to interview applicants
 */
function interviewSendEmails() {
  return forInterviewSendEmails();
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