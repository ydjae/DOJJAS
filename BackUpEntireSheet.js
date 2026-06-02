// ==========================================
// BACKUP ENTIRE SHEET - Shared Backup Utilities
// ==========================================

/**
 * Returns an existing child folder under parentFolder, or creates it if missing.
 */
function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(folderName);
}

/**
 * Backup the entire spreadsheet to the position folder containing the 4 generated folders.
 * Does not unlink the form, does not clear any source data, and explicitly trashes 
 * any auto-duplicated form copies created by Google Drive during the process.
 */
function backupWholeSheet_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const position = ss.getSheetByName(CONFIG.SHEET_NAME).getRange(CONFIG.DROPDOWN_CELL).getValue();
    if (!position) {
      throw new Error('No position selected in ' + CONFIG.SHEET_NAME + ' sheet.');
    }

    const mainFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    const positionFolder = getOrCreateFolder(mainFolder, position);

    const now = new Date();
    const backupDate = [
      String(now.getMonth() + 1).padStart(2, '0'),
      String(now.getDate()).padStart(2, '0'),
      now.getFullYear()
    ].join('-');
    const backupTime = [
      String(now.getHours()).padStart(2, '0'),
      String(now.getMinutes()).padStart(2, '0'),
      String(now.getSeconds()).padStart(2, '0')
    ].join('-');
    const backupName = `Backup_${position}_${backupDate}_${backupTime}`;
    
    // 1. Create a clean file copy of the spreadsheet.
    // Leaves your original sheet data completely safe and untouched!
    const backupFile = DriveApp.getFileById(ss.getId()).makeCopy(backupName, positionFolder);

    // 2. CRITICAL BACKGROUND CLEANUP: Erase the auto-duplicated form copy
    try {
      const ORIGINAL_FORM_ID = '1t4xIKDQWg5SFQaN6f7_lStH3Ydmekv2RG2gsRtOywrg';
      const TARGET_FORM_FOLDER_ID = '1lrgvdvE1uy9IhszAvjc0834PSeicz-M3';
      
      // Give Google Drive's asynchronous background task up to 4 seconds to finish creating the form copy
      Utilities.sleep(4000); 

      // Directly target the specific folder where Google Form copies land
      const targetFolder = DriveApp.getFolderById(TARGET_FORM_FOLDER_ID);
      const files = targetFolder.searchFiles('mimeType = "application/vnd.google-apps.form"');
      
      while (files.hasNext()) {
        const f = files.next();
        // If it's a form file inside that folder, and it's NOT our live production master form, trash it!
        if (f.getId() !== ORIGINAL_FORM_ID) {
          console.log('Automatically trashing auto-duplicated form file: ' + f.getName() + ' (' + f.getId() + ')');
          f.setTrashed(true);
        }
      }
    } catch (cleanupErr) {
      console.log('Form automatic trashing warning (non-fatal): ' + cleanupErr.message);
    }

    return { 
      success: true, 
      message: 'Backup created successfully: ' + backupFile.getName(), 
      backupUrl: backupFile.getUrl(), 
      folderUrl: positionFolder.getUrl() 
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}