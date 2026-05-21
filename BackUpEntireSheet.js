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
    
    // 1. Create a true, exact file copy of the spreadsheet.
    // This leaves your original sheet data completely safe and untouched!
    const backupFile = DriveApp.getFileById(ss.getId()).makeCopy(backupName, positionFolder);

    // 2. Erase the auto-duplicated form copy inside the backup folder.
    // Because Drive creates this duplicate asynchronously, we use a targeted lookup 
    // to find any Google Form in this exact destination folder that isn't the master form.
    try {
      const ORIGINAL_FORM_ID = '1t4xIKDQWg5SFQaN6f7_lStH3Ydmekv2RG2gsRtOywrg';
      
      // Short pause to allow Google Drive's backend to finish generating the file reference
      Utilities.sleep(2000); 

      // Search specifically for any Google Form inside the destination backup folder
      const files = positionFolder.searchFiles('mimeType = "application/vnd.google-apps.form"');
      while (files.hasNext()) {
        const f = files.next();
        // If it's a form file and it's NOT our live production form, destroy it immediately
        if (f.getId() !== ORIGINAL_FORM_ID) {
          console.log('Trashing auto-duplicated form file: ' + f.getName() + ' (' + f.getId() + ')');
          f.setTrashed(true);
        }
      }
    } catch (cleanupErr) {
      console.log('Form cleanup error (non-fatal): ' + cleanupErr.message);
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