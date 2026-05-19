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
    const backupFile = DriveApp.getFileById(ss.getId()).makeCopy(backupName, positionFolder);

    // Ensure we do NOT duplicate the linked Google Form in the backup folder.
    // If a copy of the form was created during the spreadsheet copy, trash it.
    try {
      const ORIGINAL_FORM_ID = '1t4xIKDQWg5SFQaN6f7_lStH3Ydmekv2RG2gsRtOywrg';

      const files = positionFolder.getFiles();
      while (files.hasNext()) {
        const f = files.next();
        try {
          if (f.getMimeType() === 'application/vnd.google-apps.form' && f.getId() !== ORIGINAL_FORM_ID) {
            // Trash any form in the backup folder that isn't the original form.
            // This prevents duplicating the linked Google Form.
            f.setTrashed(true);
          }
        } catch (innerErr) {
          // ignore per-file errors and continue
        }
      }
    } catch (cleanupErr) {
      // Non-fatal: don't block the backup if cleanup fails
      console.log('Form cleanup error: ' + cleanupErr.message);
    }

    return { success: true, message: 'Backup created successfully: ' + backupFile.getName(), backupUrl: backupFile.getUrl(), folderUrl: positionFolder.getUrl() };
  } catch (error) {
    return { success: false, message: error.message };
  }
}
