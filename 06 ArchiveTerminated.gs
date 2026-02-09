/**
 * Archive customers with "Terminate" status
 */
function archiveTerminated() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);
  const archivedSheet = ss.getSheetByName(CONFIG.ARCHIVED_SHEET);
  
  if (!trackerSheet || !archivedSheet) {
    Logger.log('ERROR: Required sheets not found');
    return;
  }
  
  const data = trackerSheet.getDataRange().getValues();
  const headers = data[0];
  let archivedCount = 0;
  
  // Process from bottom to top (to avoid index shifting when deleting rows)
  for (let i = data.length - 1; i >= 1; i--) {
    
    // Check monthly status columns for "terminate"
    let hasTerminate = false;
    
    for (let col = CONFIG.TRACKER_COLS.FIRST_MONTH - 1; col < data[i].length; col++) {
      if (data[i][col] === 'terminate') {
        hasTerminate = true;
        break;
      }
    }
    
    if (hasTerminate) {
      // Copy entire row to ARCHIVED
      archivedSheet.appendRow(data[i]);
      
      // Delete from TRACKER
      trackerSheet.deleteRow(i + 1);
      
      const companyName = data[i][CONFIG.TRACKER_COLS.COMPANY_NAME - 1];
      Logger.log(`Archived: ${companyName}`);
      archivedCount++;
    }
  }
  
  Logger.log(`Total customers archived: ${archivedCount}`);
  
  if (archivedCount > 0) {
    SpreadsheetApp.getUi().alert(`✅ Archived ${archivedCount} terminated customers`);
  }
}