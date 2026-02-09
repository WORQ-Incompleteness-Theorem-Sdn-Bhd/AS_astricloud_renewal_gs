/**
 * Highlight the current month column in TRACKER
 */
function highlightCurrentMonth() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);
  
  if (!sheet) {
    Logger.log('ERROR: TRACKER sheet not found');
    return;
  }
  
  const now = new Date();
  const currentMonthYear = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMM-yyyy');
  
  // Get header row
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  // Clear previous highlighting (from column K onwards)
  const monthRange = sheet.getRange(1, CONFIG.TRACKER_COLS.FIRST_MONTH, sheet.getLastRow(), lastCol - CONFIG.TRACKER_COLS.FIRST_MONTH + 1);
  monthRange.setBackground(null);
  
  // Find and highlight current month
  for (let i = 0; i < headers.length; i++) {
    if (typeof headers[i] === 'string' && headers[i].trim() === currentMonthYear) {
      const colNumber = i + 1;
      const highlightRange = sheet.getRange(1, colNumber, sheet.getLastRow(), 1);
      highlightRange.setBackground('#00FFFF'); // Cyan color as shown in your screenshot
      
      Logger.log(`Highlighted column ${colNumber}: ${currentMonthYear}`);
      break;
    }
  }
}