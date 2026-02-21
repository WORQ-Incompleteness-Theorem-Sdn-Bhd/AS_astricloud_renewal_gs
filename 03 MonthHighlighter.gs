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
  
  // Get month headers from row 2 (row 1 = year labels, row 2 = month names like Feb-2024)
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];

  // Clear previous highlighting from FIRST_MONTH column onwards (both header rows + data)
  const monthRange = sheet.getRange(1, CONFIG.TRACKER_COLS.FIRST_MONTH, sheet.getLastRow(), lastCol - CONFIG.TRACKER_COLS.FIRST_MONTH + 1);
  monthRange.setBackground(null);

  // Find and highlight current month column (handles both string and Date-formatted headers)
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    let headerStr = '';
    if (typeof header === 'string') {
      headerStr = header.trim();
    } else if (header instanceof Date) {
      headerStr = Utilities.formatDate(header, Session.getScriptTimeZone(), 'MMM-yyyy');
    }

    if (headerStr === currentMonthYear) {
      const colNumber = i + 1;
      const highlightRange = sheet.getRange(1, colNumber, sheet.getLastRow(), 1);
      highlightRange.setBackground('#00FFFF');
      Logger.log(`Highlighted column ${colNumber}: ${currentMonthYear}`);
      break;
    }
  }
}