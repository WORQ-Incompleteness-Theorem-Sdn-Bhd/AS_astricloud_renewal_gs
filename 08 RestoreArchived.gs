/**
 * Open the Restore Archived Customer dialog.
 * Lists all companies in the ARCHIVED sheet and lets the user pick which to restore.
 */
function showRestoreArchivedDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archivedSheet = ss.getSheetByName(CONFIG.ARCHIVED_SHEET);

  if (!archivedSheet) {
    SpreadsheetApp.getUi().alert('❌ ARCHIVED sheet not found.');
    return;
  }

  const data = archivedSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert('ℹ️ No archived customers found.');
    return;
  }

  // Build a JSON-serialisable list of archived companies (1-based index = position in data array)
  const companies = data.slice(1).map((row, i) => ({
    index: i + 1,
    name: toProperCase((row[CONFIG.TRACKER_COLS.COMPANY_NAME - 1] || '').toString()),
    contractStart: formatDateForDialog_(row[CONFIG.TRACKER_COLS.CONTRACT_START - 1]),
    contractEnd:   formatDateForDialog_(row[CONFIG.TRACKER_COLS.CONTRACT_END   - 1])
  }));

  const template = HtmlService.createTemplateFromFile('RestoreArchivedDialog');
  template.companies = JSON.stringify(companies);

  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(650).setHeight(440),
    '🔄 Restore Archived Customer'
  );
}

/**
 * Called from the HTML dialog.
 * Moves the selected rows from ARCHIVED back to TRACKER, then re-sorts TRACKER.
 *
 * @param {number[]} indices  1-based row indices within the ARCHIVED data (header = 0, first company = 1)
 */
function restoreArchivedCompanies(indices) {
  if (!indices || indices.length === 0) return;

  const ss            = SpreadsheetApp.getActiveSpreadsheet();
  const archivedSheet = ss.getSheetByName(CONFIG.ARCHIVED_SHEET);
  const trackerSheet  = ss.getSheetByName(CONFIG.TRACKER_SHEET);

  if (!archivedSheet || !trackerSheet) {
    throw new Error('Required sheets (TRACKER / ARCHIVED) not found.');
  }

  // Re-read ARCHIVED data fresh to get accurate row positions
  const archivedData = archivedSheet.getDataRange().getValues();

  // Delete from bottom to top so row numbers stay valid after each deletion
  const sortedDesc = [...indices].sort((a, b) => b - a);

  const restoredNames = [];

  for (const idx of sortedDesc) {
    const rowData = archivedData[idx];
    if (!rowData) continue;

    // Append columns B onwards to TRACKER — skip column A (SEQUENCE formula handles numbering)
    const newRow = trackerSheet.getLastRow() + 1;
    trackerSheet.getRange(newRow, 2, 1, rowData.length - 1).setValues([rowData.slice(1)]);
    restoredNames.push((rowData[CONFIG.TRACKER_COLS.COMPANY_NAME - 1] || '(unknown)').toString());

    // Remove from ARCHIVED  (sheet row = idx + 1 because header occupies row 1)
    archivedSheet.deleteRow(idx + 1);
  }

  // Re-sort TRACKER by contract start date (silent — no alert)
  sortTrackerSilently_();

  // Show confirmation
  const count = restoredNames.length;
  const list  = restoredNames.map(n => `• ${n}`).join('\n');
  SpreadsheetApp.getUi().alert(
    `✅ Restored ${count} customer${count === 1 ? '' : 's'} to TRACKER\n\n${list}\n\nTRACKER has been re-sorted by Contract Start Date.`
  );
}

// ─── Private helpers ───────────────────────────────────────────────────────────

/**
 * Format a date value (Date object, string, or empty) for display in the dialog.
 */
function formatDateForDialog_(value) {
  if (!value && value !== 0) return '';
  const d = (value instanceof Date) ? value : new Date(value);
  if (isNaN(d.getTime())) return value.toString();
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd MMM yyyy');
}

/**
 * Sort TRACKER rows by Contract Start Date without showing an alert.
 * Mirrors the logic in sortByContractStartDate() (07 SortByContractDate.gs)
 * but is used internally so the caller controls what feedback to show.
 */
function sortTrackerSilently_() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);
  if (!trackerSheet) return;

  const lastRow = trackerSheet.getLastRow();
  const lastCol = trackerSheet.getLastColumn();
  if (lastRow < 3 || lastCol < 2) return;

  const numDataRows        = lastRow - 2; // rows 3 → lastRow
  const CONTRACT_START_IDX = CONFIG.TRACKER_COLS.CONTRACT_START - 1; // 0-based

  const data = trackerSheet.getRange(3, 1, numDataRows, lastCol).getValues();

  const rowsWithDate    = [];
  const rowsWithoutDate = [];

  for (const row of data) {
    const companyName   = row[CONFIG.TRACKER_COLS.COMPANY_NAME - 1];
    const contractStart = row[CONTRACT_START_IDX];

    if (!companyName || companyName.toString().trim() === '') {
      rowsWithoutDate.push(row);
      continue;
    }

    if (contractStart && contractStart !== '') {
      rowsWithDate.push(row);
    } else {
      rowsWithoutDate.push(row);
    }
  }

  rowsWithDate.sort((a, b) => {
    const dateA = new Date(a[CONTRACT_START_IDX]);
    const dateB = new Date(b[CONTRACT_START_IDX]);
    return dateA - dateB;
  });

  const sortedData = [...rowsWithDate, ...rowsWithoutDate];

  // Write back columns B onwards — column A is the SEQUENCE formula, leave untouched
  const sortedWithoutColA = sortedData.map(row => row.slice(1));
  trackerSheet.getRange(3, 2, numDataRows, lastCol - 1).setValues(sortedWithoutColA);
}