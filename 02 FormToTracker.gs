/**
 * Copy new entries from Form Responses to TRACKER
 * Only copies rows that have a pilot number but aren't yet in TRACKER
 */
function copyNewEntriesToTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName(CONFIG.FORM_RESPONSES_SHEET);
  const trackerSheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);
  
  if (!formSheet || !trackerSheet) {
    Logger.log('ERROR: Required sheets not found');
    return;
  }
  
  const formData = formSheet.getDataRange().getValues();
  const trackerData = trackerSheet.getDataRange().getValues();

  // Get existing company names in tracker (skip header)
  const trackerCompanies = trackerData.slice(1).map(row => row[CONFIG.TRACKER_COLS.COMPANY_NAME - 1]);

  // Also check ARCHIVED sheet so terminated companies are not re-added
  const archivedSheet = ss.getSheetByName(CONFIG.ARCHIVED_SHEET);
  const archivedCompanies = archivedSheet
    ? archivedSheet.getDataRange().getValues().slice(1).map(row => row[CONFIG.TRACKER_COLS.COMPANY_NAME - 1])
    : [];

  const existingCompanies = [...trackerCompanies, ...archivedCompanies];
  
  let copiedCount = 0;
  
  // Start from row 2 (skip header)
  for (let i = 1; i < formData.length; i++) {
    const companyName = formData[i][CONFIG.FORM_COLS.COMPANY_NAME - 1];
    const email = formData[i][CONFIG.FORM_COLS.EMAIL - 1];
    const location = formData[i][CONFIG.FORM_COLS.WORQ_LOCATION - 1];
    
    // Check if company exists and is not already in tracker
    if (companyName && companyName.toString().trim() !== '' && !existingCompanies.includes(companyName)) {
      
      // Calculate next NO
      const lastNo = trackerData.length > 1 ? trackerData[trackerData.length - 1][CONFIG.TRACKER_COLS.NO - 1] : 0;
      const newNo = parseInt(lastNo) + 1;
      
      // Prepare row data (columns A through H)
      const newRow = [
        newNo,                  // A - NO
        companyName,            // B - Company Name
        location,               // C - WORQ Location
        email,                  // D - Company Email
        '',                     // E - Pilot Number (empty - to be filled manually)
        '',                     // F - Renewal Status
        '',                     // G - Contract Start (auto-filled on pilot number entry)
        ''                      // H - Contract End (auto-filled on pilot number entry)
      ];
      
      // Append to tracker
      trackerSheet.appendRow(newRow);
      
      Logger.log(`Copied: ${companyName} - waiting for pilot number`);
      copiedCount++;
    }
  }
  
  Logger.log(`Total new entries copied: ${copiedCount}`);
  
  if (copiedCount > 0) {
    SpreadsheetApp.getUi().alert(`✅ Copied ${copiedCount} new entries to TRACKER\n\nPlease add pilot numbers to activate tracking.`);
  }
}

/**
 * Installable trigger: runs when a cell is edited.
 * In Apps Script Triggers UI, point the "On edit" trigger to this function.
 */
function onPilotNumberEdit(e) {
  // Guard against manual execution (no event object)
  if (!e || !e.source) {
    Logger.log('onEdit must be triggered by an edit event, not run manually.');
    return;
  }

  const range = e.range;
  const sheet = range.getSheet(); // Use range.getSheet() — reliable for installable triggers
  
  // Only process edits in TRACKER sheet, column E (Pilot Number)
  if (sheet.getName() !== CONFIG.TRACKER_SHEET || range.getColumn() !== CONFIG.TRACKER_COLS.PILOT_NUMBER) {
    return;
  }
  
  const row = range.getRow();
  if (row <= 2) return; // Skip both header rows (row 1 = year labels, row 2 = month labels)
  
  const pilotNumber = range.getValue();
  
  // Check if pilot number was just added (not empty)
  if (pilotNumber && pilotNumber.toString().trim() !== '') {

    // Check for duplicate pilot number in the entire TRACKER sheet (skip header rows and current row)
    const allData = sheet.getDataRange().getValues();
    const pilotStr = pilotNumber.toString().trim();
    const duplicate = allData.some((r, i) => {
      if (i < 2) return false; // skip both header rows
      if (i === row - 1) return false; // skip current row (0-indexed)
      return r[CONFIG.TRACKER_COLS.PILOT_NUMBER - 1].toString().trim() === pilotStr;
    });

    if (duplicate) {
      SpreadsheetApp.getUi().alert(`❌ Duplicate Pilot Number\n\n"${pilotStr}" already exists in the tracker.\n\nPlease enter a unique pilot number.`);
      range.clearContent();
      return;
    }

    // Check if contract start is empty
    const contractStartCell = sheet.getRange(row, CONFIG.TRACKER_COLS.CONTRACT_START);
    
    if (!contractStartCell.getValue() || contractStartCell.getValue() === '') {
      
      // Set contract start to 1st of the current month
      const now = new Date();
      const startDate = new Date(now.getFullYear(), now.getMonth(), 1);
      contractStartCell.setValue(startDate);

      // Set contract end to last day of the 12th month (e.g. 1 Feb 2026 → 31 Jan 2027)
      const endDate = new Date(startDate.getFullYear(), startDate.getMonth() + 12, 0);
      sheet.getRange(row, CONFIG.TRACKER_COLS.CONTRACT_END).setValue(endDate);
      
      // Populate 12 months of "paid" status
      populate12MonthsPaid(sheet, row);
      
      Logger.log(`Auto-populated contract dates and payment status for row ${row}`);
    }
  }
}

/**
 * Populate 12 months of "paid" status for a given row
 */
function populate12MonthsPaid(sheet, rowNumber) {
  const contractStart = sheet.getRange(rowNumber, CONFIG.TRACKER_COLS.CONTRACT_START).getValue();
  
  if (!contractStart) {
    Logger.log(`No contract start date for row ${rowNumber}`);
    return;
  }
  
  const startDate = new Date(contractStart);
  
  // Get month headers from row 2 (row 1 = year labels, row 2 = month names like Feb-2024)
  const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Populate 12 months from contract start
  for (let monthOffset = 0; monthOffset < 12; monthOffset++) {
    const targetDate = new Date(startDate);
    targetDate.setMonth(startDate.getMonth() + monthOffset);
    
    const targetMonthYear = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'MMM-yyyy');
    
    // Find matching column header (handles both string and Date-formatted cells)
    const colIndex = headers.findIndex(header => {
      if (typeof header === 'string') {
        return header.trim() === targetMonthYear;
      }
      if (header instanceof Date) {
        return Utilities.formatDate(header, Session.getScriptTimeZone(), 'MMM-yyyy') === targetMonthYear;
      }
      return false;
    });
    
    if (colIndex !== -1) {
      // Set dropdown to "paid"
      const cell = sheet.getRange(rowNumber, colIndex + 1);
      cell.setValue('paid');
      
      // Apply dropdown data validation if not already set
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['paid', 'renew', 'terminate', 'not proceed'], true)
        .build();
      cell.setDataValidation(rule);
    }
  }
  
  Logger.log(`Populated 12 months for row ${rowNumber}`);
}

/**
 * Backfill missing paid status for rows that have contract dates but no monthly data.
 * Rows with existing monthly data are left untouched.
 */
function backfillMissingPaidStatus() {
  const trackerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.TRACKER_SHEET);
  if (!trackerSheet) {
    Logger.log('ERROR: TRACKER sheet not found');
    return;
  }

  const data = trackerSheet.getDataRange().getValues();
  let filledCount = 0;
  let skippedCount = 0;

  // Data starts at row 3 (i=2)
  for (let i = 2; i < data.length; i++) {
    const companyName  = data[i][CONFIG.TRACKER_COLS.COMPANY_NAME - 1];
    const contractStart = data[i][CONFIG.TRACKER_COLS.CONTRACT_START - 1];

    if (!contractStart) continue;

    // Check if any monthly data already exists (FIRST_MONTH column onwards)
    const monthlyData = data[i].slice(CONFIG.TRACKER_COLS.FIRST_MONTH - 1);
    const hasExistingData = monthlyData.some(val => val !== '' && val !== null && val !== undefined);

    if (hasExistingData) {
      Logger.log(`Row ${i + 1} SKIPPED — ${companyName} | Monthly data already exists`);
      skippedCount++;
      continue;
    }

    // No monthly data — populate paid from contract start
    populate12MonthsPaid(trackerSheet, i + 1);
    Logger.log(`Row ${i + 1} FILLED — ${companyName} | Populated paid from ${new Date(contractStart).toDateString()}`);
    filledCount++;
  }

  Logger.log(`Backfill complete: ${filledCount} filled, ${skippedCount} skipped (existing data preserved)`);
  SpreadsheetApp.getUi().alert(`✅ Backfill complete\n\n${filledCount} rows populated\n${skippedCount} rows skipped (existing data preserved)`);
}

/**
 * DEBUG HELPER: Logs the type and value of month column headers starting from FIRST_MONTH.
 * Run this once to confirm what format the headers are stored in.
 */
function debugHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);
  const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (let i = CONFIG.TRACKER_COLS.FIRST_MONTH - 1; i < Math.min(headers.length, CONFIG.TRACKER_COLS.FIRST_MONTH + 5); i++) {
    const h = headers[i];
    Logger.log(`Col ${i+1}: type=${typeof h}, value=${h}, isDate=${h instanceof Date}, formatted=${h instanceof Date ? Utilities.formatDate(h, Session.getScriptTimeZone(), 'MMM-yyyy') : 'N/A'}`);
  }
}