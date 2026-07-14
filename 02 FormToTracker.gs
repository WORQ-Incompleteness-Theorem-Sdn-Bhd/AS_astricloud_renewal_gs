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
  const copiedCompanies = [];

  // Start from row 2 (skip header)
  for (let i = 1; i < formData.length; i++) {
    const companyName = formData[i][CONFIG.FORM_COLS.COMPANY_NAME - 1];
    const email = formData[i][CONFIG.FORM_COLS.EMAIL - 1];
    const location = formData[i][CONFIG.FORM_COLS.WORQ_LOCATION - 1];

    // Check if company exists and is not already in tracker
    if (companyName && companyName.toString().trim() !== '' && !existingCompanies.includes(companyName)) {

      // Prepare row data (columns A through H)
      // Column A (NO) is left empty — auto-populated by the SEQUENCE formula in A3
      const newRow = [
        '',                     // A - NO (handled by =SEQUENCE(COUNTA(B3:B)) formula)
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
      copiedCompanies.push(companyName);
      copiedCount++;
    }
  }

  Logger.log(`Total new entries copied: ${copiedCount}`);

  let message;
  if (copiedCount > 0) {
    const list = copiedCompanies.map(name => `• ${name}`).join('\n');
    message = `✅ Copied ${copiedCount} new entr${copiedCount === 1 ? 'y' : 'ies'} to TRACKER\n\n${list}\n\nPlease add pilot numbers to activate tracking.`;
  } else {
    message = `ℹ️ No new entries to copy.\n\nAll form responses are already in TRACKER or ARCHIVED.`;
  }
  SpreadsheetApp.getUi().alert(message);
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

  // Only process edits in TRACKER sheet
  if (sheet.getName() !== CONFIG.TRACKER_SHEET) return;

  const col = range.getColumn();
  const row = range.getRow();
  if (row <= 2) return; // Skip both header rows (row 1 = year labels, row 2 = month labels)

  // ── Handle Pilot Number edit (column E) ──────────────────────────────────
  if (col === CONFIG.TRACKER_COLS.PILOT_NUMBER) {
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
    return;
  }

  // ── Handle Contract Start date edit (column G) ───────────────────────────
  if (col === CONFIG.TRACKER_COLS.CONTRACT_START) {
    const newStartValue = range.getValue();

    // Ignore if cleared or not a valid date
    if (!newStartValue || newStartValue === '') return;
    const rawDate = new Date(newStartValue);
    if (isNaN(rawDate.getTime())) return;

    // Warn the user — this action resets all monthly paid data for the row
    const companyName = sheet.getRange(row, CONFIG.TRACKER_COLS.COMPANY_NAME).getValue() || `Row ${row}`;
    const newMonthLabel = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'MMM-yyyy');
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '⚠️ Confirm Contract Start Date Change',
      `Changing the contract start date for "${companyName}" to ${newMonthLabel} will:\n\n` +
      `• Clear ALL existing monthly paid data for this row\n` +
      `• Recalculate Contract End to 12 months from the new start\n` +
      `• Repopulate 12 months of "paid" from ${newMonthLabel}\n\n` +
      `This cannot be undone automatically. Continue?`,
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      // Restore the previous value (or clear if there was none).
      // e.oldValue for date cells is the Sheets serial number as a string (e.g. "46054").
      // Convert it back to a Date using the Sheets epoch (Dec 30, 1899, local time)
      // so setValue() receives a Date object and preserves the cell's date formatting.
      if (e.oldValue !== undefined && e.oldValue !== null && e.oldValue !== '') {
        const serial = parseFloat(e.oldValue);
        if (!isNaN(serial)) {
          const sheetsEpoch = new Date(1899, 11, 30); // Dec 30, 1899 local time
          const oldDate = new Date(sheetsEpoch.getTime() + serial * 24 * 60 * 60 * 1000);
          range.setValue(oldDate);
        } else {
          range.setValue(e.oldValue);
        }
      } else {
        range.clearContent();
      }
      return;
    }

    // Normalize to 1st of the entered month (system is month-based)
    const startDate = new Date(rawDate.getFullYear(), rawDate.getMonth(), 1);
    range.setValue(startDate);

    // Recalculate contract end (last day of the 12th month from new start)
    const endDate = new Date(startDate.getFullYear(), startDate.getMonth() + 12, 0);
    sheet.getRange(row, CONFIG.TRACKER_COLS.CONTRACT_END).setValue(endDate);

    // Clear existing monthly data for this row, then repopulate from new start
    clearMonthlyData(sheet, row);
    populate12MonthsPaid(sheet, row);

    Logger.log(`Contract start manually updated for row ${row}. Repopulated monthly paid from ${Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'MMM-yyyy')}`);
  }

  // ── Handle Contract End date edit (column H) ─────────────────────────────
  if (col === CONFIG.TRACKER_COLS.CONTRACT_END) {
    const newEndValue = range.getValue();
    if (!newEndValue || newEndValue === '') return;
    const newEndDate = new Date(newEndValue);
    if (isNaN(newEndDate.getTime())) return;

    // Normalize the cell to the last day of the entered month
    const lastDayOfMonth = new Date(newEndDate.getFullYear(), newEndDate.getMonth() + 1, 0);
    range.setValue(lastDayOfMonth);

    // Month reference (1st of month) used for fill logic comparisons
    const newEndMonth = new Date(newEndDate.getFullYear(), newEndDate.getMonth(), 1);

    const lastCol   = sheet.getLastColumn();
    const headers   = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
    const rowValues = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

    // Find the last monthly column that already has a value
    const monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    let lastFilledDate = null;

    for (let c = lastCol - 1; c >= CONFIG.TRACKER_COLS.FIRST_MONTH - 1; c--) {
      if (rowValues[c] && rowValues[c] !== '') {
        const header = headers[c];
        if (typeof header === 'string') {
          const parts = header.trim().split('-');
          if (parts.length === 2) {
            const m = monthNames.indexOf(parts[0]);
            const y = parseInt(parts[1]);
            if (m !== -1 && !isNaN(y)) { lastFilledDate = new Date(y, m, 1); break; }
          }
        } else if (header instanceof Date) {
          lastFilledDate = new Date(header.getFullYear(), header.getMonth(), 1);
          break;
        }
      }
    }

    // Determine where to start filling
    // - If existing monthly data: fill from the month AFTER the last filled month
    // - If no monthly data yet: fall back to contract start date
    let fillStart = null;
    const isExtension = !!lastFilledDate;

    if (lastFilledDate) {
      fillStart = new Date(lastFilledDate.getFullYear(), lastFilledDate.getMonth() + 1, 1);
    } else {
      const contractStart = sheet.getRange(row, CONFIG.TRACKER_COLS.CONTRACT_START).getValue();
      if (contractStart) {
        const cs = new Date(contractStart);
        fillStart = new Date(cs.getFullYear(), cs.getMonth(), 1);
      }
    }

    // Nothing to fill if the end date doesn't extend beyond existing data
    if (!fillStart || fillStart > newEndMonth) return;

    // Extend month header columns if the new end date goes beyond existing headers
    extendMonthHeaders(sheet, lastDayOfMonth);

    // Re-read headers after potential extension
    const updatedLastCol  = sheet.getLastColumn();
    const updatedHeaders  = sheet.getRange(2, 1, 1, updatedLastCol).getValues()[0];

    let current      = new Date(fillStart);
    let renewMarked  = false; // first new month of an extension gets 'renew'
    let filledCount  = 0;

    while (current <= newEndMonth) {
      const targetMonthYear = Utilities.formatDate(current, Session.getScriptTimeZone(), 'MMM-yyyy');

      const colIndex = updatedHeaders.findIndex(h => {
        if (typeof h === 'string') return h.trim() === targetMonthYear;
        if (h instanceof Date) return Utilities.formatDate(h, Session.getScriptTimeZone(), 'MMM-yyyy') === targetMonthYear;
        return false;
      });

      if (colIndex !== -1) {
        const cell = sheet.getRange(row, colIndex + 1);
        // Only fill empty cells — never overwrite existing data
        if (!cell.getValue() || cell.getValue() === '') {
          const value = (isExtension && !renewMarked) ? 'renew' : 'paid';
          cell.setValue(value);
          const rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(['paid', 'renew', 'terminate', 'not proceed'], true)
            .build();
          cell.setDataValidation(rule);
          if (value === 'renew') renewMarked = true;
          filledCount++;
        }
      }

      current = new Date(current.getFullYear(), current.getMonth() + 1, 1);
    }

    const companyName = sheet.getRange(row, CONFIG.TRACKER_COLS.COMPANY_NAME).getValue() || `Row ${row}`;
    if (filledCount > 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `${filledCount} month(s) filled for ${companyName} (${Utilities.formatDate(fillStart, Session.getScriptTimeZone(), 'MMM-yyyy')} → ${Utilities.formatDate(newEndMonth, Session.getScriptTimeZone(), 'MMM-yyyy')})`,
        '✅ Contract Extended',
        6
      );
      Logger.log(`Contract end updated for row ${row} (${companyName}): filled ${filledCount} months from ${Utilities.formatDate(fillStart, Session.getScriptTimeZone(), 'MMM-yyyy')}`);
    }
  }
}

/**
 * Clear all monthly status cells (FIRST_MONTH column onwards) for a given row.
 */
function clearMonthlyData(sheet, rowNumber) {
  const lastCol = sheet.getLastColumn();
  const firstMonthCol = CONFIG.TRACKER_COLS.FIRST_MONTH;
  if (lastCol < firstMonthCol) return;

  const numCols = lastCol - firstMonthCol + 1;
  sheet.getRange(rowNumber, firstMonthCol, 1, numCols).clearContent();
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
