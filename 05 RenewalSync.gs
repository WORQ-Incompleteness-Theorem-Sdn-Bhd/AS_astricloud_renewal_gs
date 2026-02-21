/**
 * Sync renewals from TRACKER col F (Renewal Status)
 * When Renewal Status = "Renew", extend contract by 12 months
 */
function syncRenewals() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);

  if (!trackerSheet) {
    Logger.log('ERROR: TRACKER sheet not found');
    return;
  }

  const trackerData = trackerSheet.getDataRange().getValues();
  let renewalsProcessed = 0;

  // Skip header rows (data starts row 3, i=2)
  for (let i = 2; i < trackerData.length; i++) {
    const companyName   = trackerData[i][CONFIG.TRACKER_COLS.COMPANY_NAME - 1];
    const companyEmail  = trackerData[i][CONFIG.TRACKER_COLS.COMPANY_EMAIL - 1];
    const pilotNumber   = trackerData[i][CONFIG.TRACKER_COLS.PILOT_NUMBER - 1];
    const renewalStatus = trackerData[i][CONFIG.TRACKER_COLS.RENEWAL_STATUS - 1];
    const contractEnd   = trackerData[i][CONFIG.TRACKER_COLS.CONTRACT_END - 1];

    if (!contractEnd) continue;

    const currentEndDate = new Date(contractEnd);

    // --- Handle Not Renewing ---
    if (renewalStatus === 'Not Renewing') {
      markMonthCell(trackerSheet, i + 1, currentEndDate, 'terminate');
      trackerSheet.getRange(i + 1, CONFIG.TRACKER_COLS.RENEWAL_STATUS).setValue('Terminated');
      if (companyEmail) {
        sendTerminationEmail(companyName, companyEmail, pilotNumber, currentEndDate);
      }
      Logger.log(`Marked ${companyName} as Terminated. Termination email sent. Run Archive Terminated Customers when ready.`);
      continue;
    }

    if (renewalStatus !== 'Renew') continue;

    // New tenure: starts 1st of the month after current end, ends 12 months later
    const newStartDate = new Date(currentEndDate.getFullYear(), currentEndDate.getMonth() + 1, 1);
    const newEndDate   = new Date(currentEndDate);
    newEndDate.setFullYear(newEndDate.getFullYear() + 1);

    // Update contract end date in TRACKER
    trackerSheet.getRange(i + 1, CONFIG.TRACKER_COLS.CONTRACT_END).setValue(newEndDate);

    // Populate another 12 months — use day 1 to avoid month-end overflow (e.g. Mar 31 + 1 month = May 1, not Apr 1)
    populate12MonthsFromDate(trackerSheet, i + 1, newStartDate);

    // Send thank you email with new tenure
    if (companyEmail) {
      sendRenewalConfirmationEmail(companyName, companyEmail, pilotNumber, newStartDate, newEndDate);
    }

    // Mark as Renewed in col F
    trackerSheet.getRange(i + 1, CONFIG.TRACKER_COLS.RENEWAL_STATUS).setValue('Renewed');

    Logger.log(`Extended contract for ${companyName}: ${newStartDate.toDateString()} – ${newEndDate.toDateString()}`);
    renewalsProcessed++;
  }

  Logger.log(`Renewals processed: ${renewalsProcessed}`);

  if (renewalsProcessed > 0) {
    SpreadsheetApp.getUi().alert(`✅ Processed ${renewalsProcessed} renewals`);
  }
}

/**
 * Populate 12 months from a specific start date
 */
function populate12MonthsFromDate(sheet, rowNumber, startDate) {
  // Month headers are in row 2 (row 1 = year labels)
  const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];

  for (let monthOffset = 0; monthOffset < 12; monthOffset++) {
    const targetDate = new Date(startDate);
    targetDate.setMonth(startDate.getMonth() + monthOffset);

    const targetMonthYear = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'MMM-yyyy');

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
      const cell = sheet.getRange(rowNumber, colIndex + 1);
      // First month of renewal cycle = 'renew', remaining months = 'paid'
      cell.setValue(monthOffset === 0 ? 'renew' : 'paid');

      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['paid', 'renew', 'terminate', 'not proceed'], true)
        .build();
      cell.setDataValidation(rule);
    } else {
      Logger.log(`Column not found for ${targetMonthYear} — header may be missing or out of range`);
    }
  }
}

/**
 * Set a specific month column cell to a given value for a row
 */
function markMonthCell(sheet, rowNumber, targetDate, value) {
  const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  const targetMonthYear = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'MMM-yyyy');

  const colIndex = headers.findIndex(header => {
    if (typeof header === 'string') return header.trim() === targetMonthYear;
    if (header instanceof Date) return Utilities.formatDate(header, Session.getScriptTimeZone(), 'MMM-yyyy') === targetMonthYear;
    return false;
  });

  if (colIndex !== -1) {
    const cell = sheet.getRange(rowNumber, colIndex + 1);
    cell.setValue(value);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['paid', 'renew', 'terminate', 'not proceed'], true)
      .build();
    cell.setDataValidation(rule);
    Logger.log(`Set ${targetMonthYear} → '${value}' for row ${rowNumber}`);
  } else {
    Logger.log(`Column not found for ${targetMonthYear} — could not mark as '${value}'`);
  }
}