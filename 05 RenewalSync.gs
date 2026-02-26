/**
 * Sync renewals from TRACKER col F (Renewal Status).
 * Shows a preview of all changes and asks for confirmation before applying.
 * When Renewal Status = "Renew", extends contract by 12 months.
 * When Renewal Status = "Not Renewing", marks as terminated.
 */
function syncRenewals() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);
  const ui           = SpreadsheetApp.getUi();
  const tz           = Session.getScriptTimeZone();

  if (!trackerSheet) {
    Logger.log('ERROR: TRACKER sheet not found');
    return;
  }

  const trackerData = trackerSheet.getDataRange().getValues();

  // ── Pass 1: scan rows, build preview lists ─────────────────────────────────
  const toRenew     = []; // { rowIdx, companyName, companyEmail, pilotNumber, currentEndDate, newStartDate, newEndDate }
  const toTerminate = []; // { rowIdx, companyName, companyEmail, pilotNumber, currentEndDate }

  for (let i = 2; i < trackerData.length; i++) {
    const companyName   = trackerData[i][CONFIG.TRACKER_COLS.COMPANY_NAME   - 1];
    const companyEmail  = trackerData[i][CONFIG.TRACKER_COLS.COMPANY_EMAIL  - 1];
    const pilotNumber   = trackerData[i][CONFIG.TRACKER_COLS.PILOT_NUMBER   - 1];
    const renewalStatus = trackerData[i][CONFIG.TRACKER_COLS.RENEWAL_STATUS - 1];
    const contractEnd   = trackerData[i][CONFIG.TRACKER_COLS.CONTRACT_END   - 1];
    const worqLocation  = trackerData[i][CONFIG.TRACKER_COLS.WORQ_LOCATION  - 1];

    if (!contractEnd) continue;

    const currentEndDate = new Date(contractEnd);

    if (renewalStatus === 'Not Renewing') {
      toTerminate.push({ rowIdx: i, companyName, companyEmail, pilotNumber, currentEndDate, worqLocation });
    } else if (renewalStatus === 'Renew') {
      const newStartDate = new Date(currentEndDate.getFullYear(), currentEndDate.getMonth() + 1, 1);
      const newEndDate   = new Date(currentEndDate);
      newEndDate.setFullYear(newEndDate.getFullYear() + 1);
      toRenew.push({ rowIdx: i, companyName, companyEmail, pilotNumber, currentEndDate, newStartDate, newEndDate, worqLocation });
    }
  }

  if (toRenew.length === 0 && toTerminate.length === 0) {
    ui.alert('ℹ️ Nothing to sync.\n\nNo companies are marked "Renew" or "Not Renewing".');
    return;
  }

  // ── Preview dialog ─────────────────────────────────────────────────────────
  const fmt = date => Utilities.formatDate(date, tz, 'dd MMM yyyy');

  let previewMsg = '📋 The following changes will be applied:\n';

  if (toRenew.length > 0) {
    previewMsg += `\nRenewals to extend (${toRenew.length}):\n`;
    previewMsg += toRenew.map(r =>
      `• ${r.companyName}\n  ${fmt(r.currentEndDate)} → ${fmt(r.newEndDate)}`
    ).join('\n');
  }

  if (toTerminate.length > 0) {
    previewMsg += `\n\nTerminations to process (${toTerminate.length}):\n`;
    previewMsg += toTerminate.map(r =>
      `• ${r.companyName}  (ends ${fmt(r.currentEndDate)})`
    ).join('\n');
  }

  const emailCount = [...toRenew, ...toTerminate].filter(r => r.companyEmail).length;
  if (emailCount > 0) {
    previewMsg += `\n\n📧 ${emailCount} confirmation email(s) will be sent to customers.`;
  }
  previewMsg += `\n📨 1 vendor notification will be sent to ${CONFIG.VENDOR_EMAIL}.`;

  previewMsg += '\n\nProceed?';

  const response = ui.alert('Sync Renewals — Preview', previewMsg, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) {
    ui.alert('ℹ️ Sync cancelled. No changes were made.');
    return;
  }

  // ── Pass 2: apply changes ──────────────────────────────────────────────────
  let renewalsProcessed = 0;
  const renewedCompanies    = [];
  const terminatedCompanies = [];
  let emailsSent = 0;

  for (const r of toTerminate) {
    markMonthCell(trackerSheet, r.rowIdx + 1, r.currentEndDate, 'terminate');
    trackerSheet.getRange(r.rowIdx + 1, CONFIG.TRACKER_COLS.RENEWAL_STATUS).setValue('Terminated');
    if (r.companyEmail) {
      sendTerminationEmail(r.companyName, r.companyEmail, r.pilotNumber, r.currentEndDate, r.worqLocation);
      emailsSent++;
    }
    terminatedCompanies.push(r.companyName);
    Logger.log(`Marked ${r.companyName} as Terminated. Run [05] Archive Terminated Customers when ready.`);
  }

  for (const r of toRenew) {
    // New tenure: starts 1st of month after current end, ends 12 months later
    trackerSheet.getRange(r.rowIdx + 1, CONFIG.TRACKER_COLS.CONTRACT_END).setValue(r.newEndDate);
    extendMonthHeaders(trackerSheet, r.newEndDate);
    populate12MonthsFromDate(trackerSheet, r.rowIdx + 1, r.newStartDate);
    if (r.companyEmail) {
      sendRenewalConfirmationEmail(r.companyName, r.companyEmail, r.pilotNumber, r.newStartDate, r.newEndDate, r.worqLocation);
      emailsSent++;
    }
    trackerSheet.getRange(r.rowIdx + 1, CONFIG.TRACKER_COLS.RENEWAL_STATUS).setValue('Renewed');

    const newStartStr = Utilities.formatDate(r.newStartDate, tz, 'MMM yyyy');
    const newEndStr   = Utilities.formatDate(r.newEndDate,   tz, 'MMM yyyy');
    renewedCompanies.push(`${r.companyName} (${newStartStr} – ${newEndStr})`);
    Logger.log(`Extended contract for ${r.companyName}: ${r.newStartDate.toDateString()} – ${r.newEndDate.toDateString()}`);
    renewalsProcessed++;
  }

  Logger.log(`Renewals processed: ${renewalsProcessed}`);

  // ── Vendor notification ────────────────────────────────────────────────────
  sendVendorNotificationEmail(toRenew, toTerminate);

  // ── Summary alert ──────────────────────────────────────────────────────────
  const parts = [];
  if (renewedCompanies.length > 0) {
    parts.push(`Renewals extended: ${renewedCompanies.length}\n${renewedCompanies.map(n => `• ${n}`).join('\n')}`);
  }
  if (terminatedCompanies.length > 0) {
    parts.push(`Terminated: ${terminatedCompanies.length}\n${terminatedCompanies.map(n => `• ${n}`).join('\n')}`);
  }

  const emailLine = emailsSent > 0 ? `\n\nCustomer emails sent: ${emailsSent}` : '';
  ui.alert(`✅ Sync complete\n\n${parts.join('\n\n')}${emailLine}\nVendor notification sent to ${CONFIG.VENDOR_EMAIL}`);
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
 * Extend TRACKER month header columns (rows 1 & 2) up to the given date.
 * Row 2: adds MMM-yyyy labels. Row 1: adds year number at the first month of each new year.
 * No-ops if headers already cover the required range.
 */
function extendMonthHeaders(sheet, upToDate) {
  const monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];

  // Find the last month column by scanning backwards
  let lastMonthDate = null;
  let lastMonthCol = -1; // 1-based column number

  for (let i = headers.length - 1; i >= CONFIG.TRACKER_COLS.FIRST_MONTH - 1; i--) {
    const header = headers[i];
    let parsed = null;

    if (typeof header === 'string' && header.trim() !== '') {
      const parts = header.trim().split('-');
      if (parts.length === 2) {
        const m = monthNames.indexOf(parts[0]);
        const y = parseInt(parts[1]);
        if (m !== -1 && !isNaN(y)) parsed = new Date(y, m, 1);
      }
    } else if (header instanceof Date) {
      parsed = new Date(header.getFullYear(), header.getMonth(), 1);
    }

    if (parsed) {
      lastMonthDate = parsed;
      lastMonthCol = i + 1; // convert to 1-based
      break;
    }
  }

  if (!lastMonthDate) {
    Logger.log('extendMonthHeaders: Could not find last month header');
    return;
  }

  const upToMonth = new Date(upToDate.getFullYear(), upToDate.getMonth(), 1);
  if (upToMonth <= lastMonthDate) {
    Logger.log('extendMonthHeaders: Headers already cover up to ' + Utilities.formatDate(upToMonth, Session.getScriptTimeZone(), 'MMM-yyyy'));
    return;
  }

  // Append new month columns one by one
  let current = new Date(lastMonthDate.getFullYear(), lastMonthDate.getMonth() + 1, 1);
  let colNumber = lastMonthCol + 1;
  let added = 0;

  while (current <= upToMonth) {
    const monthLabel = Utilities.formatDate(current, Session.getScriptTimeZone(), 'MMM-yyyy');

    // Row 2: month label (e.g. Jan-2028)
    sheet.getRange(2, colNumber).setValue(monthLabel);

    // Row 1: year label only on the first month of each new year (January)
    if (current.getMonth() === 0) {
      sheet.getRange(1, colNumber).setValue(current.getFullYear());
    }

    Logger.log('extendMonthHeaders: Added column ' + colNumber + ' → ' + monthLabel);
    current = new Date(current.getFullYear(), current.getMonth() + 1, 1);
    colNumber++;
    added++;
  }

  Logger.log('extendMonthHeaders: Added ' + added + ' new column(s) up to ' + Utilities.formatDate(upToMonth, Session.getScriptTimeZone(), 'MMM-yyyy'));
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