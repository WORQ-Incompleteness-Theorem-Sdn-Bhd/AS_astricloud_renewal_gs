/**
 * Find TRACKER rows whose Contract End Date has passed with no recorded renewal
 * decision — customers who quietly lapsed without a "Renew" or "Not Renewing"
 * update in col F. This includes rows stuck at "Last Reminder Sent": the
 * reminder ladder ends at the final notice and never auto-terminates, so a
 * non-responding customer lands here and needs a manual decision.
 *
 * Shared detection used by both findLapsedContracts() (menu, on demand) and the
 * monthly reminder run summary — one rule, so the two can never disagree.
 *
 * @returns {Array<{rowNum:number, companyName:string, renewalStatus:string,
 *                  endDate:Date, daysLapsed:number}>} oldest lapse first
 */
function getLapsedContracts_() {
  const trackerSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.TRACKER_SHEET);
  if (!trackerSheet) return [];

  const data  = trackerSheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Statuses that mean the company has already been handled — skip these
  const HANDLED = new Set(['Renewed', 'Terminated', 'Not Renewing']);

  const lapsed = [];

  for (let i = 2; i < data.length; i++) {
    const companyName   = data[i][CONFIG.TRACKER_COLS.COMPANY_NAME   - 1];
    const renewalStatus = (data[i][CONFIG.TRACKER_COLS.RENEWAL_STATUS - 1] || '').toString().trim();
    const contractEnd   = data[i][CONFIG.TRACKER_COLS.CONTRACT_END   - 1];

    if (!companyName || companyName.toString().trim() === '') continue;
    if (!contractEnd  || contractEnd  === '')                  continue;
    if (HANDLED.has(renewalStatus))                            continue;

    const endDate = new Date(contractEnd);
    endDate.setHours(0, 0, 0, 0);
    if (endDate >= today) continue;

    lapsed.push({
      rowNum: i + 1,
      companyName: companyName.toString().trim(),
      renewalStatus: renewalStatus,
      endDate: endDate,
      daysLapsed: Math.floor((today - endDate) / (1000 * 60 * 60 * 24))
    });
  }

  // Longest-lapsed first — most urgent at the top
  lapsed.sort((a, b) => b.daysLapsed - a.daysLapsed);
  return lapsed;
}

/**
 * Scan TRACKER for rows whose Contract End Date has passed but have no recorded
 * renewal decision — i.e. customers who quietly lapsed without a "Renew" or
 * "Not Renewing" update in col F.
 */
function findLapsedContracts() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);

  if (!trackerSheet) {
    SpreadsheetApp.getUi().alert('❌ TRACKER sheet not found.');
    return;
  }

  const tz = Session.getScriptTimeZone();

  const lapsed = getLapsedContracts_().map(c => {
    const endStr     = Utilities.formatDate(c.endDate, tz, 'dd MMM yyyy');
    const statusNote = c.renewalStatus ? ` [${c.renewalStatus}]` : '';
    return `• ${c.companyName}${statusNote} — ended ${endStr} (${c.daysLapsed}d ago)`;
  });

  if (lapsed.length === 0) {
    SpreadsheetApp.getUi().alert(
      '✅ No lapsed contracts found.\n\n' +
      'All customers with a past Contract End Date have a recorded renewal decision.'
    );
    return;
  }

  // Truncate display to 20 entries to avoid alert overflow
  const MAX_DISPLAY = 20;
  const displayList = lapsed.slice(0, MAX_DISPLAY);
  const overflow    = lapsed.length > MAX_DISPLAY ? `\n\n...and ${lapsed.length - MAX_DISPLAY} more.` : '';

  SpreadsheetApp.getUi().alert(
    `⚠️ ${lapsed.length} lapsed contract${lapsed.length === 1 ? '' : 's'} found\n\n` +
    `These companies have a past Contract End Date with no renewal decision:\n\n` +
    displayList.join('\n') +
    overflow +
    `\n\nNext step: set their Renewal Status (col F) to "Renew" or "Not Renewing", then run [04] Sync Renewals.`
  );
}
