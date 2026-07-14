/**
 * Color-code TRACKER data rows (cols B–H) by renewal urgency, matching the
 * reminder stages (3 → 1st, 2 → 2nd, 1 → 3rd, 0 → Last):
 *   Light green  — 3 months away, or Renewal Status = "Renewed"
 *   Light yellow — 2 months away
 *   Light orange — 1 month away
 *   Light red    — Contract End is this month or already past
 *   Clear        — > 3 months away, no date, Terminated, or Not Renewing
 *
 * This applies STATIC colors to cols B–H and must be re-run as dates change.
 * For self-updating color on the Contract End Date column only, use
 * setupUrgencyFormatting() (conditional formatting — no re-run needed).
 */
function highlightRenewalUrgency() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);

  if (!trackerSheet) {
    SpreadsheetApp.getUi().alert('❌ TRACKER sheet not found.');
    return;
  }

  const lastRow = trackerSheet.getLastRow();
  if (lastRow < 3) {
    SpreadsheetApp.getUi().alert('ℹ️ No data rows to highlight.');
    return;
  }

  const numDataRows   = lastRow - 2;
  const HIGHLIGHT_END = 7; // cols B–H (7 columns starting at col 2)
  const today         = new Date();
  today.setHours(0, 0, 0, 0);

  // Read cols B–H for all data rows (0-based within this 7-col slice)
  const RENEWAL_IDX  = CONFIG.TRACKER_COLS.RENEWAL_STATUS - 2; // F(6) - B(2) = 4
  const END_DATE_IDX = CONFIG.TRACKER_COLS.CONTRACT_END   - 2; // H(8) - B(2) = 6

  const data        = trackerSheet.getRange(3, 2, numDataRows, HIGHLIGHT_END).getValues();
  const backgrounds = [];

  const counts = { red: 0, orange: 0, yellow: 0, green: 0, renewed: 0, clear: 0 };

  for (const row of data) {
    const companyName   = row[0];
    const renewalStatus = row[RENEWAL_IDX];
    const contractEnd   = row[END_DATE_IDX];

    if (!companyName || companyName.toString().trim() === '') {
      backgrounds.push(Array(HIGHLIGHT_END).fill(null));
      counts.clear++;
      continue;
    }

    let color = null;
    let isRenewed = false;

    if (renewalStatus === 'Renewed') {
      color = '#D9EAD3'; // light green — already renewed (safe)
      isRenewed = true;
    } else if (renewalStatus === 'Terminated' || renewalStatus === 'Not Renewing') {
      color = null; // about to be archived or decision already made
    } else if (contractEnd && contractEnd !== '') {
      const endDate = new Date(contractEnd);
      endDate.setHours(0, 0, 0, 0);
      const months = urgencyMonthsDiff_(today, endDate);

      // Same scheme as setupUrgencyFormatting() on col H:
      // 3 → green, 2 → yellow, 1 → orange, 0 or past → red
      if (months <= 0)       color = '#F4CCCC'; // light red    — expired / expiring this month
      else if (months === 1) color = '#FCE5CD'; // light orange — 1 month away
      else if (months === 2) color = '#FFF2CC'; // light yellow — 2 months away
      else if (months === 3) color = '#D9EAD3'; // light green  — 3 months away
    }

    backgrounds.push(Array(HIGHLIGHT_END).fill(color));

    if      (isRenewed)           counts.renewed++;
    else if (color === '#F4CCCC') counts.red++;
    else if (color === '#FCE5CD') counts.orange++;
    else if (color === '#FFF2CC') counts.yellow++;
    else if (color === '#D9EAD3') counts.green++;
    else                          counts.clear++;
  }

  trackerSheet.getRange(3, 2, numDataRows, HIGHLIGHT_END).setBackgrounds(backgrounds);

  SpreadsheetApp.getUi().alert(
    `✅ Renewal urgency highlighted\n\n` +
    `🟥 Expired / expiring this month : ${counts.red}\n` +
    `🟧 1 month away                  : ${counts.orange}\n` +
    `🟨 2 months away                 : ${counts.yellow}\n` +
    `🟩 3 months away                 : ${counts.green}\n` +
    `🟩 Renewed (also green)          : ${counts.renewed}\n` +
    `⬜ No urgency / no date          : ${counts.clear}`
  );
}

/**
 * Clear urgency highlight colors from TRACKER data rows (cols B–H).
 */
function clearRenewalHighlights() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);

  if (!trackerSheet) {
    SpreadsheetApp.getUi().alert('❌ TRACKER sheet not found.');
    return;
  }

  const lastRow = trackerSheet.getLastRow();
  if (lastRow < 3) return;

  trackerSheet.getRange(3, 2, lastRow - 2, 7).setBackground(null);
  SpreadsheetApp.getUi().alert('✅ Urgency highlights cleared.');
}

/** Whole-month difference: positive if endDate is in the future. */
function urgencyMonthsDiff_(today, endDate) {
  return (endDate.getFullYear() - today.getFullYear()) * 12 +
         (endDate.getMonth()    - today.getMonth());
}

/**
 * Rebuild conditional formatting on the Contract End Date column (col H) so the
 * cell color reflects months-to-expiry, matching the reminder stages:
 *
 *   3 months left → green   (1st Reminder)
 *   2 months left → yellow  (2nd Reminder)
 *   1 month  left → orange  (3rd Reminder)
 *   0 months left → red     (Last Reminder / expiry month)
 *
 * The month difference is computed IN the rule formula against TODAY(), so the
 * colors update on their own as time passes — no need to re-run this function.
 * Rows already decided (Renewed / Not Renewing / Terminated) are left uncolored.
 *
 * Replaces any existing rules on col H (including the legacy "Pending"-based
 * rules, which broke when the Pending status was replaced by reminder stages).
 */
function setupUrgencyFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);
  if (!trackerSheet) {
    SpreadsheetApp.getUi().alert('❌ TRACKER sheet not found.');
    return;
  }

  const lastRow = Math.max(trackerSheet.getLastRow(), 3);
  const endCol  = CONFIG.TRACKER_COLS.CONTRACT_END;      // H
  const target  = trackerSheet.getRange(3, endCol, lastRow - 2, 1);
  const a1      = target.getA1Notation();

  // Whole-month difference between today's month and the end date's month,
  // mirroring getMonthsDifference() in the reminder engine:
  //   (year(H)-year(today))*12 + (month(H)-month(today))
  const H = `$${columnLetter_(endCol)}3`;
  const monthsExpr =
    `((YEAR(${H})-YEAR(TODAY()))*12+(MONTH(${H})-MONTH(TODAY())))`;

  // Guard: only color rows that have an end date and aren't already decided.
  const F = `$${columnLetter_(CONFIG.TRACKER_COLS.RENEWAL_STATUS)}3`;
  const active =
    `AND(${H}<>"",NOT(OR(${F}="Renewed",${F}="Not Renewing",${F}="Terminated")))`;

  const specs = [
    { months: 3, color: '#D9EAD3' }, // green  — 3 months left
    { months: 2, color: '#FFF2CC' }, // yellow — 2 months left
    { months: 1, color: '#FCE5CD' }, // orange — 1 month left
    { months: 0, color: '#F4CCCC' }  // red    — expiry month
  ];

  const newRules = specs.map(s =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(${active},${monthsExpr}=${s.months})`)
      .setBackground(s.color)
      .setRanges([target])
      .build()
  );

  // Keep any existing rules that don't target col H; replace the rest.
  const kept = trackerSheet.getConditionalFormatRules().filter(rule =>
    !rule.getRanges().some(r => r.getColumn() === endCol)
  );
  const removed = trackerSheet.getConditionalFormatRules().length - kept.length;

  trackerSheet.setConditionalFormatRules(kept.concat(newRules));

  SpreadsheetApp.getUi().alert(
    `✅ Urgency formatting applied to Contract End Date (${a1})\n\n` +
    `🟩 3 months left — green\n` +
    `🟨 2 months left — yellow\n` +
    `🟧 1 month left  — orange\n` +
    `🟥 Expiry month  — red\n\n` +
    `Renewed / Not Renewing / Terminated rows stay uncolored.\n` +
    `Colors update automatically as dates approach — no need to re-run.` +
    (removed > 0 ? `\n\n(${removed} previous rule(s) on this column replaced)` : '')
  );
}

/** Convert a 1-based column number to its A1 letter (1 → A, 8 → H). */
function columnLetter_(col) {
  let letter = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

