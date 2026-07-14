/**
 * Menu / manual entry point — shows a UI alert with the run summary.
 * The scheduled trigger uses runRenewalReminders() directly (no UI).
 */
function checkAndSendReminders() {
  const summary = runRenewalReminders();
  SpreadsheetApp.getUi().alert(summary.uiMessage);
}

/**
 * Scheduled entry point — installed via setupAutoReminderTrigger() to run on
 * the 1st of every month. Wraps the run in error alerting (unattended) and
 * emails a summary of what was sent so ops has passive confidence it ran.
 */
function monthlyRenewalReminders() {
  try {
    const summary = runRenewalReminders();
    sendReminderRunSummary_(summary);
  } catch (e) {
    notifyError_('monthlyRenewalReminders', e);
    throw e; // still surface in the Apps Script execution log
  }
}

/**
 * Core reminder engine (no UI, no email-summary side effects — returns a result
 * object). Escalates each awaiting-reply company through the reminder stages as
 * expiry approaches: 3 months → 1st, 2 → 2nd, 1 → Last. A stage is applied at
 * most once and status is never downgraded, so re-running within the same month
 * is a safe no-op. Admin-set 'Renew' / 'Renewed' / 'Not Renewing' are left alone.
 *
 * @returns {{sent:Array, skipped:number, remindersSent:number, uiMessage:string}}
 */
function runRenewalReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);

  if (!trackerSheet) {
    Logger.log('ERROR: TRACKER sheet not found');
    return { sent: [], skipped: 0, remindersSent: 0, uiMessage: '❌ TRACKER sheet not found.' };
  }

  const data = trackerSheet.getDataRange().getValues();
  const now = new Date();
  const reminderDetails = [];
  let skippedHandled = 0;
  let skippedSameStage = 0;

  Logger.log(`--- runRenewalReminders started | Today: ${now.toDateString()} | Stages: 3→1st, 2→2nd, 1→Last ---`);

  // Statuses that mean the company is already decided — never remind these
  const DONE = new Set(['Renew', 'Renewed', 'Not Renewing', 'Terminated']);

  const dropdownRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.RENEWAL_STATUS_VALUES, true)
    .build();

  // Skip header rows (row 1 = header, row 2 = empty/subheader, data starts row 3)
  for (let i = 2; i < data.length; i++) {
    const contractEnd    = data[i][CONFIG.TRACKER_COLS.CONTRACT_END - 1];
    const companyName    = data[i][CONFIG.TRACKER_COLS.COMPANY_NAME - 1];
    const companyEmail   = data[i][CONFIG.TRACKER_COLS.COMPANY_EMAIL - 1];
    const pilotNumber    = data[i][CONFIG.TRACKER_COLS.PILOT_NUMBER - 1];
    const renewalStatus  = data[i][CONFIG.TRACKER_COLS.RENEWAL_STATUS - 1];
    const worqLocation   = data[i][CONFIG.TRACKER_COLS.WORQ_LOCATION - 1];

    if (!contractEnd || !companyEmail) {
      Logger.log(`Row ${i + 1} SKIPPED — ${companyName || '(no name)'} | Missing: ${!contractEnd ? 'contractEnd ' : ''}${!companyEmail ? 'companyEmail' : ''}`);
      continue;
    }

    // Already decided — leave alone
    if (DONE.has(renewalStatus)) {
      Logger.log(`Row ${i + 1} SKIPPED — ${companyName} | Already handled (${renewalStatus})`);
      skippedHandled++;
      continue;
    }

    const endDate = new Date(contractEnd);
    const monthsUntilExpiry = getMonthsDifference(now, endDate);

    // Determine the target stage for this month's months-left value
    const targetStage = CONFIG.REMINDER_STAGES[monthsUntilExpiry]; // undefined if not 3/2/1/0
    Logger.log(`Row ${i + 1} | ${companyName} | End: ${endDate.toDateString()} | Months left: ${monthsUntilExpiry} | Status: ${renewalStatus || '(empty)'} | Target stage: ${targetStage || '—'}`);

    // Only act at 3, 2, 1, 0 months out. monthsUntilExpiry is a whole-month
    // difference, so 0 means "expires this month" — still worth reminding even
    // if the exact end date falls earlier in the month. Anything already past
    // (negative months) has no target stage and is skipped.
    if (!targetStage) {
      continue;
    }

    // Never downgrade or repeat: only send if the target stage is strictly
    // stronger than the current stage.
    const currentRank = REMINDER_STAGE_ORDER.indexOf(renewalStatus); // -1 if empty/unknown
    const targetRank  = REMINDER_STAGE_ORDER.indexOf(targetStage);

    if (targetRank <= currentRank) {
      Logger.log(`Row ${i + 1} SKIPPED — ${companyName} | Already at "${renewalStatus}" (≥ target "${targetStage}")`);
      skippedSameStage++;
      continue;
    }

    sendRenewalReminderEmail(companyName, companyEmail, pilotNumber, endDate, monthsUntilExpiry, worqLocation);

    const statusCell = trackerSheet.getRange(i + 1, CONFIG.TRACKER_COLS.RENEWAL_STATUS);
    statusCell.setValue(targetStage);
    statusCell.setDataValidation(dropdownRule);

    reminderDetails.push({ name: companyName, monthsLeft: monthsUntilExpiry, stage: targetStage });
  }

  const remindersSent = reminderDetails.length;
  Logger.log(`Renewal reminders sent: ${remindersSent}`);

  const skipNote = (skippedHandled > 0 || skippedSameStage > 0)
    ? `\n\nSkipped: ${skippedHandled} already decided, ${skippedSameStage} already at this stage`
    : '';

  let uiMessage;
  if (remindersSent > 0) {
    const list = reminderDetails
      .map(d => `• ${d.name} — ${d.stage} (${d.monthsLeft} month${d.monthsLeft === 1 ? '' : 's'} left)`)
      .join('\n');
    uiMessage = `✅ ${remindersSent} renewal reminder${remindersSent === 1 ? '' : 's'} sent\n\n${list}${skipNote}`;
  } else {
    uiMessage = `ℹ️ No renewal reminders sent.\n\nNo awaiting-reply contracts hit a 3/2/1-month threshold this run.${skipNote}`;
  }

  return {
    sent: reminderDetails,
    skipped: skippedHandled + skippedSameStage,
    remindersSent: remindersSent,
    uiMessage: uiMessage
  };
}

/**
 * Email a summary of the monthly reminder run to REMINDER_SUMMARY_EMAIL so
 * unattended scheduled runs leave a visible trace (success confidence).
 */
function sendReminderRunSummary_(summary) {
  const tz = Session.getScriptTimeZone();
  const monthYear = Utilities.formatDate(new Date(), tz, 'MMMM yyyy');
  const subject = `Renewal Reminder Run — ${monthYear} (${summary.remindersSent} sent)`;

  const rows = summary.sent.length
    ? summary.sent.map(d =>
        `<tr><td style="padding:4px 10px;border:1px solid #ccc;">${toProperCase(d.name)}</td>` +
        `<td style="padding:4px 10px;border:1px solid #ccc;">${d.stage}</td>` +
        `<td style="padding:4px 10px;border:1px solid #ccc;text-align:center;">${d.monthsLeft}</td></tr>`
      ).join('')
    : `<tr><td colspan="3" style="padding:4px 10px;border:1px solid #ccc;">No reminders sent this run.</td></tr>`;

  const htmlBody = `
<p>Monthly renewal reminder run completed for <strong>${monthYear}</strong>.</p>
<p><strong>${summary.remindersSent}</strong> reminder(s) sent, ${summary.skipped} row(s) skipped.</p>
<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:13px;">
  <tr style="background:#f2f2f2;">
    <th style="padding:6px 10px;border:1px solid #ccc;text-align:left;">Company</th>
    <th style="padding:6px 10px;border:1px solid #ccc;text-align:left;">Stage</th>
    <th style="padding:6px 10px;border:1px solid #ccc;">Months Left</th>
  </tr>
  ${rows}
</table>
<p style="color:#888;font-size:12px;">Automated message from the AstriCloud Tracker.</p>`;

  try {
    MailApp.sendEmail({
      to: CONFIG.REMINDER_SUMMARY_EMAIL,
      subject: subject,
      htmlBody: htmlBody,
      name: 'AstriCloud Tracker'
    });
    Logger.log(`Reminder run summary emailed to ${CONFIG.REMINDER_SUMMARY_EMAIL}`);
  } catch (e) {
    Logger.log(`ERROR emailing reminder run summary: ${e.message}`);
  }
}

/**
 * ONE-TIME MIGRATION — realign reminder stages in col F to the current
 * months-to-expiry, WITHOUT sending any email.
 *
 * Use after changing the stage scheme (e.g. legacy 'Pending' rows, or rows set
 * under the old 3→1st/2→2nd/1→Last mapping that now needs 1→3rd/0→Last).
 * Only touches rows whose status is blank, 'Pending', or a reminder stage —
 * never rows already decided ('Renew'/'Renewed'/'Not Renewing'/'Terminated').
 * Shows a preview and asks for confirmation before writing.
 */
function migrateReminderStages() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);
  if (!trackerSheet) {
    ui.alert('❌ TRACKER sheet not found.');
    return;
  }

  const data = trackerSheet.getDataRange().getValues();
  const now = new Date();

  // Statuses safe to realign: blank, legacy 'Pending', or any reminder stage.
  const MIGRATABLE = new Set(['', 'Pending'].concat(REMINDER_STAGE_ORDER));

  const changes = []; // { rowNum, name, from, to }

  for (let i = 2; i < data.length; i++) {
    const companyName   = data[i][CONFIG.TRACKER_COLS.COMPANY_NAME - 1];
    const contractEnd   = data[i][CONFIG.TRACKER_COLS.CONTRACT_END - 1];
    const renewalStatus = (data[i][CONFIG.TRACKER_COLS.RENEWAL_STATUS - 1] || '').toString().trim();

    if (!companyName || !contractEnd) continue;
    if (!MIGRATABLE.has(renewalStatus)) continue;

    const monthsUntilExpiry = getMonthsDifference(now, new Date(contractEnd));
    const targetStage = CONFIG.REMINDER_STAGES[monthsUntilExpiry];

    // Outside the 3/2/1/0 window — nothing meaningful to set
    if (!targetStage) continue;
    if (targetStage === renewalStatus) continue;

    changes.push({ rowNum: i + 1, name: companyName, from: renewalStatus || '(blank)', to: targetStage });
  }

  if (changes.length === 0) {
    ui.alert('ℹ️ Nothing to migrate.\n\nAll reminder stages already match their months-to-expiry.');
    return;
  }

  const preview = changes
    .map(c => `• ${c.name}: ${c.from} → ${c.to}`)
    .join('\n');

  const response = ui.alert(
    'Migrate Reminder Stages — Preview',
    `${changes.length} row(s) will be updated. NO emails will be sent.\n\n${preview}\n\nProceed?`,
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) {
    ui.alert('Migration cancelled. No changes made.');
    return;
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.RENEWAL_STATUS_VALUES, true)
    .build();

  for (const c of changes) {
    const cell = trackerSheet.getRange(c.rowNum, CONFIG.TRACKER_COLS.RENEWAL_STATUS);
    cell.setValue(c.to);
    cell.setDataValidation(rule);
    Logger.log(`Migrated row ${c.rowNum} | ${c.name} | ${c.from} → ${c.to}`);
  }

  ui.alert(`✅ Migrated ${changes.length} row(s).\n\n${preview}\n\nNo emails were sent.`);
}

/**
 * Apply dropdown validation to the entire Renewal Status column (col F)
 * Run once to set up dropdowns on all existing rows
 */
function setupRenewalStatusDropdown() {
  const trackerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.TRACKER_SHEET);
  if (!trackerSheet) return;

  const lastRow = trackerSheet.getLastRow();
  if (lastRow < 3) return;

  const range = trackerSheet.getRange(3, CONFIG.TRACKER_COLS.RENEWAL_STATUS, lastRow - 2, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.RENEWAL_STATUS_VALUES, true)
    .build();
  range.setDataValidation(rule);

  Logger.log(`Renewal Status dropdown applied to rows 3–${lastRow}`);
  SpreadsheetApp.getUi().alert('✅ Renewal Status dropdown applied to all rows in col F.');
}

/**
 * Calculate months between two dates
 */
function getMonthsDifference(date1, date2) {
  const months = (date2.getFullYear() - date1.getFullYear()) * 12 + (date2.getMonth() - date1.getMonth());
  return Math.round(months);
}

/**
 * Send renewal reminder email
 */
function sendRenewalReminderEmail(companyName, email, pilotNumber, expiryDate, monthsLeft, worqLocation) {
  companyName = toProperCase(companyName);
  const expiryStr = Utilities.formatDate(expiryDate, Session.getScriptTimeZone(), 'dd MMM yyyy');
  pilotNumber = formatPilotNumber(pilotNumber);

  // monthsLeft === 0 means the contract expires THIS month — final notice with
  // distinct urgent wording. 3/2/1 use the standard countdown reminder.
  const isFinalNotice = monthsLeft === 0;

  const subject = isFinalNotice
    ? `⚠️ Virtual Landline Expiring This Month - Action Required`
    : `⏰ Virtual Landline Renewal Reminder - ${monthsLeft} Month${monthsLeft > 1 ? 's' : ''} Until Expiry`;

  const htmlBody = isFinalNotice ? `
<p>Dear ${companyName},</p>
<p><strong>Your virtual landline service expires this month.</strong> This is our final reminder before the service ends.</p>
<p>
  ${pilotNumber ? `<strong>📞 Pilot Number: ${pilotNumber}</strong><br>` : ''}
  <strong>📅 Expiry Date: ${expiryStr}</strong>
</p>
<p>To avoid interruption to your service, please confirm your renewal by replying to this email <strong>as soon as possible</strong>.</p>
<p>If we do not hear from you, your virtual landline will be deactivated after the expiry date above.</p>
<p>If you have any questions or need assistance, please reach out to us right away.</p>
<p>
  <strong>Best regards,</strong><br>
  <strong>WORQ Operations Team</strong>
</p>
` : `
<p>Dear ${companyName},</p>
<p>This is a friendly reminder that your virtual landline service is expiring soon.</p>
<p>
  ${pilotNumber ? `<strong>📞 Pilot Number: ${pilotNumber}</strong><br>` : ''}
  <strong>📅 Expiry Date: ${expiryStr}</strong><br>
  <strong>⏳ Time Remaining: ${monthsLeft} month${monthsLeft > 1 ? 's' : ''}</strong>
</p>
<p>To ensure uninterrupted service, please confirm your renewal by replying to this email at your earliest convenience.</p>
<p>If you have any questions, please don't hesitate to reach out.</p>
<p>
  <strong>Best regards,</strong><br>
  <strong>WORQ Operations Team</strong>
</p>
`;

  const locationEmail = getLocationEmail(worqLocation);

  try {
    const options = { to: email, subject: subject, htmlBody: htmlBody, name: 'WORQ Operations Team', replyTo: 'it@worq.space' };
    if (locationEmail) options.cc = locationEmail;
    MailApp.sendEmail(options);

    Logger.log(`Reminder sent to ${companyName} (${email})${locationEmail ? ` | CC: ${locationEmail}` : ''}`);
  } catch (e) {
    Logger.log(`ERROR sending email to ${email}: ${e.message}`);
  }
}

/**
 * Send renewal confirmation (thank you) email with new contract tenure
 */
function sendRenewalConfirmationEmail(companyName, email, pilotNumber, newStartDate, newEndDate, worqLocation) {
  companyName = toProperCase(companyName);
  const newStartStr = Utilities.formatDate(newStartDate, Session.getScriptTimeZone(), 'dd MMM yyyy');
  const newEndStr   = Utilities.formatDate(newEndDate,   Session.getScriptTimeZone(), 'dd MMM yyyy');
  pilotNumber = formatPilotNumber(pilotNumber);

  const subject = `✅ Virtual Landline Renewal Confirmed - Thank You, ${companyName}!`;

  const htmlBody = `
<p>Dear ${companyName},</p>
<p>Thank you for renewing your virtual landline service with WORQ! We truly appreciate your continued support.</p>
<p>Your service has been successfully renewed with the following details:</p>
<p>
  ${pilotNumber ? `<strong>📞 Pilot Number: ${pilotNumber}</strong><br>` : ''}
  <strong>📅 New Service Period: ${newStartStr} – ${newEndStr}</strong>
</p>
<p>Your virtual landline will remain active without interruption throughout this period.</p>
<p>If you have any questions or require any assistance, please don't hesitate to reach out to us.</p>
<p>Thank you once again for choosing WORQ. We look forward to serve you.</p>
<p>
  <strong>Best regards,</strong><br>
  <strong>WORQ Operations Team</strong>
</p>
`;

  const locationEmail = getLocationEmail(worqLocation);

  try {
    const options = { to: email, subject: subject, htmlBody: htmlBody, name: 'WORQ Operations Team', replyTo: 'it@worq.space' };
    if (locationEmail) options.cc = locationEmail;
    MailApp.sendEmail(options);

    Logger.log(`Renewal confirmation sent to ${companyName} (${email})${locationEmail ? ` | CC: ${locationEmail}` : ''}`);
  } catch (e) {
    Logger.log(`ERROR sending confirmation to ${email}: ${e.message}`);
  }
}

/**
 * Send termination confirmation email to customer who chose not to renew
 */
function sendTerminationEmail(companyName, email, pilotNumber, contractEndDate, worqLocation) {
  companyName = toProperCase(companyName);
  const endStr = Utilities.formatDate(contractEndDate, Session.getScriptTimeZone(), 'dd MMM yyyy');
  pilotNumber = formatPilotNumber(pilotNumber);

  const subject = `Virtual Landline Service Termination Confirmation - ${companyName}`;

  const htmlBody = `
<p>Dear ${companyName},</p>
<p>We have received and acknowledged your decision not to renew your virtual landline service with WORQ.</p>
<p>This email serves as confirmation of your termination request.</p>
<p>
  ${pilotNumber ? `<strong>📞 Pilot Number: ${pilotNumber}</strong><br>` : ''}
  <strong>📅 Service End Date: ${endStr}</strong>
</p>
<p>Your virtual landline service will remain active until the above date, after which it will be deactivated.</p>
<p>We're sorry to see you go and hope to have the opportunity to serve you again in the future. If you change your mind or have any questions before your service ends, please do not hesitate to reach out.</p>
<p>Thank you for being a valued WORQ customer.</p>
<p>
  <strong>Best regards,</strong><br>
  <strong>WORQ Operations Team</strong>
</p>
`;

  const locationEmail = getLocationEmail(worqLocation);

  try {
    const options = { to: email, subject: subject, htmlBody: htmlBody, name: 'WORQ Operations Team', replyTo: 'it@worq.space' };
    if (locationEmail) options.cc = locationEmail;
    MailApp.sendEmail(options);

    Logger.log(`Termination confirmation sent to ${companyName} (${email})${locationEmail ? ` | CC: ${locationEmail}` : ''}`);
  } catch (e) {
    Logger.log(`ERROR sending termination email to ${email}: ${e.message}`);
  }
}

/**
 * Ensure pilot number has a leading zero (Google Sheets drops it when stored as a number).
 */
function formatPilotNumber(pilotNumber) {
  if (!pilotNumber) return null;
  const str = String(pilotNumber).trim();
  return str.startsWith('0') ? str : '0' + str;
}

/**
 * Convert a string to Proper Case (Title Case).
 * e.g. "INCOMPLETENESS THEOREM SDN BHD" → "Incompleteness Theorem Sdn Bhd"
 */
function toProperCase(str) {
  if (!str) return str;
  return String(str).trim().toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
}

/**
 * Send a single vendor notification email listing all renewals and terminations
 * processed in this sync run.
 * @param {Array} renewals    - items from toRenew  (needs: companyName, newEndDate)
 * @param {Array} terminations - items from toTerminate (needs: companyName, currentEndDate)
 */
function sendVendorNotificationEmail(renewals, terminations) {
  const tz = Session.getScriptTimeZone();
  const monthYear = Utilities.formatDate(new Date(), tz, 'MMMM yyyy');
  const subject = `Astricloud Landline Subscription Update - ${monthYear}`;

  // Build table rows — terminations first, then renewals (matches screenshot order)
  let tableRows = '';

  for (const r of terminations) {
    const endStr = Utilities.formatDate(r.currentEndDate, tz, 'dd-MMM-yyyy');
    const name   = toProperCase(r.companyName);
    tableRows += `
      <tr>
        <td style="border:1px solid #000;padding:6px 12px;">${name}</td>
        <td style="border:1px solid #000;padding:6px 12px;text-align:center;">${endStr}</td>
        <td style="border:1px solid #000;padding:6px 12px;text-align:center;color:#cc0000;font-weight:bold;">Terminate</td>
        <td style="border:1px solid #000;padding:6px 12px;">Remove from WORQ Billing</td>
      </tr>`;
  }

  for (const r of renewals) {
    const endStr = Utilities.formatDate(r.newEndDate, tz, 'dd-MMM-yyyy');
    const name   = toProperCase(r.companyName);
    tableRows += `
      <tr>
        <td style="border:1px solid #000;padding:6px 12px;">${name}</td>
        <td style="border:1px solid #000;padding:6px 12px;text-align:center;">${endStr}</td>
        <td style="border:1px solid #000;padding:6px 12px;text-align:center;color:#007700;font-weight:bold;">Renew</td>
        <td style="border:1px solid #000;padding:6px 12px;">Continue on WORQ Billing</td>
      </tr>`;
  }

  const htmlBody = `
<p>Hi Theinosha,</p>
<p>Please assist to process the below renewal / termination decision by the customer.</p>
<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:13px;">
  <tr style="background:#f2f2f2;">
    <th style="border:1px solid #000;padding:8px 14px;text-align:left;">Companies</th>
    <th style="border:1px solid #000;padding:8px 14px;">End</th>
    <th style="border:1px solid #000;padding:8px 14px;">Final Status</th>
    <th style="border:1px solid #000;padding:8px 14px;text-align:left;">AstriCloud Action</th>
  </tr>
  ${tableRows}
</table>
<br>
<p>
  <strong>Best regards,</strong><br>
  <strong>WORQ Operations Team</strong>
</p>`;

  const vendorEmail = getVendorEmail();
  const vendorCc    = getVendorCc();

  try {
    MailApp.sendEmail({
      to:       vendorEmail,
      cc:       vendorCc,
      subject:  subject,
      htmlBody: htmlBody,
      name:     'WORQ Operations Team',
      replyTo:  'it@worq.space'
    });
    Logger.log(`Vendor notification sent to ${vendorEmail} | CC: ${vendorCc}`);
  } catch (e) {
    Logger.log(`ERROR sending vendor notification: ${e.message}`);
  }
}

/**
 * Read the vendor recipient ("To") address from the Config sheet at runtime.
 * Falls back to CONFIG.VENDOR_EMAIL if the sheet or key is missing, so email
 * never breaks even before the Config sheet is set up.
 * @returns {string}
 */
function getVendorEmail() {
  return getConfigValue('VENDOR_EMAIL') || CONFIG.VENDOR_EMAIL;
}

/**
 * Read the vendor CC list from the Config sheet at runtime.
 * Falls back to CONFIG.VENDOR_CC if the sheet or key is missing.
 * @returns {string} comma-separated CC addresses
 */
function getVendorCc() {
  return getConfigValue('VENDOR_CC') || CONFIG.VENDOR_CC;
}

/**
 * Generic lookup for a value in the Config sheet.
 * Config sheet layout: col A = Key, col B = Value (row 1 = header).
 * Returns the trimmed value string, or null if the sheet/key is missing or blank.
 * Mirrors the getLocationEmail() lookup pattern.
 * @param {string} key
 * @returns {string|null}
 */
function getConfigValue(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG.CONFIG_SHEET);
  if (!configSheet) return null;

  const data = configSheet.getDataRange().getValues();
  const keyStr = key.toString().trim().toLowerCase();

  // Row 1 is the header — start from row index 1
  for (let i = 1; i < data.length; i++) {
    const rowKey = data[i][0]; // Column A
    if (rowKey && rowKey.toString().trim().toLowerCase() === keyStr) {
      const value = data[i][1]; // Column B
      const valueStr = value ? value.toString().trim() : '';
      return valueStr || null;
    }
  }

  return null;
}

/**
 * Create the Config sheet (if missing) and seed it with the current
 * hardcoded VENDOR_EMAIL / VENDOR_CC values from CONFIG. Non-destructive:
 * existing keys are left untouched so admin edits are never overwritten.
 */
function setupConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName(CONFIG.CONFIG_SHEET);
  let created = false;

  if (!configSheet) {
    configSheet = ss.insertSheet(CONFIG.CONFIG_SHEET);
    configSheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]).setFontWeight('bold');
    configSheet.setFrozenRows(1);
    configSheet.setColumnWidth(1, 160);
    configSheet.setColumnWidth(2, 600);
    created = true;
  }

  // Collect existing keys so we never overwrite admin-edited values
  const data = configSheet.getDataRange().getValues();
  const existingKeys = {};
  for (let i = 1; i < data.length; i++) {
    const k = data[i][0];
    if (k) existingKeys[k.toString().trim().toLowerCase()] = true;
  }

  const seed = [
    ['VENDOR_EMAIL', CONFIG.VENDOR_EMAIL],
    ['VENDOR_CC',    CONFIG.VENDOR_CC]
  ];
  const toAppend = seed.filter(row => !existingKeys[row[0].toLowerCase()]);

  if (toAppend.length > 0) {
    configSheet.getRange(configSheet.getLastRow() + 1, 1, toAppend.length, 2).setValues(toAppend);
  }

  SpreadsheetApp.getUi().alert(
    `✅ Config sheet ready${created ? ' (created)' : ''}.\n\n` +
    (toAppend.length > 0
      ? `Seeded ${toAppend.length} key(s): ${toAppend.map(r => r[0]).join(', ')}\n\n`
      : `All keys already present — nothing overwritten.\n\n`) +
    `Edit vendor recipients directly in the "${CONFIG.CONFIG_SHEET}" sheet (col B). ` +
    `Changes apply immediately — no code deploy needed.`
  );
}

/**
 * Look up the location email from the Addresses sheet by matching the site name (col A).
 * Returns the email string (col C) or null if not found.
 */
function getLocationEmail(worqLocation) {
  if (!worqLocation) return null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const addressSheet = ss.getSheetByName('Addresses');
  if (!addressSheet) {
    Logger.log('getLocationEmail: Addresses sheet not found');
    return null;
  }

  const data = addressSheet.getDataRange().getValues();
  const locationStr = worqLocation.toString().trim().toLowerCase();

  // Row 1 is the header — start from row index 1
  for (let i = 1; i < data.length; i++) {
    const siteName = data[i][0]; // Column A
    if (siteName && siteName.toString().trim().toLowerCase() === locationStr) {
      const locationEmail = data[i][2]; // Column C
      return locationEmail ? locationEmail.toString().trim() : null;
    }
  }

  Logger.log(`getLocationEmail: No match found for location "${worqLocation}"`);
  return null;
}
