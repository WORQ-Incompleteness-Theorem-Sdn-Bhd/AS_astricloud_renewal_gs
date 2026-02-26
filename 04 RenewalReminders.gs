/**
 * Check for upcoming contract expirations and send reminder emails
 */
function checkAndSendReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);

  if (!trackerSheet) {
    Logger.log('ERROR: TRACKER sheet not found');
    return;
  }

  const data = trackerSheet.getDataRange().getValues();
  const now = new Date();
  let remindersSent = 0;
  const reminderDetails = [];
  let skippedPending = 0;
  let skippedRenew = 0;

  Logger.log(`--- checkAndSendReminders started | Today: ${now.toDateString()} | REMINDER_MONTHS: [${CONFIG.REMINDER_MONTHS}] ---`);

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

    const endDate = new Date(contractEnd);
    const monthsUntilExpiry = getMonthsDifference(now, endDate);

    Logger.log(`Row ${i + 1} | ${companyName} | End: ${endDate.toDateString()} | Months until expiry: ${monthsUntilExpiry} | Renewal Status: ${renewalStatus || '(empty)'}`);

    // Check if we should send reminder (3, 2, 1, or 0 months before — 0 = expiring this month)
    if (CONFIG.REMINDER_MONTHS.includes(monthsUntilExpiry) && endDate >= now) {

      if (renewalStatus === 'Renew') {
        Logger.log(`Row ${i + 1} SKIPPED — ${companyName} | Already marked Renew`);
        skippedRenew++;
        continue;
      }

      if (renewalStatus === 'Pending') {
        Logger.log(`Row ${i + 1} SKIPPED — ${companyName} | Reminder already sent (Pending)`);
        skippedPending++;
        continue;
      }

      sendRenewalReminderEmail(companyName, companyEmail, pilotNumber, endDate, monthsUntilExpiry, worqLocation);

      // Mark as Pending in TRACKER col F with dropdown
      const pendingCell = trackerSheet.getRange(i + 1, CONFIG.TRACKER_COLS.RENEWAL_STATUS);
      pendingCell.setValue('Pending');
      pendingCell.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(['Pending', 'Renew', 'Renewed', 'Not Renewing'], true)
          .build()
      );

      reminderDetails.push({ name: companyName, monthsLeft: monthsUntilExpiry });
      remindersSent++;
    } else {
      Logger.log(`Row ${i + 1} SKIPPED — ${companyName} | monthsUntilExpiry (${monthsUntilExpiry}) not in REMINDER_MONTHS or endDate already passed`);
    }
  }

  Logger.log(`Renewal reminders sent: ${remindersSent}`);

  let message;
  if (remindersSent > 0) {
    const list = reminderDetails
      .map(d => `• ${d.name} — ${d.monthsLeft} month${d.monthsLeft === 1 ? '' : 's'} left`)
      .join('\n');
    const skipNote = (skippedPending > 0 || skippedRenew > 0)
      ? `\n\nSkipped: ${skippedPending} already Pending, ${skippedRenew} already Renew`
      : '';
    message = `✅ ${remindersSent} renewal reminder${remindersSent === 1 ? '' : 's'} sent\n\n${list}${skipNote}`;
  } else {
    const skipNote = (skippedPending > 0 || skippedRenew > 0)
      ? `\n\nSkipped: ${skippedPending} already Pending, ${skippedRenew} already Renew`
      : '';
    message = `ℹ️ No renewal reminders sent.\n\nNo contracts are expiring in the next ${Math.max(...CONFIG.REMINDER_MONTHS)} months.${skipNote}`;
  }
  SpreadsheetApp.getUi().alert(message);
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
    .requireValueInList(['Pending', 'Renew', 'Renewed', 'Not Renewing'], true)
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

  const subject = `⏰ Virtual Landline Renewal Reminder - ${monthsLeft} Month${monthsLeft > 1 ? 's' : ''} Until Expiry`;

  const htmlBody = `
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

  try {
    MailApp.sendEmail({
      to:       CONFIG.VENDOR_EMAIL,
      cc:       CONFIG.VENDOR_CC,
      subject:  subject,
      htmlBody: htmlBody,
      name:     'WORQ Operations Team',
      replyTo:  'it@worq.space'
    });
    Logger.log(`Vendor notification sent to ${CONFIG.VENDOR_EMAIL} | CC: ${CONFIG.VENDOR_CC}`);
  } catch (e) {
    Logger.log(`ERROR sending vendor notification: ${e.message}`);
  }
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
