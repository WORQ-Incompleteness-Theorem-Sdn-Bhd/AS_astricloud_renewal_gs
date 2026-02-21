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

  Logger.log(`--- checkAndSendReminders started | Today: ${now.toDateString()} | REMINDER_MONTHS: [${CONFIG.REMINDER_MONTHS}] ---`);

  // Skip header rows (row 1 = header, row 2 = empty/subheader, data starts row 3)
  for (let i = 2; i < data.length; i++) {
    const contractEnd    = data[i][CONFIG.TRACKER_COLS.CONTRACT_END - 1];
    const companyName    = data[i][CONFIG.TRACKER_COLS.COMPANY_NAME - 1];
    const companyEmail   = data[i][CONFIG.TRACKER_COLS.COMPANY_EMAIL - 1];
    const pilotNumber    = data[i][CONFIG.TRACKER_COLS.PILOT_NUMBER - 1];
    const renewalStatus  = data[i][CONFIG.TRACKER_COLS.RENEWAL_STATUS - 1];

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
        continue;
      }

      if (renewalStatus === 'Pending') {
        Logger.log(`Row ${i + 1} SKIPPED — ${companyName} | Reminder already sent (Pending)`);
        continue;
      }

      sendRenewalReminderEmail(companyName, companyEmail, pilotNumber, endDate, monthsUntilExpiry);

      // Mark as Pending in TRACKER col F with dropdown
      const pendingCell = trackerSheet.getRange(i + 1, CONFIG.TRACKER_COLS.RENEWAL_STATUS);
      pendingCell.setValue('Pending');
      pendingCell.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(['Pending', 'Renew', 'Renewed', 'Not Renewing'], true)
          .build()
      );

      remindersSent++;
    } else {
      Logger.log(`Row ${i + 1} SKIPPED — ${companyName} | monthsUntilExpiry (${monthsUntilExpiry}) not in REMINDER_MONTHS or endDate already passed`);
    }
  }

  Logger.log(`Renewal reminders sent: ${remindersSent}`);
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
function sendRenewalReminderEmail(companyName, email, pilotNumber, expiryDate, monthsLeft) {
  const expiryStr = Utilities.formatDate(expiryDate, Session.getScriptTimeZone(), 'dd MMM yyyy');

  const subject = `⏰ Virtual Landline Renewal Reminder - ${monthsLeft} Month${monthsLeft > 1 ? 's' : ''} Until Expiry`;

  const body = `
Dear ${companyName},

This is a friendly reminder that your virtual landline service is expiring soon.

${pilotNumber ? `📞 Pilot Number: ${pilotNumber}\n` : ''}📅 Expiry Date: ${expiryStr}
⏳ Time Remaining: ${monthsLeft} month${monthsLeft > 1 ? 's' : ''}

To ensure uninterrupted service, please confirm your renewal at your earliest convenience.

If you have any questions or would like to discuss renewal options, please don't hesitate to reach out.

Best regards,
WORQ IT Operations Team
${CONFIG.EMAIL_FROM}
`;

  try {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body,
      name: 'WORQ IT Operations'
    });

    Logger.log(`Reminder sent to ${companyName} (${email})`);
  } catch (e) {
    Logger.log(`ERROR sending email to ${email}: ${e.message}`);
  }
}

/**
 * Send renewal confirmation (thank you) email with new contract tenure
 */
function sendRenewalConfirmationEmail(companyName, email, pilotNumber, newStartDate, newEndDate) {
  const newStartStr = Utilities.formatDate(newStartDate, Session.getScriptTimeZone(), 'dd MMM yyyy');
  const newEndStr   = Utilities.formatDate(newEndDate,   Session.getScriptTimeZone(), 'dd MMM yyyy');

  const subject = `✅ Virtual Landline Renewal Confirmed - Thank You, ${companyName}!`;

  const body = `
Dear ${companyName},

Thank you for renewing your virtual landline service with WORQ! We truly appreciate your continued support.

Your service has been successfully renewed with the following details:

${pilotNumber ? `📞 Pilot Number: ${pilotNumber}\n` : ''}📅 New Service Period: ${newStartStr} – ${newEndStr}

Your virtual landline will remain active without interruption throughout this period.

If you have any questions or require any assistance, please don't hesitate to reach out to us.

Thank you once again for choosing WORQ. We look forward to continuing to serve you.

Best regards,
WORQ IT Operations Team
${CONFIG.EMAIL_FROM}
`;

  try {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body,
      name: 'WORQ IT Operations'
    });

    Logger.log(`Renewal confirmation sent to ${companyName} (${email})`);
  } catch (e) {
    Logger.log(`ERROR sending confirmation to ${email}: ${e.message}`);
  }
}

/**
 * Send termination confirmation email to customer who chose not to renew
 */
function sendTerminationEmail(companyName, email, pilotNumber, contractEndDate) {
  const endStr = Utilities.formatDate(contractEndDate, Session.getScriptTimeZone(), 'dd MMM yyyy');

  const subject = `Virtual Landline Service Termination Confirmation - ${companyName}`;

  const body = `
Dear ${companyName},

We have received and acknowledged your decision not to renew your virtual landline service with WORQ.

This email serves as confirmation of your termination request.

${pilotNumber ? `📞 Pilot Number: ${pilotNumber}\n` : ''}📅 Service End Date: ${endStr}

Your virtual landline service will remain active until the above date, after which it will be deactivated.

We're sorry to see you go and hope to have the opportunity to serve you again in the future. If you change your mind or have any questions before your service ends, please do not hesitate to reach out.

Thank you for being a valued WORQ customer.

Best regards,
WORQ IT Operations Team
${CONFIG.EMAIL_FROM}
`;

  try {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body,
      name: 'WORQ IT Operations'
    });

    Logger.log(`Termination confirmation sent to ${companyName} (${email})`);
  } catch (e) {
    Logger.log(`ERROR sending termination email to ${email}: ${e.message}`);
  }
}
