/**
 * Handles new form submissions in Form Responses 1.
 * Installed as an onFormSubmit trigger via setupFormSubmitTrigger().
 *
 * On each submission:
 *  1. Clears old yellow highlight from previous submission
 *  2. Highlights the new submission row in yellow
 *  3. Checks if the email is a duplicate (already exists in a prior row)
 *     - Duplicate → move to Archived Form Responses, no vendor email
 *     - New       → send vendor notification to AstriCloud
 */
function onFormSubmit(e) {
  // Thin wrapper: run the handler under error alerting so a failed submission
  // (which happens unattended) surfaces instead of failing silently.
  try {
    handleFormSubmit_(e);
  } catch (err) {
    notifyError_('onFormSubmit', err);
    throw err; // still surface in the Apps Script execution log
  }
}

function handleFormSubmit_(e) {
  if (!e || !e.range) {
    Logger.log('onFormSubmit: no event object — must be run via trigger, not manually');
    return;
  }

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName(CONFIG.FORM_RESPONSES_SHEET);
  if (!formSheet) {
    Logger.log('onFormSubmit: Form Responses 1 sheet not found');
    return;
  }

  const rowNum  = e.range.getRow();
  const rowData = formSheet.getRange(rowNum, 1, 1, formSheet.getLastColumn()).getValues()[0];

  const timestamp    = rowData[CONFIG.FORM_COLS.TIMESTAMP    - 1];
  const email        = (rowData[CONFIG.FORM_COLS.EMAIL        - 1] || '').toString().trim();
  const companyName  = (rowData[CONFIG.FORM_COLS.COMPANY_NAME - 1] || '').toString().trim();
  const worqLocation = (rowData[CONFIG.FORM_COLS.WORQ_LOCATION - 1] || '').toString().trim();

  // 1. Clear old yellow highlight, then highlight new row
  clearFormHighlights(formSheet);
  formSheet.getRange(rowNum, 1, 1, formSheet.getLastColumn()).setBackground('#FFFF00');

  Logger.log(`onFormSubmit: row ${rowNum} | ${companyName} | ${email} | ${worqLocation}`);

  // 2. Duplicate check
  if (isEmailDuplicate(formSheet, email, rowNum)) {
    Logger.log(`onFormSubmit: DUPLICATE email "${email}" — moving to Archived Form Responses`);
    moveToArchivedFormResponses(formSheet, rowNum, rowData);
    return;
  }

  // 3. New submission — notify vendor
  sendNewSignupVendorEmail(companyName, email, worqLocation, timestamp);
}

/**
 * Clears yellow (#FFFF00) background from all data rows in Form Responses 1.
 * Other cell colors are left untouched.
 */
function clearFormHighlights(formSheet) {
  const lastRow = formSheet.getLastRow();
  const lastCol = formSheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  const range       = formSheet.getRange(2, 1, lastRow - 1, lastCol);
  const backgrounds = range.getBackgrounds();

  for (let r = 0; r < backgrounds.length; r++) {
    for (let c = 0; c < backgrounds[r].length; c++) {
      if (backgrounds[r][c].toUpperCase() === '#FFFF00') {
        formSheet.getRange(r + 2, c + 1).setBackground(null);
      }
    }
  }
}

/**
 * Returns true if `email` already exists in a prior row of Form Responses 1
 * (any row except the header row 1 and the current submission row).
 */
function isEmailDuplicate(formSheet, email, currentRowNum) {
  if (!email) return false;

  const emailNorm = email.toLowerCase();
  const data      = formSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) { // i=1 skips header row (row 1)
    const rowNum = i + 1; // 1-based row number
    if (rowNum === currentRowNum) continue;

    const existingEmail = (data[i][CONFIG.FORM_COLS.EMAIL - 1] || '').toString().trim().toLowerCase();
    if (existingEmail === emailNorm) return true;
  }

  return false;
}

/**
 * Copies rowData to Archived Form Responses sheet, then deletes the row from Form Responses 1.
 * Google Forms always appends to the last row of the sheet, so row deletion is safe.
 */
function moveToArchivedFormResponses(formSheet, rowNum, rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get or create the archive sheet
  let archiveSheet = ss.getSheetByName(CONFIG.ARCHIVED_FORM_RESPONSES_SHEET);
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet(CONFIG.ARCHIVED_FORM_RESPONSES_SHEET);
    Logger.log(`moveToArchivedFormResponses: created sheet "${CONFIG.ARCHIVED_FORM_RESPONSES_SHEET}"`);
  }

  // Append to archive — direct copy (same columns as Form Responses 1)
  archiveSheet.appendRow(rowData);

  // Remove from Form Responses 1 to keep it clean
  formSheet.deleteRow(rowNum);

  Logger.log(`moveToArchivedFormResponses: moved row ${rowNum} to "${CONFIG.ARCHIVED_FORM_RESPONSES_SHEET}"`);
}

/**
 * Sends a new signup notification to the AstriCloud vendor.
 * CC: full VENDOR_CC list + outlet email (from Addresses sheet, if found).
 */
function sendNewSignupVendorEmail(companyName, email, worqLocation, timestamp) {
  const tz          = Session.getScriptTimeZone();
  const displayName = toProperCase(companyName) || companyName;
  const timestampStr = timestamp
    ? Utilities.formatDate(new Date(timestamp), tz, 'dd MMM yyyy, hh:mm a')
    : 'N/A';

  const subject = `New Virtual Landline Signup Request - ${displayName}`;

  const htmlBody = `
<p>Hi Theinosha,</p>
<p>A new company has submitted a Virtual Landline signup request through WORQ.</p>
<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:13px;">
  <tr style="background:#f2f2f2;">
    <th style="border:1px solid #000;padding:8px 14px;text-align:left;">Field</th>
    <th style="border:1px solid #000;padding:8px 14px;text-align:left;">Details</th>
  </tr>
  <tr>
    <td style="border:1px solid #000;padding:6px 12px;"><strong>Company Name</strong></td>
    <td style="border:1px solid #000;padding:6px 12px;">${displayName}</td>
  </tr>
  <tr>
    <td style="border:1px solid #000;padding:6px 12px;"><strong>Email</strong></td>
    <td style="border:1px solid #000;padding:6px 12px;">${email}</td>
  </tr>
  <tr>
    <td style="border:1px solid #000;padding:6px 12px;"><strong>WORQ Location</strong></td>
    <td style="border:1px solid #000;padding:6px 12px;">${worqLocation || 'N/A'}</td>
  </tr>
  <tr>
    <td style="border:1px solid #000;padding:6px 12px;"><strong>Submitted At</strong></td>
    <td style="border:1px solid #000;padding:6px 12px;">${timestampStr}</td>
  </tr>
</table>
<br>
<p>You may view the full signup list in the tracker: <a href="https://docs.google.com/spreadsheets/d/1t_-C-TZjd7dN6uweYG3wdZkToVAG4pxfEKInIG-TolE/edit?gid=219063083#gid=219063083">Form Responses 1 – AstriCloud Tracker</a></p>
<p>Please assist to provision a virtual landline pilot number for this company.</p>
<p>
  <strong>Best regards,</strong><br>
  <strong>WORQ Operations Team</strong>
</p>`;

  // Build CC: VENDOR_CC (from Config sheet, fallback to CONFIG) + outlet email (if found)
  const vendorCc = getVendorCc();
  const locationEmail = getLocationEmail(worqLocation);
  const cc = locationEmail
    ? vendorCc + ',' + locationEmail
    : vendorCc;

  try {
    MailApp.sendEmail({
      to:       getVendorEmail(),
      cc:       cc,
      subject:  subject,
      htmlBody: htmlBody,
      name:     'WORQ Operations Team',
      replyTo:  'it@worq.space'
    });
    Logger.log(`sendNewSignupVendorEmail: sent for "${displayName}" | CC: ${cc}`);
  } catch (err) {
    Logger.log(`sendNewSignupVendorEmail ERROR: ${err.message}`);
  }
}

/**
 * Creates an installable onFormSubmit trigger pointing to onFormSubmit().
 * Removes any existing trigger with the same handler first to prevent duplicates.
 */
function setupFormSubmitTrigger() {
  const triggers = ScriptApp.getProjectTriggers();

  let removed = 0;
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });

  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();

  SpreadsheetApp.getUi().alert(
    `✅ Form submit trigger set\n\n` +
    `onFormSubmit() will run automatically whenever a new form response is received.` +
    (removed > 0 ? `\n\n(${removed} previous trigger(s) replaced)` : '')
  );
}

/**
 * Removes all installable onFormSubmit triggers pointing to onFormSubmit().
 */
function removeFormSubmitTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;

  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });

  SpreadsheetApp.getUi().alert(
    removed > 0
      ? `✅ Form submit trigger removed.\n\nonFormSubmit() will no longer run automatically.`
      : `ℹ️ No form submit trigger found.\n\nNothing to remove.`
  );
}