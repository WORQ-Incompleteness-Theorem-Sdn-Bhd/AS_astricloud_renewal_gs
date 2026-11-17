/**
 * Client Termination Request form.
 *
 * A second Google Form (linked to this same spreadsheet) where a client can
 * tell us they don't want to renew. The form only asks for the company name
 * (reference only), pilot number, registered company email and a reason —
 * everything else is already in TRACKER. Responses land in the "Termination Requests" tab. The existing
 * onFormSubmit trigger (10 FormSubmitHandler) routes those submissions here —
 * no extra trigger is needed.
 *
 * On each termination submission:
 *  1. Find the TRACKER row by pilot number, then check the submitted email
 *     matches that row's Company Email (col D). Pilot numbers are public phone
 *     numbers, so the email is what proves the request is genuine.
 *  2. Both match one row → set Renewal Status (col F) to "Not Renewing" and add
 *     a note on the cell. Anything else → change nothing, flag NEEDS REVIEW.
 *  3. Write the outcome next to the response in "Termination Requests".
 *  4. Email Ops (Config: TERMINATION_NOTIFY_EMAIL, default it@worq.space), CC the
 *     company's outlet email from the Addresses sheet, so they can review.
 *
 * Nothing is sent to AstriCloud or the client here. Ops still runs
 * [04] Sync Renewals as usual, which sends the client termination
 * confirmation and the vendor email. Terminations always take effect at the
 * end of the current contract (clients pay the full term up front), so
 * early-termination requests are handled the same way.
 *
 * One-time setup: run createTerminationForm() from the Apps Script editor.
 */

var TERMINATION = {
  SHEET: 'Termination Requests',
  NOTIFY_CONFIG_KEY: 'TERMINATION_NOTIFY_EMAIL',
  FORM_ID_PROPERTY: 'TERMINATION_FORM_ID',

  // Form question titles — the handler reads answers by these titles.
  // If you rename a question in the form, update the title here too.
  Q: {
    COMPANY_NAME:  'Company Name',     // for reference only — matching uses pilot + email
    PILOT_NUMBER:  'Pilot Number (Virtual Landline Number)',
    COMPANY_EMAIL: 'Registered Company Email',
    REASON:        'Reason for Termination',
    ACK:           'Acknowledgement'
  },

  // Status columns the script writes to the right of the form answers
  STATUS_HEADERS: ['Processing Status', 'Matched TRACKER Company', 'Contract End Date', 'Processed At']
};

// ─────────────────────────────────────────────────────────────────────────────
// Submission handler (called from onFormSubmit in 10 FormSubmitHandler)
// ─────────────────────────────────────────────────────────────────────────────

function handleTerminationSubmit_(e) {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet     = e.range.getSheet();
  const rowNum       = e.range.getRow();
  const trackerSheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);
  const tz           = Session.getScriptTimeZone();

  const answer = title => {
    const v = e.namedValues && e.namedValues[title];
    return (v && v.length ? v[0] : '').toString().trim();
  };

  const req = {
    companyName:  answer(TERMINATION.Q.COMPANY_NAME),
    pilotNumber:  answer(TERMINATION.Q.PILOT_NUMBER),
    companyEmail: answer(TERMINATION.Q.COMPANY_EMAIL),
    reason:       answer(TERMINATION.Q.REASON),
    submittedAt:  new Date()
  };

  Logger.log(`handleTerminationSubmit_: row ${rowNum} | ${req.companyName} | pilot ${req.pilotNumber} | ${req.companyEmail}`);

  if (!trackerSheet) throw new Error('TRACKER sheet not found');

  const match = matchTerminationRequest_(trackerSheet.getDataRange().getValues(), req);
  let status = match.status;
  let previousStatus = '';

  if (match.ok) {
    const statusCell = trackerSheet.getRange(match.rowNum, CONFIG.TRACKER_COLS.RENEWAL_STATUS);
    previousStatus = (statusCell.getValue() || '').toString().trim();

    if (previousStatus === 'Terminated' || previousStatus === 'Not Renewing') {
      status = `Already "${previousStatus}" — no change`;
    } else if (previousStatus === 'Renew' || previousStatus === 'Renewed') {
      // The client already chose (or paid for) a renewal — never undo that
      // automatically; a person must confirm which decision stands.
      status = `NEEDS REVIEW — contract is already "${previousStatus}", nothing changed`;
    } else {
      statusCell.setValue('Not Renewing');
      const stamp = Utilities.formatDate(req.submittedAt, tz, 'dd MMM yyyy HH:mm');
      const oldNote = statusCell.getNote();
      statusCell.setNote(
        (oldNote ? oldNote + '\n\n' : '') +
        `Termination form submitted ${stamp} by "${req.companyName}" (${req.companyEmail}).` +
        (previousStatus ? ` Previous status: ${previousStatus}.` : '') +
        (req.reason ? `\nReason: ${req.reason}` : '')
      );
      status = 'Set to "Not Renewing"';
    }
  } else if (!match.row && isInArchived_(ss, req)) {
    status = 'NEEDS REVIEW — pilot number is already in ARCHIVED, nothing changed';
  }

  const company = match.row ? {
    name:        match.row[CONFIG.TRACKER_COLS.COMPANY_NAME  - 1],
    location:    match.row[CONFIG.TRACKER_COLS.WORQ_LOCATION - 1],
    email:       match.row[CONFIG.TRACKER_COLS.COMPANY_EMAIL - 1],
    contractEnd: match.row[CONFIG.TRACKER_COLS.CONTRACT_END  - 1] || null
  } : null;

  writeTerminationStatus_(reqSheet, rowNum, [
    status,
    company ? company.name : '',
    company && company.contractEnd ? new Date(company.contractEnd) : '',
    req.submittedAt
  ]);

  // TRACKER and the status columns are already written, so don't let an email
  // failure throw: that would leave the submission un-marked in
  // runOncePerKey_ and a retry would overwrite the status with
  // 'Already "Not Renewing"', hiding who made the change. Alert instead.
  try {
    sendTerminationRequestOpsEmail_(req, status, company, previousStatus);
  } catch (err) {
    notifyError_(`Termination request Ops email (row ${rowNum}, pilot ${req.pilotNumber}) — result was: ${status}`, err);
  }
}

/**
 * Decide what to do with a request, given TRACKER's values (pure — no writes).
 * Auto-approval needs the pilot number to match exactly one row AND the email
 * to match that row's Company Email.
 * @returns {{ok:boolean, status:string, rowNum?:number, row?:Array}}
 *          row/rowNum are set whenever a single pilot-number match was found.
 */
function matchTerminationRequest_(data, req) {
  const C = CONFIG.TRACKER_COLS;
  const pilot = normPilot_(req.pilotNumber);
  const email = normEmail_(req.companyEmail);

  const rows = [];
  if (pilot) {
    for (let i = 2; i < data.length; i++) {          // data starts at row 3
      if (normPilot_(data[i][C.PILOT_NUMBER - 1]) === pilot) rows.push(i);
    }
  }

  if (rows.length > 1) {
    return { ok: false, status: `NEEDS REVIEW — ${rows.length} TRACKER rows have this pilot number, nothing changed` };
  }
  if (rows.length === 0) {
    return { ok: false, status: 'NEEDS REVIEW — pilot number not found in TRACKER, nothing changed' };
  }

  const i = rows[0];
  const registered = normEmail_(data[i][C.COMPANY_EMAIL - 1]);
  if (!email || email !== registered) {
    return { ok: false, rowNum: i + 1, row: data[i],
             status: 'NEEDS REVIEW — email does not match the registered Company Email, nothing changed' };
  }
  return { ok: true, rowNum: i + 1, row: data[i], status: 'matched' };
}

function isInArchived_(ss, req) {
  const sheet = ss.getSheetByName(CONFIG.ARCHIVED_SHEET);
  const pilot = normPilot_(req.pilotNumber);
  if (!sheet || !pilot) return false;
  return sheet.getDataRange().getValues().slice(2)
    .some(r => normPilot_(r[CONFIG.TRACKER_COLS.PILOT_NUMBER - 1]) === pilot);
}

/**
 * Reasons from verified termination requests, keyed by normalised pilot number
 * (latest submission wins — by Timestamp, not row order, so sorting the tab
 * can't change the result). Used by [04] Sync Renewals via
 * terminationReasonFor_(). NEEDS REVIEW rows are kept only when the pilot
 * number matched a TRACKER company — Ops approves those by setting
 * "Not Renewing" by hand, and the reason should still reach AstriCloud. Rows
 * with no matched company (unknown/archived pilot number) are skipped.
 * terminationReasonFor_ is only used for rows Sync is terminating, so an
 * unapproved request's reason is never shown.
 * @returns {Object<string,{reason:string, at:Date|null}>}
 */
function getTerminationReasons_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TERMINATION.SHEET);
  if (!sheet || sheet.getLastRow() < 2) return {};

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => (h || '').toString().trim());
  const pilotCol  = headers.indexOf(TERMINATION.Q.PILOT_NUMBER);
  const reasonCol = headers.indexOf(TERMINATION.Q.REASON);
  const statusCol = headers.indexOf(TERMINATION.STATUS_HEADERS[0]);
  const matchCol  = headers.indexOf(TERMINATION.STATUS_HEADERS[1]);
  const timeCol   = headers.indexOf('Timestamp');
  if (pilotCol < 0 || reasonCol < 0 || statusCol < 0) return {};

  const reasons = {};
  for (let i = 1; i < data.length; i++) {
    const status  = (data[i][statusCol] || '').toString();
    const matched = matchCol >= 0 && (data[i][matchCol] || '').toString().trim() !== '';
    const pilot   = normPilot_(data[i][pilotCol]);
    const reason  = (data[i][reasonCol] || '').toString().trim();
    if (!pilot || !reason || !status) continue;
    if (status.indexOf('NEEDS REVIEW') === 0 && !matched) continue;
    const at = timeCol >= 0 && data[i][timeCol] ? new Date(data[i][timeCol]) : null;
    const prev = reasons[pilot];
    const t = at && !isNaN(at.getTime()) ? at.getTime() : -Infinity;
    const prevT = prev && prev.at && !isNaN(prev.at.getTime()) ? prev.at.getTime() : -Infinity;
    if (!prev || t >= prevT) reasons[pilot] = { reason: reason, at: at };   // newest Timestamp wins
  }
  return reasons;
}

/**
 * The reason to show for a termination, or '' — only if the request was
 * submitted during the current contract cycle (the 12 months ending on
 * contractEnd), so a reason from an earlier cycle is never reused.
 */
function terminationReasonFor_(reasons, pilotNumber, contractEnd) {
  const r = reasons[normPilot_(pilotNumber)];
  if (!r || !r.reason) return '';
  if (!r.at || isNaN(r.at.getTime()) || !contractEnd) return '';
  const end = new Date(contractEnd);
  const cycleStart = new Date(end.getFullYear() - 1, end.getMonth(), end.getDate() + 1);
  return r.at >= cycleStart ? r.reason : '';
}

// Compare phone numbers by digits only: ignores spaces, dashes, dots, brackets,
// the leading 0 (Sheets drops it from stored numbers) and the Malaysian country
// code (+60 / 60 / 0060). 03-3015008, 03 301 5008, +60 3-3015008 → "33015008".
// Local numbers without their 0 are at most 9-10 digits and never start "60",
// so a 10+ digit value starting "60" can only be the country code.
function normPilot_(v) {
  let d = (v || '').toString().replace(/\D/g, '').replace(/^0+/, '');
  if (d.length >= 10 && d.indexOf('60') === 0) d = d.slice(2).replace(/^0+/, '');
  return d;
}
function normEmail_(v) { return (v || '').toString().trim().toLowerCase(); }

/** Write status values into the named status columns (created if missing). */
function writeTerminationStatus_(sheet, rowNum, values) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => (h || '').toString().trim());

  TERMINATION.STATUS_HEADERS.forEach((title, i) => {
    let col = headers.indexOf(title) + 1;
    if (col === 0) {
      col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue(title).setFontWeight('bold');
      headers[col - 1] = title;
    }
    const cell = sheet.getRange(rowNum, col);
    cell.setValue(values[i]);
    if (values[i] instanceof Date) cell.setNumberFormat(i === 3 ? 'dd-MMM-yyyy HH:mm' : 'dd-MMM-yyyy');
  });

  // Highlight rows that need a person to look at them
  if (String(values[0]).indexOf('NEEDS REVIEW') === 0) {
    sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).setBackground('#FCE8E6');
  }
}

/** Email Ops about a new termination request. Throws on failure (→ notifyError_). */
function sendTerminationRequestOpsEmail_(req, status, company, previousStatus) {
  const tz = Session.getScriptTimeZone();
  const to = getConfigValue(TERMINATION.NOTIFY_CONFIG_KEY) || CONFIG.ALERT_EMAIL;
  const needsReview = status.indexOf('NEEDS REVIEW') === 0;
  const esc = escHtml_;
  const endStr = company && company.contractEnd
    ? Utilities.formatDate(new Date(company.contractEnd), tz, 'dd MMM yyyy') : '—';

  const row = (label, value) =>
    `<tr><td style="border:1px solid #000;padding:6px 12px;"><strong>${label}</strong></td>` +
    `<td style="border:1px solid #000;padding:6px 12px;">${value}</td></tr>`;

  // One row per field. TRACKER values win; the submitted value is shown only
  // when it differs, so Ops can spot a mismatch at a glance.
  const differs = (a, b, norm) => a && b && norm(a) !== norm(b);
  const flag = (label, v) => ` <span style="color:#cc0000;">(${label}: ${esc(v)})</span>`;
  const normName = v => String(v || '').toLowerCase().replace(/[^a-z0-9]/g, '');

  let rows = '';
  rows += row('Company', company
    ? esc(company.name) + (differs(req.companyName, company.name, normName) ? flag('form says', req.companyName) : '')
    : esc(req.companyName) + ' <em>(not found in TRACKER)</em>');
  rows += row('Pilot Number', esc(req.pilotNumber));
  rows += row('Company Email', esc(req.companyEmail) +
    (company && differs(req.companyEmail, company.email, normEmail_) ? flag('registered', company.email) : ''));
  if (company) {
    rows += row('WORQ Location', esc(company.location) || '—');
    rows += row('Contract End Date', endStr);
  }
  if (previousStatus) rows += row('Previous Renewal Status', esc(previousStatus));
  if (req.reason)     rows += row('Reason', esc(req.reason));
  rows += row('Result', `<strong style="color:${needsReview ? '#cc0000' : '#007700'};">${esc(status)}</strong>`);

  const title = toProperCase(company ? company.name : req.companyName) || `Pilot ${req.pilotNumber || 'unknown'}`;
  const subject = `${needsReview ? '⚠️ [Needs review] ' : ''}Virtual Landline Termination Request - ${title}`;

  const htmlBody = `
<p>Hi Team,</p>
<p>A client has submitted the Virtual Landline <strong>Termination Request</strong> form.</p>
<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:13px;">
  ${rows}
</table>
<br>
${needsReview
  ? `<p><strong>Action needed:</strong> check the request with the client <strong>using the registered Company Email in TRACKER</strong> — not the email typed on the form, which is unverified. If it is genuine, set Renewal Status (col F) to "Not Renewing" in TRACKER manually.</p>`
  : `<p><strong>Next step:</strong> run <em>🤖 AstriCloud Automation → [04] Sync Renewals</em>. That sends the client's termination confirmation and the AstriCloud notification. If the request is not genuine, change Renewal Status (col F) back before syncing.</p>`}
<p>Note: service continues until the contract end date even if the client asked for early termination (contract is paid in full upfront).</p>
<p><a href="${SpreadsheetApp.getActiveSpreadsheet().getUrl()}">Open the AstriCloud Tracker</a></p>
<p style="color:#888;font-size:12px;">Automated message from the AstriCloud Tracker.</p>`;

  // CC the outlet (Addresses sheet) when we know which company this is
  const outletEmail = company ? getLocationEmail(company.location) : null;

  const mail = {
    to: to,
    subject: subject,
    htmlBody: htmlBody,
    name: 'AstriCloud Tracker', // internal email — same sender name as the boss's alerts/summary
    // Never the form's email: it is unverified, so on a NEEDS REVIEW request
    // "Reply" could reach an impostor. Use the registered TRACKER email.
    replyTo: (company && company.email) ? String(company.email).trim() : 'it@worq.space'
  };
  if (outletEmail) mail.cc = outletEmail;
  sendMail_(mail);
  Logger.log(`sendTerminationRequestOpsEmail_: sent to ${to}${outletEmail ? ` | CC: ${outletEmail}` : ''} | ${status}`);
}

// ─────────────────────────────────────────────────────────────────────────────
// One-time setup — run from the Apps Script editor
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Create the Termination Request form, link it to this spreadsheet, rename its
 * response tab to "Termination Requests", and add TERMINATION_NOTIFY_EMAIL to
 * the Config sheet (default it@worq.space; the outlet is CC'd automatically).
 * Safe to re-run: stops if the tab already exists.
 *
 * NOTE: when run from the editor, it waits until you click OK on the popup in
 * the spreadsheet tab.
 */
function createTerminationForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  if (ss.getSheetByName(TERMINATION.SHEET)) {
    ui.alert(`ℹ️ "${TERMINATION.SHEET}" tab already exists — the form is already set up. Nothing done.`);
    return;
  }

  // A form was created before but its tab never got renamed: finish that
  // setup instead of creating a second form.
  const existingId = PropertiesService.getScriptProperties().getProperty(TERMINATION.FORM_ID_PROPERTY);
  if (existingId) {
    const linked = ss.getSheets().find(s => (s.getFormUrl() || '').indexOf(existingId) !== -1);
    if (linked) {
      const oldName = linked.getName();
      linked.setName(TERMINATION.SHEET);
      seedTerminationNotifyConfig_(ss);
      ui.alert(`✅ Existing Termination Request form found — renamed its tab "${oldName}" to "${TERMINATION.SHEET}". No new form created.`);
    } else {
      ui.alert(
        `⚠️ A Termination Request form was already created (ID ${existingId}) but no tab in this sheet is linked to it.\n\n` +
        `Open that form → Responses → Link to Sheets → this spreadsheet, then run createTerminationForm() again. ` +
        `No new form was created.`
      );
    }
    return;
  }

  const form = FormApp.create('AstriCloud Virtual Landline - Termination Request');
  form.setDescription(
    'Use this form if you do not wish to renew your WORQ virtual landline service.\n\n' +
    'Your service will remain active until the end of your current contract period. ' +
    'We will email you a confirmation once your request has been processed.'
  );
  form.setConfirmationMessage('Thank you. We have received your termination request and will email you a confirmation once it is processed.');
  form.setCollectEmail(false);
  try { form.setRequireLogin(false); } catch (err) { Logger.log('setRequireLogin not available: ' + err.message); }

  form.addTextItem().setTitle(TERMINATION.Q.COMPANY_NAME).setRequired(true);
  form.addTextItem().setTitle(TERMINATION.Q.PILOT_NUMBER)
    .setHelpText('Your virtual landline number, e.g. 0386843817.')
    .setRequired(true);
  form.addTextItem().setTitle(TERMINATION.Q.COMPANY_EMAIL)
    .setHelpText('The company email used when you signed up for the virtual landline.')
    .setValidation(FormApp.createTextValidation().requireTextIsEmail().build())
    .setRequired(true);
  form.addParagraphTextItem().setTitle(TERMINATION.Q.REASON).setRequired(false);
  form.addCheckboxItem().setTitle(TERMINATION.Q.ACK)
    .setChoiceValues(['I understand that my virtual landline service will remain active until the end of my current contract period, and that fees already paid for the contract are not refundable.'])
    .setRequired(true);

  // Link to this spreadsheet — Google adds a new "Form Responses N" tab
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  PropertiesService.getScriptProperties().setProperty(TERMINATION.FORM_ID_PROPERTY, form.getId());

  // Find the new tab (can take a moment to appear) and rename it
  let respSheet = null;
  for (let attempt = 0; attempt < 10 && !respSheet; attempt++) {
    SpreadsheetApp.flush();
    respSheet = SpreadsheetApp.openById(ss.getId()).getSheets()
      .find(s => (s.getFormUrl() || '').indexOf(form.getId()) !== -1) || null;
    if (!respSheet) Utilities.sleep(1500);
  }
  if (respSheet) respSheet.setName(TERMINATION.SHEET);

  seedTerminationNotifyConfig_(ss);

  ui.alert(
    '✅ Termination Request form created\n\n' +
    `Share link (send to clients):\n${form.getPublishedUrl()}\n\n` +
    `Edit link:\n${form.getEditUrl()}\n\n` +
    (respSheet
      ? `Responses go to the "${TERMINATION.SHEET}" tab.\n\n`
      : `⚠️ SETUP NOT FINISHED — could not find the new response tab yet. Termination requests are IGNORED until it is named "${TERMINATION.SHEET}". ` +
        `Wait a minute and run createTerminationForm() again (it will rename the tab, not create a second form), or rename the newest "Form Responses" tab by hand.\n\n`) +
    `Ops notifications go to the ${TERMINATION.NOTIFY_CONFIG_KEY} value in the Config tab.\n\n` +
    'The form was created in your My Drive — move it next to this spreadsheet if you like.'
  );
}

/** Add TERMINATION_NOTIFY_EMAIL (it@worq.space) to Config if missing — never overwrites. */
function seedTerminationNotifyConfig_(ss) {
  if (getConfigValue(TERMINATION.NOTIFY_CONFIG_KEY)) return;
  const cfg = ss.getSheetByName(CONFIG.CONFIG_SHEET);
  if (cfg) cfg.appendRow([TERMINATION.NOTIFY_CONFIG_KEY, CONFIG.ALERT_EMAIL]);
}
