/**
 * Test mode — makes it safe to run the automation on a copy of the sheet.
 *
 * Every email the script sends goes through sendMail_(). If the Config sheet
 * has a TEST_EMAIL_OVERRIDE value, ALL emails (customers, AstriCloud, outlet
 * CCs, Ops, error alerts) go ONLY to that address instead, with "[TEST]" in
 * the subject and the original recipients listed at the top of the email.
 *
 * Leave TEST_EMAIL_OVERRIDE out of Config (or blank) on the live sheet —
 * then emails are sent exactly as before.
 *
 * Any copy of the spreadsheet (ID ≠ CONFIG.PRODUCTION_SPREADSHEET_ID) with no
 * override is blocked from sending client/vendor email, so a copy with a
 * missing/renamed Config tab or a misspelled key can never email real people.
 */

var TEST_MODE_CONFIG_KEY = 'TEST_EMAIL_OVERRIDE';

// Per-execution cache (Apps Script globals reset on every run), so a loop of
// N emails reads the Config sheet once, not N times.
var testEmailOverrideCache_;

/**
 * @returns {string|null} override address when test mode is on.
 * No try/catch on purpose (fail closed): if the Config sheet can't be read,
 * the error propagates and nothing is sent, instead of silently treating a
 * test copy as live and emailing real recipients. Callers already surface
 * send failures (alert email / "please send manually" in the summary).
 */
function getTestEmailOverride_() {
  if (testEmailOverrideCache_ === undefined) {
    testEmailOverrideCache_ = getConfigValue(TEST_MODE_CONFIG_KEY);
  }
  return testEmailOverrideCache_;
}

function isProductionSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getId() === CONFIG.PRODUCTION_SPREADSHEET_ID;
}

/**
 * Drop-in replacement for MailApp.sendEmail(options) (object form only).
 * @param {Object} options - same fields as MailApp.sendEmail
 * @param {{internal?: boolean}} [opts] - internal: true for IT-only mail
 *        (alerts, summaries) that must still go out if Config is unreadable
 */
function sendMail_(options, opts) {
  let override;
  try {
    override = getTestEmailOverride_();
  } catch (err) {
    // Fail closed for client/vendor email. Internal-only mail (error alerts,
    // monthly summary — always to it@worq.space, no CC) is still sent as-is,
    // so a Config failure is never silent.
    if (opts && opts.internal) {
      Logger.log(`sendMail_: Config unreadable (${err.message}) — sending internal mail as-is`);
      MailApp.sendEmail(Object.assign({}, options, {
        subject: '[CONFIG UNREADABLE] ' + (options.subject || ''),
        htmlBody: `<p style="color:#cc0000;"><strong>Note:</strong> the Config sheet could not be read (${escHtml_(err.message)}). Client and AstriCloud emails are blocked until this is fixed.</p>` + (options.htmlBody || '')
      }));
      return;
    }
    throw err;
  }
  if (!override) {
    if (!isProductionSpreadsheet_() && !(opts && opts.internal)) {
      throw new Error(
        `Email blocked: this is a copy of the tracker (not production) and Config has no ${TEST_MODE_CONFIG_KEY}. ` +
        `Add ${TEST_MODE_CONFIG_KEY} to the Config tab to test. Blocked: "${options.subject}" to ${options.to}`
      );
    }
    MailApp.sendEmail(options);
    return;
  }

  const esc = escHtml_;
  const test = Object.assign({}, options);
  delete test.cc;
  delete test.bcc;
  test.to = override;
  test.replyTo = override; // a reply from the tester must never reach the real client
  test.subject = '[TEST] ' + (options.subject || '');

  const banner =
    `<div style="border:2px dashed #cc0000;background:#fff4f4;padding:10px 14px;margin-bottom:16px;font-family:Arial,sans-serif;font-size:12px;">` +
    `<strong style="color:#cc0000;">🧪 TEST MODE — this email was NOT sent to the real recipients.</strong><br>` +
    `<strong>Would have gone to:</strong> ${esc(options.to) || '—'}<br>` +
    `<strong>CC:</strong> ${esc(options.cc) || '—'}` +
    (options.replyTo ? `<br><strong>Reply-To:</strong> ${esc(options.replyTo)}` : '') +
    (options.bcc ? `<br><strong>BCC:</strong> ${esc(options.bcc)}` : '') +
    `<br><span style="color:#888;">Turn off: remove ${TEST_MODE_CONFIG_KEY} from the Config tab.</span></div>`;

  if (options.htmlBody) test.htmlBody = banner + options.htmlBody;
  if (options.body || !options.htmlBody) {
    test.body = `TEST MODE — would have gone to: ${options.to || '-'} | CC: ${options.cc || '-'}\n\n` + (options.body || '');
  }

  MailApp.sendEmail(test);
  Logger.log(`sendMail_ [TEST MODE]: "${options.subject}" redirected to ${override} (was to: ${options.to} | cc: ${options.cc || '-'})`);
}
