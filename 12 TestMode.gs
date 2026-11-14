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
 */

var TEST_MODE_CONFIG_KEY = 'TEST_EMAIL_OVERRIDE';

/**
 * @returns {string|null} override address when test mode is on.
 * No try/catch on purpose (fail closed): if the Config sheet can't be read,
 * the error propagates and nothing is sent, instead of silently treating a
 * test copy as live and emailing real recipients. Callers already surface
 * send failures (alert email / "please send manually" in the summary).
 */
function getTestEmailOverride_() {
  return getConfigValue(TEST_MODE_CONFIG_KEY);
}

/**
 * Drop-in replacement for MailApp.sendEmail(options) (object form only).
 * @param {Object} options - same fields as MailApp.sendEmail
 */
function sendMail_(options) {
  const override = getTestEmailOverride_();
  if (!override) {
    MailApp.sendEmail(options);
    return;
  }

  const esc = s => String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
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
