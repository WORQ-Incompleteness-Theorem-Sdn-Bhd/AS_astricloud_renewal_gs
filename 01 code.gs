/**
 * AstriCloud Tracker Automation System
 * Main coordination file
 */

// Configuration
var CONFIG = {
  TRACKER_SHEET: 'TRACKER',
  FORM_RESPONSES_SHEET: 'Form Responses 1',
  RENEWAL_SHEET: 'Renewal Status',
  ARCHIVED_SHEET: 'ARCHIVED',
  EMAIL_FROM: 'it_worq@worq.space',

  // Column positions in TRACKER
  TRACKER_COLS: {
    NO: 1,              // A
    COMPANY_NAME: 2,    // B
    WORQ_LOCATION: 3,   // C
    COMPANY_EMAIL: 4,   // D
    PILOT_NUMBER: 5,    // E
    RENEWAL_STATUS: 6,  // F
    CONTRACT_START: 7,  // G
    CONTRACT_END: 8,    // H
    FIRST_MONTH: 9      // I (Feb-2024)
  },

  // Column positions in FORM RESPONSES 1
  FORM_COLS: {
    TIMESTAMP: 1,       // A
    EMAIL: 2,           // B
    COMPANY_NAME: 3,    // C
    REG_NUMBER: 4,      // D (ignore)
    WORQ_LOCATION: 5    // E
    // F-X: ignore
  },

  // Renewal Status columns
  RENEWAL_COLS: {
    COMPANIES: 1,       // A
    OUTLET: 2,          // B
    CURRENT_START: 3,   // C
    CURRENT_END: 4,     // D
    CUSTOMER_STATUS: 5, // E
    NEW_RENEWAL_DATE: 6,// F
    FINAL_STATUS: 7     // G
  },

  // First month in tracker (Feb-2024)
  FIRST_MONTH_DATE: new Date('2024-02-01'),

  // Archived form responses (duplicate submissions)
  ARCHIVED_FORM_RESPONSES_SHEET: 'Archived Form Responses',

  // Runtime config sheet (Key/Value) — lets admins edit vendor recipients without a code deploy
  CONFIG_SHEET: 'Config',

  // Vendor notification (AstriCloud)
  VENDOR_EMAIL: 'theinosha@astricloud.com',
  VENDOR_CC: 'shahril@astricloud.com,mgt@astricloud.com,hussaini@astricloud.com,it@worq.space,accounts@astricloud.com,operations@worq.space,accounts@worq.space,sasikala@worq.space',



  // Reminder timing (months before expiry)
  REMINDER_MONTHS: [3, 2, 1, 0],

  // Renewal reminder escalation stages (TRACKER col F).
  // The monthly job (1st of each month) advances a company through these as
  // expiry approaches: 3 months out → 1st, 2 → 2nd, 1 → 3rd, 0 → Last.
  // 0 months = the expiry month itself, which gets distinct urgent wording.
  // Each stage is set at most once (never resends the same stage, never downgrades).
  // Admin still sets 'Renew' / 'Not Renewing' manually at any point.
  REMINDER_STAGES: {
    3: '1st Reminder Sent',
    2: '2nd Reminder Sent',
    1: '3rd Reminder Sent',
    0: 'Last Reminder Sent'
  },
  // Full dropdown value list for the Renewal Status column (col F).
  RENEWAL_STATUS_VALUES: [
    '1st Reminder Sent',
    '2nd Reminder Sent',
    '3rd Reminder Sent',
    'Last Reminder Sent',
    'Renew',
    'Renewed',
    'Not Renewing'
  ],

  // Error / failure alert recipient (scheduled triggers run unattended)
  ALERT_EMAIL: 'it@worq.space',

  // Failsafe recipient for the monthly reminder run summary (unattended runs)
  REMINDER_SUMMARY_EMAIL: 'it@worq.space'
};

/**
 * Ordered list of reminder stages, weakest → strongest, used to prevent
 * downgrading a company's status (e.g. never go from 'Last' back to '1st').
 */
var REMINDER_STAGE_ORDER = [
  '1st Reminder Sent',
  '2nd Reminder Sent',
  '3rd Reminder Sent',
  'Last Reminder Sent'
];

/**
 * Create custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 AstriCloud Automation')
    .addItem('[01] Copy New Entries from Form', 'copyNewEntriesToTracker')
    .addItem('[02] Check & Send Renewal Reminders', 'checkAndSendReminders')
    .addItem('[03] Backfill Missing Paid Status', 'backfillMissingPaidStatus')
    .addItem('[04] Sync Renewals from Renewal Status', 'syncRenewals')
    .addItem('[05] Archive Terminated Customers', 'archiveTerminated')
    .addItem('[06] Restore Archived Customer', 'showRestoreArchivedDialog')
    .addSeparator()
    .addItem('Sort by Contract Start Date', 'sortByContractStartDate')
    .addItem('Highlight Renewal Urgency', 'highlightRenewalUrgency')
    .addItem('Clear Renewal Highlights', 'clearRenewalHighlights')
    .addItem('Find Lapsed Contracts', 'findLapsedContracts')
    .addToUi();

  // ---------------------------------------------------------------------------
  // Not in the menu — run these from the Apps Script editor (Run > select fn).
  // They are one-time setup / maintenance, kept out of the menu to reduce noise.
  //
  //   checkTriggerHealth()          — verify the automation triggers are still
  //                                   installed. Worth running occasionally: a
  //                                   dropped trigger silently stops reminders
  //                                   or vendor signup emails.
  //   setupConfigSheet()            — create/seed the Config sheet (vendor emails)
  //   setupAutoReminderTrigger()    — install the monthly (1st of month) reminder
  //   setupFormSubmitTrigger()      — install the onFormSubmit trigger
  //   setupUrgencyFormatting()      — (re)build urgency colors on Contract End Date
  //   setupRenewalStatusDropdown()  — apply the col F dropdown to all rows
  //   migrateReminderStages()       — realign reminder stages, sends no emails
  //   removeAutoReminderTrigger()   — uninstall the monthly reminder trigger
  //   removeFormSubmitTrigger()     — uninstall the form submit trigger
  //   removeArchivedDuplicatesFromTracker()
  // ---------------------------------------------------------------------------
}

// Handler function installed by the monthly reminder trigger. Kept as a
// constant so setup / health-check / removal all agree on the name.
var MONTHLY_REMINDER_HANDLER = 'monthlyRenewalReminders';

/**
 * Install a monthly time-based trigger that runs monthlyRenewalReminders() on
 * the 1st of every month (between 8–9 AM). Idempotent — removes any existing
 * reminder trigger first so it can't double-fire.
 *
 * NOTE: this replaces the previous daily 'checkAndSendReminders' trigger. Any
 * old daily trigger pointing at checkAndSendReminders is cleaned up too.
 */
function setupAutoReminderTrigger() {
  const removed = removeReminderTriggers_();

  ScriptApp.newTrigger(MONTHLY_REMINDER_HANDLER)
    .timeBased()
    .onMonthDay(1)   // 1st of every month
    .atHour(8)
    .create();

  SpreadsheetApp.getUi().alert(
    `✅ Monthly auto-reminder trigger set\n\n` +
    `${MONTHLY_REMINDER_HANDLER}() will run automatically on the 1st of every month (~8 AM).\n\n` +
    `Each run escalates awaiting-reply customers: 3 months out → 1st Reminder, 2 → 2nd, 1 → Last.` +
    (removed > 0 ? `\n\n(${removed} previous reminder trigger(s) replaced)` : '')
  );
}

/**
 * Remove all time-based reminder triggers (both the new monthly handler and any
 * legacy daily 'checkAndSendReminders' trigger).
 */
function removeAutoReminderTrigger() {
  const removed = removeReminderTriggers_();
  SpreadsheetApp.getUi().alert(
    removed > 0
      ? `✅ Auto-reminder trigger removed (${removed}).\n\nRenewal reminders will no longer run automatically.`
      : `ℹ️ No auto-reminder trigger found.\n\nNothing to remove.`
  );
}

/** Internal: delete reminder triggers (monthly handler + legacy daily). Returns count removed. */
function removeReminderTriggers_() {
  const legacyHandlers = new Set([MONTHLY_REMINDER_HANDLER, 'checkAndSendReminders']);
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (legacyHandlers.has(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  return removed;
}

/**
 * Email an error alert to CONFIG.ALERT_EMAIL. Called from the catch blocks of
 * unattended triggered functions so silent failures become visible.
 * @param {string} context  - name of the function/operation that failed
 * @param {Error}  error    - the caught error
 */
function notifyError_(context, error) {
  const tz = Session.getScriptTimeZone();
  const when = Utilities.formatDate(new Date(), tz, 'dd MMM yyyy HH:mm:ss');
  const message = (error && error.message) ? error.message : String(error);
  const stack   = (error && error.stack) ? error.stack : '(no stack)';

  Logger.log(`ERROR in ${context}: ${message}`);

  try {
    MailApp.sendEmail({
      to: CONFIG.ALERT_EMAIL,
      subject: `⚠️ AstriCloud Tracker error — ${context}`,
      htmlBody:
        `<p>An automated task in the AstriCloud Tracker failed.</p>` +
        `<p><strong>Function:</strong> ${context}<br>` +
        `<strong>Time:</strong> ${when}<br>` +
        `<strong>Error:</strong> ${message}</p>` +
        `<pre style="background:#f4f4f4;padding:10px;border:1px solid #ddd;white-space:pre-wrap;">${stack}</pre>` +
        `<p style="color:#888;font-size:12px;">Automated alert from the AstriCloud Tracker.</p>`,
      name: 'AstriCloud Tracker'
    });
  } catch (mailErr) {
    // Last resort — if even the alert email fails, at least log it.
    Logger.log(`notifyError_ FAILED to send alert for ${context}: ${mailErr.message}`);
  }
}

/**
 * Report the currently installed automation triggers, and warn if the expected
 * ones are missing. Time-based triggers run unattended — if one is silently
 * dropped (script copy, owner change, manual delete), all reminders stop with
 * no warning. Run this from the menu to confirm the automation is live.
 */
function checkTriggerHealth() {
  const triggers = ScriptApp.getProjectTriggers();
  const handlers = triggers.map(t => t.getHandlerFunction());

  const hasReminder = handlers.includes(MONTHLY_REMINDER_HANDLER);
  const hasFormSubmit = handlers.includes('onFormSubmit');

  const lines = triggers.length
    ? triggers.map(t => `• ${t.getHandlerFunction()} (${t.getEventType()})`).join('\n')
    : '(none installed)';

  const warnings = [];
  if (!hasReminder) {
    warnings.push(`❌ Monthly reminder trigger (${MONTHLY_REMINDER_HANDLER}) is NOT installed — renewal reminders will not run. Fix: menu → "Setup Monthly Auto-Reminder Trigger".`);
  }
  if (!hasFormSubmit) {
    warnings.push(`❌ Form-submit trigger (onFormSubmit) is NOT installed — new signups won't notify the vendor. Fix: run setupFormSubmitTrigger() from the Apps Script editor.`);
  }

  const status = warnings.length === 0
    ? '✅ All expected triggers are installed.'
    : '⚠️ Missing triggers:\n\n' + warnings.join('\n\n');

  SpreadsheetApp.getUi().alert(
    `Trigger Health Check\n\n` +
    `Installed triggers (${triggers.length}):\n${lines}\n\n` +
    status
  );
}
