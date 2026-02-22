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

  // Reminder timing (months before expiry)
  REMINDER_MONTHS: [3, 2, 1, 0]
};

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
    // .addItem('Setup Renewal Status Dropdown', 'setupRenewalStatusDropdown')
    // .addItem('Remove Archived Duplicates from Tracker', 'removeArchivedDuplicatesFromTracker')
    // .addSeparator()
    // .addItem('Setup Auto-Reminder Trigger (Daily 9 AM)', 'setupAutoReminderTrigger')
    // .addItem('Remove Auto-Reminder Trigger', 'removeAutoReminderTrigger')
}

/**
 * Install a daily time-based trigger for checkAndSendReminders at 9 AM.
 * Replaces any existing reminder trigger to prevent duplicates.
 */
function setupAutoReminderTrigger() {
  const triggers = ScriptApp.getProjectTriggers();

  // Remove any existing reminder triggers
  let removed = 0;
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'checkAndSendReminders') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });

  // Create new daily trigger at 9 AM
  ScriptApp.newTrigger('checkAndSendReminders')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  SpreadsheetApp.getUi().alert(
    `✅ Auto-reminder trigger set\n\n` +
    `checkAndSendReminders() will run automatically every day at 9 AM.` +
    (removed > 0 ? `\n\n(${removed} previous trigger(s) replaced)` : '')
  );
}

/**
 * Remove all time-based triggers for checkAndSendReminders.
 */
function removeAutoReminderTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;

  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'checkAndSendReminders') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });

  SpreadsheetApp.getUi().alert(
    removed > 0
      ? `✅ Auto-reminder trigger removed.\n\ncheckAndSendReminders() will no longer run automatically.`
      : `ℹ️ No auto-reminder trigger found.\n\nNothing to remove.`
  );
}
