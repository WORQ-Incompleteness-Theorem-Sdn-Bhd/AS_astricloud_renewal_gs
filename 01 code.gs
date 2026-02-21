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
    CONTRACT_START: 6,  // F
    CONTRACT_END: 7,    // G
    FIRST_MONTH: 8      // H (Feb-2024)
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
  REMINDER_MONTHS: [3, 2, 1]
};

/**
 * Create custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 AstriCloud Automation')
    .addItem('📋 Copy New Entries from Form', 'copyNewEntriesToTracker')
    .addItem('🎨 Highlight Current Month', 'highlightCurrentMonth')
    .addItem('📧 Check & Send Renewal Reminders', 'checkAndSendReminders')
    .addItem('🔄 Sync Renewals from Renewal Status', 'syncRenewals')
    .addItem('📦 Archive Terminated Customers', 'archiveTerminated')
    .addSeparator()
    .addItem('⚙️ Setup Time-based Triggers', 'setupTriggers')
    .addItem('🗑️ Remove All Triggers', 'removeAllTriggers')
    .addToUi();
}

/**
 * Install time-based triggers for automation
 */
function setupTriggers() {
  // Remove existing triggers first
  removeAllTriggers();
  
  // Daily at 8 AM - Copy new entries
  ScriptApp.newTrigger('copyNewEntriesToTracker')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  
  // Daily at 9 AM - Check reminders
  ScriptApp.newTrigger('checkAndSendReminders')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
  
  // Daily at 7 AM - Highlight current month
  ScriptApp.newTrigger('highlightCurrentMonth')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
  
  // Daily at 10 AM - Sync renewals
  ScriptApp.newTrigger('syncRenewals')
    .timeBased()
    .everyDays(1)
    .atHour(10)
    .create();
  
  SpreadsheetApp.getUi().alert('✅ Triggers installed successfully!\n\nAutomation will run daily at:\n- 7 AM: Highlight current month\n- 8 AM: Copy new entries\n- 9 AM: Send reminders\n- 10 AM: Sync renewals');
}

/**
 * Remove all project triggers
 */
function removeAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  Logger.log('All triggers removed');
}