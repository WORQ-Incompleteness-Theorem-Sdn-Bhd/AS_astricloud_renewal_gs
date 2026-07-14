# AstriCloud Tracker Automation System

## Overview

A Google Apps Script automation system that manages customer contract tracking for WORQ's virtual landline (AstriCloud) service. The system automates the full lifecycle of customer contracts: intake from a Google Form, payment tracking on a monthly grid, renewal reminders via email, renewal processing, archival of terminated customers, and restoration of archived customers who renew at the last minute.

## Google Workspace Links

| Resource | Testing | Production |
|----------|---------|------------|
| Google Sheet | [Testing Sheet](https://docs.google.com/spreadsheets/d/1USIQXESClxQ_DHVD7qSLtAh8I1JbJUtLSg3KVDNlbn0/edit?pli=1&gid=266634575#gid=266634575) | [Production Sheet](https://docs.google.com/spreadsheets/d/1t_-C-TZjd7dN6uweYG3wdZkToVAG4pxfEKInIG-TolE/edit?pli=1&gid=266634575#gid=266634575) |
| Apps Script | [Testing Script](https://script.google.com/u/0/home/projects/1T-z8HKXroRAMV8a9WBF0NdlQG38FrjHxPmLlJ5cXblBuioVFHwlObFlK) | [Production Script](https://script.google.com/u/0/home/projects/1RCLJ-hDSe_9ESl4LpHEfaYhaM5BhD8MpOEJ0XgLYDIx9ZSVxO9n_EQvH/) |
| Google Form | [Testing Form](https://docs.google.com/forms/d/1rC0suleIU9AqaQWaM3uPkZs7KlLo1O9wWugOgYYrYvc/edit) | [Production Form](https://docs.google.com/forms/d/1UAEJlLobYMUNB5JJFadbjHIEs1TmA90k6-1plIgERhs/preview) |

---

## Sheet Structure

The spreadsheet contains 6 active sheets:

### 1. TRACKER (Main sheet)
The central tracking sheet with a monthly payment grid. It has **two header rows**:
- **Row 1**: Year group labels (2024, 2025, 2026, ...) spanning the month columns
- **Row 2**: Month column names (Feb-2024, Mar-2024, ...)
- **Row 3+**: Customer data

| Column | Field | Description |
|--------|-------|-------------|
| A | NO | Auto-incremented row number via SEQUENCE formula in A3 — never written to directly by scripts |
| B | Company Name | Customer company name |
| C | WORQ Location | Which WORQ outlet the customer belongs to |
| D | Company Email | Contact email for the customer |
| E | Pilot Number | Virtual landline number (triggers auto-population when entered) |
| F | Renewal Status | Tracks renewal lifecycle: `1st/2nd/3rd/Last Reminder Sent`, `Renew`, `Renewed`, `Not Renewing`, `Terminated` |
| G | Contract Start | Auto-set to 1st of the current month when pilot number is entered |
| H | Contract End | Auto-set to last day of the 12th month from contract start |
| I+ | Monthly columns (Feb-2024, ...) | Payment status: `paid`, `renew`, `terminate`, `not proceed` |

### 2. Form Responses 1
Auto-populated by the linked Google Form.

| Column | Field | Used? |
|--------|-------|-------|
| A | Timestamp | Ignored |
| B | Email | Copied to TRACKER col D |
| C | Company Name | Copied to TRACKER col B |
| D | Registration Number | Ignored |
| E | WORQ Location | Copied to TRACKER col C |
| F-X | Other fields | Ignored |

### 3. ARCHIVED
Storage for terminated customer records. Same column structure as TRACKER (col A = SEQUENCE formula, col B onwards = customer data). Terminated companies are never re-added to TRACKER by the copy function. Archived companies can be restored back to TRACKER via `[06] Restore Archived Customer`.

### 4. Archived Form Responses
Storage for duplicate form submissions. When the `onFormSubmit` trigger detects that an incoming submission's email already exists in a prior row of Form Responses 1, the row is moved here automatically. Same column structure as Form Responses 1 (Timestamp, Email, Company Name, Registration Number, WORQ Location, ...). Created automatically on first duplicate detection if the sheet does not exist.

### 5. Addresses
Lookup table mapping WORQ location names to their location-specific email addresses. Used by `getLocationEmail()` to CC the correct location inbox on customer emails.

| Column | Field | Description |
|--------|-------|-------------|
| A | Site | WORQ location name (must match exactly what appears in TRACKER col C) |
| B | Address | Full address of the location (informational only) |
| C | Emails | Location-specific email address (e.g. `ttdi@worq.space`) |

### 6. Config
Runtime Key/Value settings sheet. Lets admins change vendor email recipients **without a code deploy** — the scripts read these values at send time via `getConfigValue()`, falling back to the hardcoded `CONFIG` values in `01 code.gs` if the sheet or a key is missing/blank. Created and seeded on demand via the **Setup / Refresh Config Sheet (Vendor Emails)** menu item (`setupConfigSheet()`).

| Column | Field | Description |
|--------|-------|-------------|
| A | Key | Setting name (must match exactly, case-insensitive). Row 1 is the header. |
| B | Value | Setting value |

**Recognised keys:**
| Key | Used by | Meaning |
|-----|---------|---------|
| `VENDOR_EMAIL` | `getVendorEmail()` | Main "To" recipient for vendor notification emails (AstriCloud) |
| `VENDOR_CC` | `getVendorCc()` | Comma-separated CC list for vendor notification emails |

> Edit column B to add/remove recipients — changes apply on the next email, no `clasp push` required. The Config sheet is the source of truth; the `CONFIG.VENDOR_EMAIL` / `CONFIG.VENDOR_CC` values in code are only a fallback used when the sheet or key is absent.

---

## File Structure & Functions

### 01 code.gs — Configuration & Triggers
The main entry point. Contains the global `CONFIG` object (declared with `var` for cross-file accessibility in Apps Script V8) and menu/trigger management.

**CONFIG.TRACKER_COLS:**
| Key | Column | Description |
|-----|--------|-------------|
| NO | A (1) | Row number |
| COMPANY_NAME | B (2) | Company name |
| WORQ_LOCATION | C (3) | WORQ outlet |
| COMPANY_EMAIL | D (4) | Contact email |
| PILOT_NUMBER | E (5) | Virtual landline number |
| RENEWAL_STATUS | F (6) | Renewal lifecycle status |
| CONTRACT_START | G (7) | Contract start date |
| CONTRACT_END | H (8) | Contract end date |
| FIRST_MONTH | I (9) | First month column (Feb-2024) |

**CONFIG.ARCHIVED_FORM_RESPONSES_SHEET:** Sheet name for duplicate form submissions (`'Archived Form Responses'`).

**CONFIG.CONFIG_SHEET:** Sheet name for the runtime Key/Value config sheet (`'Config'`).

**CONFIG.VENDOR_EMAIL:** *Fallback* main recipient for vendor notification emails (AstriCloud). At runtime the value is read from the **Config** sheet via `getVendorEmail()`; this hardcoded value is used only if the sheet/key is missing.

**CONFIG.VENDOR_CC:** *Fallback* comma-separated CC list for vendor notification emails (AstriCloud team + WORQ internal recipients). At runtime the value is read from the **Config** sheet via `getVendorCc()`; this hardcoded value is used only if the sheet/key is missing.

**CONFIG.REMINDER_MONTHS:** `[3, 2, 1, 0]` — the months-before-expiry thresholds at which reminders fire (0 = the expiry month itself).

**CONFIG.REMINDER_STAGES:** Maps months-to-expiry → the status written to TRACKER col F. This is the escalation ladder the monthly job walks:

| Months left | Status set | Email sent |
|---|---|---|
| 3 | `1st Reminder Sent` | Standard countdown reminder |
| 2 | `2nd Reminder Sent` | Standard countdown reminder |
| 1 | `3rd Reminder Sent` | Standard countdown reminder |
| 0 | `Last Reminder Sent` | **Final notice** — distinct urgent wording |

**CONFIG.RENEWAL_STATUS_VALUES:** The full dropdown list for col F — the four reminder stages above plus `Renew`, `Renewed`, `Not Renewing`.

**REMINDER_STAGE_ORDER** (module-level, not in CONFIG): the four reminder stages ordered weakest → strongest. Used to guarantee a company's status is only ever *advanced*, never downgraded or re-sent.

**CONFIG.ALERT_EMAIL:** Recipient for automated failure alerts (`it@worq.space`). See `notifyError_()`.

**CONFIG.REMINDER_SUMMARY_EMAIL:** Recipient for the monthly reminder run summary (`it@worq.space`).

| Function | Type | Description |
|----------|------|-------------|
| `onOpen()` | Auto-trigger | Creates the custom menu when the spreadsheet is opened |
| `setupAutoReminderTrigger()` | Editor | Installs the monthly trigger — runs `monthlyRenewalReminders()` on the **1st of every month (~8 AM)**. Idempotent; also removes the legacy daily `checkAndSendReminders` trigger. |
| `removeAutoReminderTrigger()` | Editor | Removes the reminder trigger (monthly + any legacy daily). |
| `removeReminderTriggers_()` | Internal | Deletes reminder triggers; returns the count removed. |
| `notifyError_(context, error)` | Internal | Emails an error alert (function name, time, message, stack) to `CONFIG.ALERT_EMAIL`. Called from the catch blocks of unattended triggered functions so silent failures become visible. |
| `checkTriggerHealth()` | Editor | Lists installed triggers and warns if `monthlyRenewalReminders` or `onFormSubmit` is missing. **Worth running periodically** — a dropped trigger silently stops reminders or vendor signup emails with no other warning. |
| `columnLetter_(col)` | Internal | Converts a 1-based column number to its A1 letter (8 → H). Lives in `03 MonthHighlighter.gs`. |

**Menu items (in order):**
| # | Label | Function |
|---|-------|----------|
| [01] | Copy New Entries from Form | `copyNewEntriesToTracker()` |
| [02] | Check & Send Renewal Reminders | `checkAndSendReminders()` |
| [03] | Backfill Missing Paid Status | `backfillMissingPaidStatus()` |
| [04] | Sync Renewals from Renewal Status | `syncRenewals()` |
| [05] | Archive Terminated Customers | `archiveTerminated()` |
| [06] | Restore Archived Customer | `showRestoreArchivedDialog()` |
| — | *(separator)* | — |
| | Sort by Contract Start Date | `sortByContractStartDate()` |
| | Highlight Renewal Urgency | `highlightRenewalUrgency()` |
| | Clear Renewal Highlights | `clearRenewalHighlights()` |
| | Find Lapsed Contracts | `findLapsedContracts()` |

#### Editor-only functions (not in the menu)

One-time setup and maintenance functions are deliberately kept **out of the menu** to reduce clutter. Run them from the [Apps Script editor](https://script.google.com/u/0/home/projects/1RCLJ-hDSe_9ESl4LpHEfaYhaM5BhD8MpOEJ0XgLYDIx9ZSVxO9n_EQvH/): select the function in the dropdown at the top, then click **Run**. A matching list is kept in a comment block at the bottom of `onOpen()` in `01 code.gs`.

| Function | When to run |
|----------|-------------|
| `checkTriggerHealth()` | Periodically — verifies the automation triggers are still installed |
| `setupConfigSheet()` | Once — creates/seeds the **Config** sheet (vendor emails) |
| `setupAutoReminderTrigger()` | Once — installs the monthly reminder trigger |
| `setupFormSubmitTrigger()` | Once — installs the `onFormSubmit` trigger |
| `setupUrgencyFormatting()` | Once, or after changing the urgency colors — rebuilds conditional formatting on Contract End Date |
| `setupRenewalStatusDropdown()` | Once, or after changing `RENEWAL_STATUS_VALUES` — applies the col F dropdown to all rows |
| `migrateReminderStages()` | Rarely — realigns reminder stages to months-to-expiry. **Sends no emails.** |
| `removeAutoReminderTrigger()` | Only to disable automated reminders |
| `removeFormSubmitTrigger()` | Only to disable automated form handling |
| `removeArchivedDuplicatesFromTracker()` | One-time cleanup |

---

### 02 FormToTracker.gs — Form Data Processing
Handles new customer intake and pilot number activation.

| Function | Type | Description |
|----------|------|-------------|
| `copyNewEntriesToTracker()` | Scheduled / Manual | Copies new form responses to TRACKER. Deduplicates by company name against both TRACKER and ARCHIVED to prevent re-adding terminated companies. |
| `onPilotNumberEdit(e)` | Installable trigger | When a pilot number is entered in col E, validates for duplicates, then auto-sets contract dates and 12 months of "paid" |
| `populate12MonthsPaid(sheet, rowNumber)` | Internal | Populates 12 monthly columns with "paid" + dropdown validation from the contract start date |
| `backfillMissingPaidStatus()` | Manual | Scans all TRACKER rows with a contract start date but no monthly data, and backfills with "paid". Rows that already have monthly data are skipped. |

> **Trigger setup**: In the Apps Script Triggers UI, create an installable **On edit** trigger pointing to `onPilotNumberEdit` (not the built-in `onEdit`). Using a non-reserved function name prevents double-firing.

**Workflow:**
1. Customer fills out Google Form
2. `copyNewEntriesToTracker()` copies Company Name, Email, Location to TRACKER with empty pilot/contract fields
   - Skips companies already present in TRACKER **or** ARCHIVED
3. Admin manually enters a Pilot Number in column E
4. `onPilotNumberEdit()` fires:
   - Checks for duplicate pilot numbers across all TRACKER rows — shows error modal and clears the cell if duplicate found
   - Sets Contract Start = **1st of the current month**
   - Sets Contract End = **last day of the 12th month** (e.g. 1 Feb 2026 → 31 Jan 2027)
5. `populate12MonthsPaid()` fills the next 12 month columns with "paid" and adds dropdown validation
   - Reads month headers from **row 2** (handles both string and Date-formatted header cells)

---

### 03 MonthHighlighter.gs — Visual Highlighting
Highlights the current month column, and colors rows/dates by renewal urgency.

| Function | Type | Description |
|----------|------|-------------|
| `highlightCurrentMonth()` | Scheduled / Manual | Clears all month column backgrounds, then highlights the current month column in cyan (#00FFFF) |
| `setupUrgencyFormatting()` | Editor | Rebuilds **conditional formatting** on the Contract End Date column (col H) so the color reflects months-to-expiry. **Self-updating — no need to re-run.** |
| `highlightRenewalUrgency()` | Manual / Menu | Applies **static** urgency colors to cols B–H. Must be re-run as dates change. |
| `clearRenewalHighlights()` | Manual / Menu | Clears the static urgency colors from cols B–H |
| `urgencyMonthsDiff_(today, endDate)` | Internal | Whole-month difference; positive if endDate is in the future |
| `columnLetter_(col)` | Internal | Converts a 1-based column number to its A1 letter (8 → H) |

- `highlightCurrentMonth()` reads month headers from **row 2** (handles both string and Date-formatted header cells) and clears backgrounds from column I (FIRST_MONTH = 9) across all rows

**Urgency color scheme** — shared by both functions, and aligned to the reminder stages:

| Months to expiry | Color | Hex | Reminder stage |
|---|---|---|---|
| 3 | 🟩 Light green | `#D9EAD3` | 1st Reminder |
| 2 | 🟨 Light yellow | `#FFF2CC` | 2nd Reminder |
| 1 | 🟧 Light orange | `#FCE5CD` | 3rd Reminder |
| 0 / past | 🟥 Light red | `#F4CCCC` | Last Reminder (final notice) |

Rows with status `Renewed` / `Not Renewing` / `Terminated` are left uncolored by `setupUrgencyFormatting()` — a decided contract isn't urgent. (`highlightRenewalUrgency()` additionally paints `Renewed` rows the same light green, since they're safe.)

> **Which to use:** `setupUrgencyFormatting()` is preferred. It computes the month difference **inside the rule formula** (`YEAR`/`MONTH` against `TODAY()`), so cells roll from green → yellow → orange → red on their own as time passes. `highlightRenewalUrgency()` writes static backgrounds and goes stale until re-run. If you use both, the static one can overwrite the conditional colors.

> **Note:** `setupUrgencyFormatting()` replaces any existing conditional format rules **on col H only** — rules targeting other columns are preserved. It supersedes the legacy rules that keyed off the old `Pending` status, which silently stopped working when the reminder stages replaced `Pending`.

---

### 04 RenewalReminders.gs — Email Reminder System
Sends automated renewal reminder emails at 3, 2, 1 and 0 months (expiry month) before contract end. **Runs automatically on the 1st of every month.** Tracks renewal status directly in TRACKER col F via an escalating stage ladder.

| Function | Type | Description |
|----------|------|-------------|
| `checkAndSendReminders()` | Manual / Menu | Menu entry point — runs the engine and shows a UI alert with the run summary |
| `monthlyRenewalReminders()` | **Scheduled trigger** | Automated entry point (1st of each month). Runs the engine under error alerting and emails a run summary. Installed via `setupAutoReminderTrigger()`. |
| `runRenewalReminders()` | Internal | The core engine — no UI, no side effects beyond sending. Returns a result object. Shared by both entry points above. |
| `sendReminderRunSummary_(summary)` | Internal | Emails a table of who was reminded to `CONFIG.REMINDER_SUMMARY_EMAIL`, so unattended runs leave a visible trace. Also surfaces **lapsed contracts** (see below). |
| `migrateReminderStages()` | Editor | **One-time migration.** Realigns col F reminder stages to each row's current months-to-expiry, **without sending any email**. Shows a preview and asks for confirmation before writing. Only touches rows that are blank, legacy `Pending`, or already at a reminder stage — never `Renew`/`Renewed`/`Not Renewing`/`Terminated`. |
| `setupRenewalStatusDropdown()` | Editor | Applies dropdown validation (`CONFIG.RENEWAL_STATUS_VALUES`) to all rows in TRACKER col F |
| `getMonthsDifference(date1, date2)` | Internal | Calculates the whole-month difference between two dates |
| `sendRenewalReminderEmail(companyName, email, pilotNumber, expiryDate, monthsLeft, worqLocation)` | Internal | Sends the renewal reminder email via `MailApp.sendEmail()`. CC's the location inbox and sets Reply-To to `it@worq.space`. Company name is normalised to Proper Case before use. |
| `sendRenewalConfirmationEmail(companyName, email, pilotNumber, newStartDate, newEndDate, worqLocation)` | Internal | Sends a thank-you confirmation email when a customer renews, including their new effective tenure. CC's the location inbox. Company name is normalised to Proper Case. |
| `sendTerminationEmail(companyName, email, pilotNumber, endDate, worqLocation)` | Internal | Sends a termination confirmation email when a customer chooses Not Renewing. CC's the location inbox. Company name is normalised to Proper Case. |
| `sendVendorNotificationEmail(renewals, terminations)` | Internal | Sends a **single batched HTML table email** to the AstriCloud vendor (Theinosha) listing all renewals and terminations processed in the sync run. Terminations appear first (red), renewals second (green). Subject is `Astricloud Landline Subscription Update - {Month Year}`. Recipients are configured in `CONFIG.VENDOR_EMAIL` and `CONFIG.VENDOR_CC`. |
| `toProperCase(str)` | Internal | Converts a company name string to Title Case (e.g. `"INCOMPLETENESS THEOREM SDN BHD"` → `"Incompleteness Theorem Sdn Bhd"`). Applied in all customer-facing and vendor email functions, and in the Restore Archived dialog. |
| `formatPilotNumber(pilotNumber)` | Internal | Restores the leading zero on pilot numbers that Google Sheets strips when stored as a numeric value (e.g. `327746340` → `0327746340`). |
| `getLocationEmail(worqLocation)` | Internal | Looks up the location-specific email from the **Addresses** sheet by matching col A (Site) to the customer's WORQ Location. Returns the email in col C, or `null` if no match found. |
| `getVendorEmail()` | Internal | Returns the vendor "To" address from the **Config** sheet (`VENDOR_EMAIL` key), falling back to `CONFIG.VENDOR_EMAIL` if the sheet or key is missing/blank. |
| `getVendorCc()` | Internal | Returns the vendor CC list from the **Config** sheet (`VENDOR_CC` key), falling back to `CONFIG.VENDOR_CC` if the sheet or key is missing/blank. |
| `getConfigValue(key)` | Internal | Generic Key/Value lookup in the **Config** sheet (col A = Key, col B = Value, row 1 = header). Returns the trimmed value string, or `null` if the sheet/key is missing or blank. Mirrors the `getLocationEmail()` lookup pattern. |
| `setupConfigSheet()` | Manual / Menu | Creates the **Config** sheet if missing and seeds it with the current `CONFIG.VENDOR_EMAIL` / `VENDOR_CC` values. Non-destructive — existing keys are never overwritten. |

**Reminder escalation (the core model):**

Each run computes `monthsUntilExpiry` and looks up the target stage in `CONFIG.REMINDER_STAGES` (3 → 1st, 2 → 2nd, 1 → 3rd, 0 → Last). It sends **only if the target stage is strictly stronger than the row's current stage**, per `REMINDER_STAGE_ORDER`.

This single rule gives three important properties:
- **Idempotent** — re-running in the same month is a no-op (months-left hasn't changed, so the target stage isn't stronger). Safe to run manually any time.
- **No downgrades** — a company at `3rd Reminder Sent` can never drop back to `1st`.
- **No duplicate sends** — each stage fires at most once per contract cycle.

**Skip logic in `runRenewalReminders()`:**
- Skips rows with no contract end date or no email
- Skips rows already decided — col F is `Renew`, `Renewed`, `Not Renewing`, or `Terminated`
- Skips rows outside the 3/2/1/0-month window (including contracts already past — negative months have no target stage)
- Skips rows whose current stage already ≥ the target stage
- After sending, sets col F to the target stage with dropdown validation

> **No auto-termination.** After the final notice (0 months), a non-responding customer stays at `Last Reminder Sent` — the system never auto-sends a termination email or auto-sets `Not Renewing`. Handling stragglers is a deliberate manual step.

**What happens after the final notice:** once the end date passes, months-to-expiry goes negative, no target stage exists, and the row is skipped by every subsequent run — no further emails, no status change. The company becomes a **lapsed contract** and is surfaced two ways:

1. **Automatically** — the monthly run summary email leads with a red "⚠️ N lapsed contracts need a decision" block, and the count appears in the subject line. This is the push mechanism; it needs no one to remember anything.
2. **On demand** — the **Find Lapsed Contracts** menu item.

Both use the same detection helper (`getLapsedContracts_()` in `09 HealthCheck.gs`), so they can never disagree.

**Email Details:**
- Format: HTML (`htmlBody`) — text wraps naturally at the reader's window width
- Sender display name: "WORQ Operations Team" (sent from the script owner's Google account)
- Reply-To: `it@worq.space` — customer replies route to the IT inbox regardless of who the script runs as
- CC: Location-specific email from the Addresses sheet (e.g. `ttdi@worq.space`) — only the matching location is CC'd; omitted if no match is found
- Reminder subject (3/2/1 months): `⏰ Virtual Landline Renewal Reminder - X Month(s) Until Expiry`
- **Final notice subject (0 months):** `⚠️ Virtual Landline Expiring This Month - Action Required` — distinct urgent wording noting the service will be deactivated after the expiry date
- Confirmation subject: `✅ Virtual Landline Renewal Confirmed - Thank You, {Company}!`
- Termination subject: `Virtual Landline Service Termination Confirmation - {Company}`

> **Note on Reply All:** When a customer clicks Reply All, their email client addresses the reply to `it@worq.space` (Reply-To) and CC's the location email. This is the expected behaviour for actual customer accounts. Testing from the same account that sent the email will not reflect Reply-To correctly due to a Gmail self-reply quirk.

---

### 05 RenewalSync.gs — Renewal Processing
Processes renewal decisions from TRACKER col F and updates contract dates and monthly statuses. Also handles Not Renewing customers.

| Function | Type | Description |
|----------|------|-------------|
| `syncRenewals()` | Scheduled / Manual | Reads TRACKER col F — processes `Renew` and `Not Renewing` rows. Also reads `worqLocation` (col C) to pass to email functions for CC routing. |
| `populate12MonthsFromDate(sheet, rowNumber, startDate)` | Internal | Populates 12 monthly columns from a given start date: first month = `renew`, subsequent months = `paid` |
| `extendMonthHeaders(sheet, upToDate)` | Internal | Auto-extends TRACKER month header columns (rows 1 & 2) up to the required date. Called automatically before populating months. No-ops if headers already cover the range. |
| `markMonthCell(sheet, rowNumber, targetDate, value)` | Internal | Sets a single month column cell to a given value by matching the month header |

**Renew workflow (col F = "Renew"):**
1. Reads current Contract End from col H
2. New Contract Start = **1st of the month after current end** (avoids day-overflow issues, e.g. Jan 31 + 1 month ≠ Feb 28)
3. New Contract End = current end date + 1 year
4. Updates col G and col H with new dates
5. **`extendMonthHeaders()`** — automatically adds new month columns if the new end date goes beyond the last existing header (e.g. renewing into 2028 adds Jan-2028 through the required month; year label added in row 1 at each new January)
6. Populates 12 months from new start: first month = `renew`, rest = `paid`
7. Sets col F to `Renewed`
8. Sends renewal confirmation email (if email is present)

**Not Renewing workflow (col F = "Not Renewing"):**
1. Marks the contract end month cell as `terminate`
2. Sets col F to `Terminated`
3. Sends termination confirmation email (if email is present)
4. Admin then runs `archiveTerminated()` manually to move the row to ARCHIVED

---

### 06 ArchiveTerminated.gs — Customer Archival
Moves terminated customers from TRACKER to ARCHIVED.

| Function | Type | Description |
|----------|------|-------------|
| `archiveTerminated()` | Manual | Scans monthly status columns (col I onwards) for `terminate`, copies columns B onwards to ARCHIVED (skipping col A), deletes from TRACKER. Processes bottom-to-top to avoid index shifting. |
| `removeArchivedDuplicatesFromTracker()` | Manual | Scans TRACKER for any company that already exists in ARCHIVED (case-insensitive) and removes it. Use as a one-time cleanup if companies were re-added before the duplicate-check fix. |

- Writes columns B onwards only — column A (SEQUENCE formula) is never overwritten in either sheet
- Once archived, `copyNewEntriesToTracker()` will never re-add the company

---

### 07 SortByContractDate.gs — TRACKER Sorting
Sorts TRACKER data rows by Contract Start Date, oldest to newest.

| Function | Type | Description |
|----------|------|-------------|
| `sortByContractStartDate()` | Manual | Sorts all TRACKER data rows (row 3+) by Contract Start Date ascending. Rows without a contract date are pushed to the bottom. Shows a summary alert on completion. |

- Both header rows (row 1 = year labels, row 2 = month labels) are preserved
- Writes columns B onwards only — column A (SEQUENCE formula) is left untouched
- Rows with no company name are treated as undated and moved to the bottom

---

### 08 RestoreArchived.gs — Restore Archived Customers
Restores archived customers back to TRACKER for last-minute renewals where the contract end date has lapsed but the company decides to renew.

| Function | Type | Description |
|----------|------|-------------|
| `showRestoreArchivedDialog()` | Manual | Opens an HTML modal dialog listing all companies in the ARCHIVED sheet. User selects one or more to restore. Shows an alert if ARCHIVED is empty. |
| `restoreArchivedCompanies(indices)` | Called from dialog | Moves selected rows from ARCHIVED back to TRACKER (columns B onwards, skipping col A), deletes them from ARCHIVED, then re-sorts TRACKER by Contract Start Date. Shows a confirmation alert listing restored companies. |
| `formatDateForDialog_(value)` | Internal | Formats a date value (Date object or string) to `dd MMM yyyy` for display in the HTML dialog. |
| `sortTrackerSilently_()` | Internal | Sorts TRACKER by Contract Start Date without showing an alert — used internally after restore so the caller controls the feedback shown to the user. |

**Restore workflow:**
1. Admin runs `[07] Restore Archived Customer` from the menu
2. HTML modal opens listing all archived companies with their Contract Start and End dates
3. Admin ticks one or more companies (checkbox per row; "Select all" header checkbox available)
4. Admin clicks **Restore Selected**
5. Selected rows are moved from ARCHIVED to TRACKER (columns B onwards, skipping col A)
6. TRACKER is automatically re-sorted by Contract Start Date (oldest → newest)
7. Confirmation alert shows which companies were restored

> After restoring, the admin should update the company's Renewal Status (col F) and run `[04] Sync Renewals` as needed to extend the contract dates and repopulate monthly statuses.

---

### RestoreArchivedDialog.html — Restore Dialog UI
HTML modal rendered by `showRestoreArchivedDialog()` via `HtmlService.createTemplateFromFile()`.

- Checkbox table listing all archived companies (Company Name, Contract Start, Contract End)
- "Select all" header checkbox with indeterminate state support
- Selected count shown in the footer
- **Restore Selected** button (disabled until at least one company is checked)
- Calls `google.script.run.restoreArchivedCompanies(indices)` on submit
- Uses `<?!= companies ?>` (unescaped scriptlet) to inject the JSON array — required to prevent HtmlService from HTML-encoding the double quotes

---

### 09 HealthCheck.gs — Lapsed Contract Detection
Finds customers whose contract ended with no renewal decision recorded — the end of the line for the reminder ladder, since it never auto-terminates.

| Function | Type | Description |
|----------|------|-------------|
| `getLapsedContracts_()` | Internal | Shared detection. Returns structured data for every TRACKER row whose Contract End Date has passed and whose col F is **not** `Renewed` / `Terminated` / `Not Renewing`. Sorted longest-lapsed first. |
| `findLapsedContracts()` | Manual / Menu | Shows the lapsed list in a UI alert (truncated to 20 entries) |

A row lapses when its end date passes without an admin decision — typically sitting at `Last Reminder Sent` after the customer ignored all four reminders, but also any row left blank or mid-ladder.

`getLapsedContracts_()` is the single source of truth, used by both `findLapsedContracts()` (on demand) and `sendReminderRunSummary_()` (the monthly push). Changing the lapse rule in one place changes it everywhere.

---

### 10 FormSubmitHandler.gs — Automatic Form Submit Processing
Handles new Google Form submissions automatically via an installable `onFormSubmit` trigger. On each submission it highlights the new row, detects duplicate emails, routes duplicates to a separate archive sheet, and notifies the AstriCloud vendor for genuine new signups.

| Function | Type | Description |
|----------|------|-------------|
| `onFormSubmit(e)` | Installable trigger | Thin wrapper — runs `handleFormSubmit_(e)` inside a try/catch that routes any error to `notifyError_()`, so a failed submission surfaces instead of failing silently. Re-throws so the error still appears in the execution log. |
| `handleFormSubmit_(e)` | Internal | Main handler: clears old yellow highlight → highlights new row → checks for duplicate email → moves duplicate to Archived Form Responses or sends vendor notification |
| `clearFormHighlights(formSheet)` | Internal | Scans all data rows in Form Responses 1 and removes yellow (`#FFFF00`) backgrounds only, leaving other cell colours untouched |
| `isEmailDuplicate(formSheet, email, currentRowNum)` | Internal | Returns `true` if the submitted email already exists in any prior row of Form Responses 1 (case-insensitive). Skips the header row and the current new row. |
| `moveToArchivedFormResponses(formSheet, rowNum, rowData)` | Internal | Appends the full row to the **Archived Form Responses** sheet (same column structure — direct copy), then deletes the row from Form Responses 1 to keep it clean. Creates the archive sheet automatically if it does not exist. |
| `sendNewSignupVendorEmail(companyName, email, worqLocation, timestamp)` | Internal | Sends an HTML new-signup notification to the AstriCloud vendor. **To:** `CONFIG.VENDOR_EMAIL`. **CC:** `CONFIG.VENDOR_CC` + outlet email from `getLocationEmail()` (appended if found). **Subject:** `New Virtual Landline Signup Request - {Company Name}`. **ReplyTo:** `it@worq.space`. |
| `setupFormSubmitTrigger()` | Manual (editor) | Creates an installable `onFormSubmit` trigger for the spreadsheet. Removes any existing trigger with the same handler first (idempotent). |
| `removeFormSubmitTrigger()` | Manual (editor) | Removes all installable `onFormSubmit` triggers pointing to `onFormSubmit`. |

**Duplicate check logic:**
- Email from Col B (FORM_COLS.EMAIL) is normalised with `.toLowerCase().trim()`
- All rows in Form Responses 1 are scanned except the header (row 1) and the current new row
- If any prior row has the same email → duplicate → row is moved to Archived Form Responses; no vendor email is sent

**Yellow highlight logic:**
- On every form submit: all yellow (`#FFFF00`) backgrounds are cleared first, then the new submission row is highlighted yellow
- Ensures only one row is ever highlighted at a time

**Vendor email CC:**
- Base: `CONFIG.VENDOR_CC` (full AstriCloud + WORQ recipient list)
- The outlet email from the Addresses sheet is appended if the submitted WORQ Location matches a row in col A

> **Trigger setup**: Run `setupFormSubmitTrigger()` once from the Apps Script editor (or uncomment its menu item) to install the trigger. Verify in Apps Script > Triggers that an `onFormSubmit` trigger pointing to `onFormSubmit` is present.

---

## Automation & Reliability

### Installed triggers

| Trigger | Handler | Schedule | Installed by |
|---------|---------|----------|--------------|
| Time-based (CLOCK) | `monthlyRenewalReminders` | **1st of every month, ~8 AM** | `setupAutoReminderTrigger()` |
| On form submit | `onFormSubmit` | On each new form response | `setupFormSubmitTrigger()` |

Both setup functions are **idempotent** — they remove any existing trigger with the same handler before creating a new one, so they can't double-fire. `setupAutoReminderTrigger()` also cleans up the legacy daily `checkAndSendReminders` trigger it replaced.

> Everything else (`copyNewEntriesToTracker`, `syncRenewals`, `archiveTerminated`, …) is **manual by design**. Those mutate contracts or archive customers based on admin-entered decisions, so they keep a human gate.

### Failure alerting

Time-based and form triggers run unattended — without alerting, a failure is invisible until someone notices a customer never got an email. Both automated entry points are wrapped:

| Entry point | On error |
|-------------|----------|
| `monthlyRenewalReminders()` | `notifyError_('monthlyRenewalReminders', e)` → alert email, then re-throws |
| `onFormSubmit(e)` | `notifyError_('onFormSubmit', e)` → alert email, then re-throws |

`notifyError_()` emails `CONFIG.ALERT_EMAIL` (`it@worq.space`) with the function name, timestamp, error message, and stack trace. Errors are re-thrown so they still land in the Apps Script execution log.

The monthly run additionally emails a **success summary** (`sendReminderRunSummary_()`) listing every company reminded and its new stage — so a quiet month is distinguishable from a broken job.

### Trigger health

Run **`checkTriggerHealth()`** from the Apps Script editor periodically. It lists the installed triggers and warns if either expected trigger is missing.

> **Why this matters:** triggers can be silently dropped by a script copy, an owner change, or a manual deletion — with no warning. A missing `monthlyRenewalReminders` means **no renewal reminders go out at all**; a missing `onFormSubmit` means **new signups never reach the vendor**. This check is the only thing that surfaces either condition.

### Known gaps

- **Email typos fail silently.** `MailApp.sendEmail` does not throw on a syntactically valid but wrong address (e.g. a misspelled recipient in `VENDOR_CC`) — the mail is sent to the valid recipients and the bounce returns to the sending account, not the script. Blank values are safe (they fall back to `CONFIG`); only outright malformed strings throw. The **Config** sheet is protected to limit who can introduce a typo, but an authorized editor can still make one.
- **No auto-termination.** Deliberate — see the reminder section.

---

## Data Flow Diagram

```
Google Form
    |
    v
[Form Responses 1]  ←── onFormSubmit() fires automatically on each submission
    |   |                 1. Clears old yellow highlight
    |   |                 2. Highlights new row yellow
    |   |                 3. Checks email for duplicates in prior rows
    |   |
    |   |── duplicate email?
    |   |       YES → move row to [Archived Form Responses] (no vendor email)
    |   |       NO  → sendNewSignupVendorEmail() to AstriCloud
    |   |                 (CC: VENDOR_CC + outlet email)
    |
    | copyNewEntriesToTracker()  (manual, [01] menu)
    | (skips companies already in TRACKER or ARCHIVED)
    v
[TRACKER]
    |--- onPilotNumberEdit()  ← pilot number entry triggers auto-population
    |
    |--- monthlyRenewalReminders()  ← AUTOMATIC, 1st of every month (~8 AM)
    |       |   escalates each awaiting-reply customer by months-to-expiry:
    |       |     3 months → "1st Reminder Sent"   (countdown email)
    |       |     2 months → "2nd Reminder Sent"   (countdown email)
    |       |     1 month  → "3rd Reminder Sent"   (countdown email)
    |       |     0 months → "Last Reminder Sent"  (FINAL NOTICE email)
    |       └── emails a run summary to it@worq.space
    |           (errors → notifyError_() alerts it@worq.space)
    |
    |    Admin updates col F: "Renew" or "Not Renewing"
    |    (no auto-termination — stragglers stay at "Last Reminder Sent")
    |
    |--- syncRenewals()
    |       |── Renew      → extend contract, populate months, col F → "Renewed", send confirmation email
    |       └── Not Renewing → mark month as "terminate", col F → "Terminated", send termination email
    |
    |--- archiveTerminated() ──→ [ARCHIVED]
    |       (manual step after syncRenewals sets "Terminated")
    |                               |
    |                               | showRestoreArchivedDialog()
    |                               | (last-minute renewal — company decides to renew after lapse)
    |                               |
    |◄──────────────────────────────┘
    |    restoreArchivedCompanies() moves row back, re-sorts TRACKER
    |    Admin then updates col F and runs syncRenewals() to extend contract
    |
    |--- sortByContractStartDate() → re-sorts TRACKER rows oldest → newest
    |
    |--- highlightCurrentMonth() → highlights current month column in cyan
    |
    |--- backfillMissingPaidStatus() → fills historical rows with no monthly data
```

---

## Renewal Status Values (TRACKER col F)

| Value | Set by | Meaning |
|-------|--------|---------|
| *(empty)* | — | New entry, no action needed yet |
| `1st Reminder Sent` | Monthly reminder job | Reminder sent at 3 months to expiry — awaiting customer reply |
| `2nd Reminder Sent` | Monthly reminder job | Reminder sent at 2 months to expiry — awaiting customer reply |
| `3rd Reminder Sent` | Monthly reminder job | Reminder sent at 1 month to expiry — awaiting customer reply |
| `Last Reminder Sent` | Monthly reminder job | Final notice sent in the expiry month — awaiting customer reply |
| `Renew` | Admin (manual) | Customer confirmed renewal |
| `Renewed` | `syncRenewals()` | Contract extended, months populated |
| `Not Renewing` | Admin (manual) | Customer confirmed termination |
| `Terminated` | `syncRenewals()` | Termination processed, ready to archive |

> The four reminder stages all mean *"awaiting customer reply"* — they're a record of how far the escalation has gone. Only `Renew` and `Not Renewing` (set by the admin) move a company out of the reminder ladder and into `syncRenewals()`. A company sitting at `Last Reminder Sent` past its end date is a **lapsed contract** and appears in **Find Lapsed Contracts**.
>
> `Pending` is the **legacy** value from the pre-escalation system and is no longer used or valid. Run `migrateReminderStages()` to convert any remaining `Pending` rows.

## Monthly Status Values (TRACKER col I onwards)

| Value | Meaning |
|-------|---------|
| `paid` | Customer has paid for this month |
| `renew` | First month of a renewal cycle |
| `terminate` | Customer is terminating (triggers archival) |
| `not proceed` | Customer decided not to proceed |

---

## clasp Setup

Files are synced to Google Apps Script via [clasp](https://github.com/google/clasp).

```bash
# Push local changes to Apps Script
clasp push --force

# Pull latest from Apps Script
clasp pull
```

**Requirements:**
1. Apps Script API must be enabled at [script.google.com/home/usersettings](https://script.google.com/home/usersettings)
2. Run `clasp login` once to authenticate

`.claspignore` excludes: `.clasp.json`, `.claspignore`, `*.md`, `*.claude`, `*.js`

> **Note:** `clasp pull` downloads files with a `.js` extension, creating duplicates alongside the `.gs` files. To avoid this, always make changes locally in `.gs` files and use `clasp push --force` to sync up. If you do run `clasp pull`, manually copy the content from the `.js` files into the corresponding `.gs` files and delete the `.js` files. The `*.js` entry in `.claspignore` prevents any `.js` files from being accidentally pushed back.
