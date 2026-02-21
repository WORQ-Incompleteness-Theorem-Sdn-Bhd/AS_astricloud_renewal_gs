# AstriCloud Tracker Automation System

## Overview

A Google Apps Script automation system that manages customer contract tracking for WORQ's virtual landline (AstriCloud) service. The system automates the full lifecycle of customer contracts: intake from a Google Form, payment tracking on a monthly grid, renewal reminders via email, renewal processing, and archival of terminated customers.

## Google Workspace Links

| Resource | Testing | Production |
|----------|---------|------------|
| Google Sheet | [Testing Sheet](https://docs.google.com/spreadsheets/d/1USIQXESClxQ_DHVD7qSLtAh8I1JbJUtLSg3KVDNlbn0/edit?gid=219063083#gid=219063083) | [Production Sheet](https://docs.google.com/spreadsheets/d/1t_-C-TZjd7dN6uweYG3wdZkToVAG4pxfEKInIG-TolE/edit?usp=drive_link) |
| Google Form | [Testing Form](https://docs.google.com/forms/d/1rC0suleIU9AqaQWaM3uPkZs7KlLo1O9wWugOgYYrYvc/edit) | [Production Form](https://docs.google.com/forms/d/1UAEJlLobYMUNB5JJFadbjHIEs1TmA90k6-1plIgERhs/preview) |

---

## Sheet Structure

The spreadsheet contains 3 active sheets:

### 1. TRACKER (Main sheet)
The central tracking sheet with a monthly payment grid. It has **two header rows**:
- **Row 1**: Year group labels (2024, 2025, 2026, ...) spanning the month columns
- **Row 2**: Month column names (Feb-2024, Mar-2024, ...)
- **Row 3+**: Customer data

| Column | Field | Description |
|--------|-------|-------------|
| A | NO | Auto-incremented row number |
| B | Company Name | Customer company name |
| C | WORQ Location | Which WORQ outlet the customer belongs to |
| D | Company Email | Contact email for the customer |
| E | Pilot Number | Virtual landline number (triggers auto-population when entered) |
| F | Renewal Status | Tracks renewal lifecycle: `Pending`, `Renew`, `Renewed`, `Not Renewing`, `Terminated` |
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
Storage for terminated customer records. Same column structure as TRACKER. Terminated companies are never re-added to TRACKER by the copy function.

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

**CONFIG.REMINDER_MONTHS:** `[3, 2, 1, 0]` — sends reminders at 3, 2, 1 months before expiry, and again in the expiry month itself.

| Function | Type | Description |
|----------|------|-------------|
| `onOpen()` | Auto-trigger | Creates the custom menu when the spreadsheet is opened |

**Menu items (in order):**
| # | Label | Function |
|---|-------|----------|
| [01] | Highlight Current Month | `highlightCurrentMonth()` |
| [02] | Copy New Entries from Form | `copyNewEntriesToTracker()` |
| [03] | Check & Send Renewal Reminders | `checkAndSendReminders()` |
| [04] | Backfill Missing Paid Status | `backfillMissingPaidStatus()` |
| [05] | Sync Renewals from Renewal Status | `syncRenewals()` |
| [06] | Archive Terminated Customers | `archiveTerminated()` |
| — | *(separator)* | — |
| | Setup Renewal Status Dropdown | `setupRenewalStatusDropdown()` |
| | Remove Archived Duplicates from Tracker | `removeArchivedDuplicatesFromTracker()` |

> `setupTriggers()` and `removeAllTriggers()` are currently commented out. Triggers are managed manually via the Apps Script Triggers UI.

---

### 02 FormToTracker.gs — Form Data Processing
Handles new customer intake and pilot number activation.

| Function | Type | Description |
|----------|------|-------------|
| `copyNewEntriesToTracker()` | Scheduled / Manual | Copies new form responses to TRACKER. Deduplicates by company name against both TRACKER and ARCHIVED to prevent re-adding terminated companies. |
| `onPilotNumberEdit(e)` | Installable trigger | When a pilot number is entered in col E, validates for duplicates, then auto-sets contract dates and 12 months of "paid" |
| `populate12MonthsPaid(sheet, rowNumber)` | Internal | Populates 12 monthly columns with "paid" + dropdown validation from the contract start date |
| `backfillMissingPaidStatus()` | Manual | Scans all TRACKER rows with a contract start date but no monthly data, and backfills with "paid". Rows that already have monthly data are skipped. |
| `debugHeaders()` | Debug helper | Logs the type and value of month column headers from row 2 to the execution log |

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
Highlights the current month column for easy visual reference.

| Function | Type | Description |
|----------|------|-------------|
| `highlightCurrentMonth()` | Scheduled / Manual | Clears all month column backgrounds, then highlights the current month column in cyan (#00FFFF) |

- Reads month headers from **row 2** (handles both string and Date-formatted header cells)
- Clears backgrounds starting from column I (FIRST_MONTH = 9) across all rows

---

### 04 RenewalReminders.gs — Email Reminder System
Sends automated renewal reminder emails at 3, 2, 1 month(s) and 0 months (expiry month) before contract end. Tracks renewal status directly in TRACKER col F.

| Function | Type | Description |
|----------|------|-------------|
| `checkAndSendReminders()` | Scheduled / Manual | Iterates TRACKER rows, checks months until expiry, sends reminder emails at configured thresholds |
| `setupRenewalStatusDropdown()` | Manual | Applies dropdown validation (`Pending`, `Renew`, `Renewed`, `Not Renewing`) to all rows in TRACKER col F |
| `getMonthsDifference(date1, date2)` | Internal | Calculates the whole-month difference between two dates |
| `sendRenewalReminderEmail(companyName, email, pilotNumber, expiryDate, monthsLeft)` | Internal | Sends the renewal reminder email via `MailApp.sendEmail()` |
| `sendRenewalConfirmationEmail(companyName, email, pilotNumber, newStartDate, newEndDate)` | Internal | Sends a thank-you confirmation email when a customer renews, including their new effective tenure |
| `sendTerminationEmail(companyName, email, pilotNumber, endDate)` | Internal | Sends a termination confirmation email when a customer chooses Not Renewing |

**Skip logic in `checkAndSendReminders()`:**
- Skips rows with no contract end date or no email
- Skips rows where col F is already `Renew` or `Pending` (reminder already sent / customer already confirmed)
- Skips rows where contract end is in the past
- After sending a reminder, sets col F to `Pending` with dropdown validation

**Email Details:**
- Sender: `it_worq@worq.space` (display name: "WORQ IT Operations")
- Reminder subject: `Virtual Landline Renewal Reminder - X Month(s) Until Expiry`
- Confirmation subject: `Virtual Landline Renewal Confirmed`
- Termination subject: `Virtual Landline Service Termination Confirmed`

---

### 05 RenewalSync.gs — Renewal Processing
Processes renewal decisions from TRACKER col F and updates contract dates and monthly statuses. Also handles Not Renewing customers.

| Function | Type | Description |
|----------|------|-------------|
| `syncRenewals()` | Scheduled / Manual | Reads TRACKER col F — processes `Renew` and `Not Renewing` rows |
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
| `archiveTerminated()` | Manual | Scans monthly status columns (col I onwards) for `terminate`, copies the full row to ARCHIVED, deletes from TRACKER. Processes bottom-to-top to avoid index shifting. |
| `removeArchivedDuplicatesFromTracker()` | Manual | Scans TRACKER for any company that already exists in ARCHIVED (case-insensitive) and removes it. Use as a one-time cleanup if companies were re-added before the duplicate-check fix. |

- Preserves all data (entire row is copied as-is to ARCHIVED)
- Once archived, `copyNewEntriesToTracker()` will never re-add the company

---

## Data Flow Diagram

```
Google Form
    |
    v
[Form Responses 1]
    |
    | copyNewEntriesToTracker()
    | (skips companies already in TRACKER or ARCHIVED)
    v
[TRACKER]
    |--- onPilotNumberEdit()  ← pilot number entry triggers auto-population
    |
    |--- checkAndSendReminders() ──→ Email to customer (col F → "Pending")
    |
    |    Admin updates col F: "Renew" or "Not Renewing"
    |
    |--- syncRenewals()
    |       |── Renew      → extend contract, populate months, col F → "Renewed", send confirmation email
    |       └── Not Renewing → mark month as "terminate", col F → "Terminated", send termination email
    |
    |--- archiveTerminated() ──→ [ARCHIVED]
    |       (manual step after syncRenewals sets "Terminated")
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
| `Pending` | `checkAndSendReminders()` | Reminder sent, awaiting customer decision |
| `Renew` | Admin (manual) | Customer confirmed renewal |
| `Renewed` | `syncRenewals()` | Contract extended, months populated |
| `Not Renewing` | Admin (manual) | Customer confirmed termination |
| `Terminated` | `syncRenewals()` | Termination processed, ready to archive |

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
