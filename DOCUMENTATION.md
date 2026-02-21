# AstriCloud Tracker Automation System

## Overview

A Google Apps Script automation system that manages customer contract tracking for WORQ's virtual landline (AstriCloud) service. The system automates the lifecycle of customer contracts: intake from a Google Form, payment tracking on a monthly grid, renewal reminders via email, renewal processing, and archival of terminated customers.

## Google Workspace Links

| Resource | Testing | Production |
|----------|---------|------------|
| Google Sheet | [Testing Sheet](https://docs.google.com/spreadsheets/d/1USIQXESClxQ_DHVD7qSLtAh8I1JbJUtLSg3KVDNlbn0/edit?gid=219063083#gid=219063083) | [Production Sheet](https://docs.google.com/spreadsheets/d/1t_-C-TZjd7dN6uweYG3wdZkToVAG4pxfEKInIG-TolE/edit?usp=drive_link) |
| Google Form | [Testing Form](https://docs.google.com/forms/d/1rC0suleIU9AqaQWaM3uPkZs7KlLo1O9wWugOgYYrYvc/edit) | [Production Form](https://docs.google.com/forms/d/1UAEJlLobYMUNB5JJFadbjHIEs1TmA90k6-1plIgERhs/preview) |

---

## Sheet Structure

The spreadsheet contains 4 sheets:

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
| F | Contract Start | Auto-set to 1st of the current month when pilot number is entered |
| G | Contract End | Auto-set to last day of the 12th month from contract start |
| H+ | Monthly columns (Feb-2024, Mar-2024, ...) | Payment status: `paid`, `renew`, `terminate`, `not proceed` |

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

### 3. Renewal Status
Tracks companies approaching contract expiry.

| Column | Field | Description |
|--------|-------|-------------|
| A | Companies | Company name |
| B | Outlet | WORQ location |
| C | Current Start | Current contract start date |
| D | Current End | Current contract end date |
| E | Customer Status | `Pending` or `Renewed` |
| F | New Renewal Date (EOM) | New renewal date if renewing |
| G | Final Status | `Pending`, `Renew`, or `Terminate` |

### 4. ARCHIVED
Storage for terminated customer records. Same column structure as TRACKER.

---

## File Structure & Functions

### 01 code.gs - Configuration & Triggers
The main entry point. Contains the global `CONFIG` object (declared with `var` for cross-file accessibility in Apps Script V8) and menu/trigger management.

**CONFIG.TRACKER_COLS:**
| Key | Column | Description |
|-----|--------|-------------|
| NO | A (1) | Row number |
| COMPANY_NAME | B (2) | Company name |
| WORQ_LOCATION | C (3) | WORQ outlet |
| COMPANY_EMAIL | D (4) | Contact email |
| PILOT_NUMBER | E (5) | Virtual landline number |
| CONTRACT_START | F (6) | Contract start date |
| CONTRACT_END | G (7) | Contract end date |
| FIRST_MONTH | H (8) | First month column (Feb-2024) |

| Function | Type | Description |
|----------|------|-------------|
| `onOpen()` | Auto-trigger | Creates the custom menu when the spreadsheet is opened |
| `setupTriggers()` | Manual | Installs 4 daily time-based triggers (7 AM, 8 AM, 9 AM, 10 AM) |
| `removeAllTriggers()` | Manual | Deletes all project triggers |

**Trigger Schedule:**
- 7:00 AM - `highlightCurrentMonth()`
- 8:00 AM - `copyNewEntriesToTracker()`
- 9:00 AM - `checkAndSendReminders()`
- 10:00 AM - `syncRenewals()`

### 02 FormToTracker.gs - Form Data Processing
Handles new customer intake and pilot number activation.

| Function | Type | Description |
|----------|------|-------------|
| `copyNewEntriesToTracker()` | Scheduled / Manual | Copies new form responses to TRACKER (deduplicates by company name) |
| `onPilotNumberEdit(e)` | Installable trigger | When a pilot number is entered in column E, validates for duplicates, then auto-sets contract dates and 12 months of "paid" |
| `populate12MonthsPaid(sheet, rowNumber)` | Internal | Populates 12 monthly columns with "paid" + dropdown validation from the contract start date |
| `debugHeaders()` | Debug helper | Logs the type and value of month column headers from row 2 to the execution log |

> **Trigger setup**: In the Apps Script Triggers UI, create an installable **On edit** trigger pointing to `onPilotNumberEdit` (not the built-in `onEdit`). Using a non-reserved function name prevents double-firing.

**Workflow:**
1. Customer fills out Google Form
2. `copyNewEntriesToTracker()` copies Company Name, Email, Location to TRACKER with empty pilot/contract fields
3. Admin manually enters a Pilot Number in column E
4. `onPilotNumberEdit()` fires:
   - Checks for duplicate pilot numbers across all TRACKER rows — shows error modal and clears the cell if duplicate found
   - Sets Contract Start = **1st of the current month**
   - Sets Contract End = **last day of the 12th month** (e.g. 1 Feb 2026 → 31 Jan 2027)
5. `populate12MonthsPaid()` fills the next 12 month columns with "paid" and adds dropdown validation
   - Reads month headers from **row 2** (handles both string and Date-formatted header cells)

### 03 MonthHighlighter.gs - Visual Highlighting
Highlights the current month column for easy visual reference.

| Function | Type | Description |
|----------|------|-------------|
| `highlightCurrentMonth()` | Scheduled / Manual | Clears all month column backgrounds, then highlights the current month column in cyan (#00FFFF) |

- Reads month headers from **row 2** (handles both string and Date-formatted header cells)
- Clears backgrounds starting from column H (FIRST_MONTH = 8) across all rows

### 04 RenewalReminders.gs - Email Reminder System
Sends automated renewal reminder emails at 3, 2, and 1 months before contract expiry.

| Function | Type | Description |
|----------|------|-------------|
| `checkAndSendReminders()` | Scheduled / Manual | Iterates TRACKER rows, checks months until expiry, sends emails for 3/2/1 month thresholds |
| `getMonthsDifference(date1, date2)` | Internal | Calculates month difference between two dates |
| `checkIfAlreadyRenewing(renewalSheet, companyName)` | Internal | Returns `true` if company already has `Renew` or `Pending` status in Renewal Status sheet |
| `sendRenewalReminderEmail(companyName, email, pilotNumber, expiryDate, monthsLeft)` | Internal | Sends the renewal reminder email via `MailApp.sendEmail()` |
| `addToRenewalStatus(renewalSheet, rowData, contractEnd)` | Internal | Adds a new row to Renewal Status with `Pending` status (deduplicates by company name) |

**Email Details:**
- Sender name: "WORQ IT Operations"
- Subject: "Virtual Landline Renewal Reminder - X Month(s) Until Expiry"
- Contains: company name, pilot number, expiry date, months remaining

### 05 RenewalSync.gs - Renewal Processing
Processes approved renewals and extends contracts.

| Function | Type | Description |
|----------|------|-------------|
| `syncRenewals()` | Scheduled / Manual | Finds rows in Renewal Status with Final Status = "Renew", extends the contract in TRACKER by 12 months, populates next 12 months as "paid", marks Customer Status as "Renewed" |
| `populate12MonthsFromDate(sheet, rowNumber, startDate)` | Internal | Same as `populate12MonthsPaid` but accepts an arbitrary start date |

**Workflow:**
1. Admin sets Final Status = "Renew" in Renewal Status sheet
2. `syncRenewals()` finds matching company in TRACKER
3. Contract End extended by +12 months from current end
4. Next 12 month columns populated with "paid" + dropdown validation
5. Customer Status in Renewal Status updated to "Renewed"

### 06 ArchiveTerminated.gs - Customer Archival
Moves terminated customers from TRACKER to ARCHIVED.

| Function | Type | Description |
|----------|------|-------------|
| `archiveTerminated()` | Manual | Scans monthly status columns for "terminate", copies full row to ARCHIVED sheet, deletes from TRACKER. Processes bottom-to-top to avoid index shifting. |

---

## Data Flow Diagram

```
Google Form
    |
    v
[Form Responses 1]
    |
    | copyNewEntriesToTracker()
    v
[TRACKER] <--- onPilotNumberEdit() (pilot number triggers auto-population)
    |
    |--- checkAndSendReminders() ---> Email to customer
    |                              |
    |                              v
    |                        [Renewal Status]
    |                              |
    |--- syncRenewals() <---------+  (when Final Status = "Renew")
    |
    |--- archiveTerminated() ----> [ARCHIVED]
```

---

## Known Issues / Potential Bugs

1. **`CONFIG.COLS.*` references in files 05 and 06**: `05 RenewalSync.gs` and `06 ArchiveTerminated.gs` may reference `CONFIG.COLS.*` instead of `CONFIG.TRACKER_COLS.*`, which will cause `undefined` errors at runtime.
   - `05 RenewalSync.gs:31,36,44` - `CONFIG.COLS.COMPANY_NAME`, `CONFIG.COLS.CONTRACT_END` should be `CONFIG.TRACKER_COLS.*`
   - `06 ArchiveTerminated.gs:24,38` - `CONFIG.COLS.FIRST_MONTH`, `CONFIG.COLS.COMPANY_NAME` should be `CONFIG.TRACKER_COLS.*`

2. **`copyNewEntriesToTracker()` row numbering**: The `lastNo` calculation uses the last row of `trackerData` (fetched once before the loop). If multiple new entries are copied in one run, all will get the same `newNo` because `trackerData` is not refreshed after each `appendRow()`.

---

## Monthly Status Values

The dropdown validation in monthly columns allows these values:
- `paid` - Customer has paid for this month
- `renew` - Customer intends to renew
- `terminate` - Customer is terminating (triggers archival)
- `not proceed` - Customer decided not to proceed
