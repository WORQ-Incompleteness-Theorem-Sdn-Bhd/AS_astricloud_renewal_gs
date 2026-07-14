# astricloud_renewal_gs

Google Apps Script automation for WORQ's AstriCloud (virtual landline) service — manages the full customer contract lifecycle: intake, payment tracking, renewal reminders, renewal processing, archival of terminated customers, and restoration of archived customers.

## Features

| # | Menu Item | Description |
|---|-----------|-------------|
| [01] | Copy New Entries from Form | Copies new Google Form responses to the TRACKER sheet (deduplicates against TRACKER and ARCHIVED) |
| [02] | Check & Send Renewal Reminders | Sends renewal reminder emails at 3, 2, 1 month(s) and 0 months before contract expiry |
| [03] | Backfill Missing Paid Status | Fills historical rows that have a contract date but no monthly payment data |
| [04] | Sync Renewals from Renewal Status | Extends contracts for "Renew" rows and marks "Not Renewing" rows as terminated |
| [05] | Archive Terminated Customers | Moves rows with `terminate` status from TRACKER to ARCHIVED sheet |
| [06] | Restore Archived Customer | Opens a dialog to select and move archived companies back to TRACKER (for last-minute renewals) |

## Sheets

| Sheet | Purpose |
|-------|---------|
| TRACKER | Main customer tracking grid with monthly payment status columns |
| Form Responses 1 | Auto-populated by the linked Google Form |
| ARCHIVED | Terminated customer records; can be restored back to TRACKER via [06] |
| Archived Form Responses | Duplicate form submissions (same email re-submitted); moved here automatically by the form submit trigger |
| Addresses | WORQ location → email lookup table; used to CC the correct location inbox on customer emails |

## File Structure

```
01 code.gs                  — CONFIG object, onOpen() menu
02 FormToTracker.gs         — Form intake, pilot number activation, backfill
03 MonthHighlighter.gs      — Highlight current month column
04 RenewalReminders.gs      — Renewal reminder and confirmation emails
05 RenewalSync.gs           — Contract extension and termination processing
06 ArchiveTerminated.gs     — Archive terminated customers / cleanup duplicates
07 SortByContractDate.gs    — Sort TRACKER by contract start date
08 RestoreArchived.gs       — Restore archived customers back to TRACKER
09 HealthCheck.gs           — Health check utilities
10 FormSubmitHandler.gs     — Form submit trigger: highlight, duplicate check, vendor email
RestoreArchivedDialog.html  — HTML modal UI for the restore dialog
```

## Setup

1. Enable the Apps Script API at [script.google.com/home/usersettings](https://script.google.com/home/usersettings)
2. Install [clasp](https://github.com/google/clasp): `npm install -g @google/clasp`
3. Authenticate: `clasp login`
4. Push changes: `clasp push`

Install the following triggers. Most have a setup function — run it from the Apps Script editor (select the function, click **Run**) rather than wiring it by hand:

- **Time-based** → `monthlyRenewalReminders` — sends escalating renewal reminders on the 1st of every month. Run `setupAutoReminderTrigger()`.
- **On form submit** → `onFormSubmit` — highlights new submissions, detects duplicates, sends vendor email. Run `setupFormSubmitTrigger()`.
- **On edit** → `onPilotNumberEdit` — auto-populates contract dates when a pilot number is entered. Create manually in the Apps Script Triggers UI (use this non-reserved name, not `onEdit`, to prevent double-firing).

Then run `checkTriggerHealth()` to confirm the automation is live — and re-run it occasionally, since a silently dropped trigger stops reminders or vendor emails with no other warning.

First-time setup also needs: `setupConfigSheet()` (vendor email recipients), `setupRenewalStatusDropdown()` (col F dropdown), and `setupUrgencyFormatting()` (Contract End Date colors).

## Documentation

See [DOCUMENTATION.md](DOCUMENTATION.md) for the full function reference, sheet structure, data flow diagram, and renewal status values.
