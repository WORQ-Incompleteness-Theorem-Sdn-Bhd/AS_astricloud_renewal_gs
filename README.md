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
| [06] | Sort by Contract Start Date | Re-sorts TRACKER rows oldest → newest by contract start date |
| [07] | Restore Archived Customer | Opens a dialog to select and move archived companies back to TRACKER (for last-minute renewals) |

## Sheets

| Sheet | Purpose |
|-------|---------|
| TRACKER | Main customer tracking grid with monthly payment status columns |
| Form Responses 1 | Auto-populated by the linked Google Form |
| ARCHIVED | Terminated customer records; can be restored back to TRACKER via [07] |
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
RestoreArchivedDialog.html  — HTML modal UI for the restore dialog
```

## Setup

1. Enable the Apps Script API at [script.google.com/home/usersettings](https://script.google.com/home/usersettings)
2. Install [clasp](https://github.com/google/clasp): `npm install -g @google/clasp`
3. Authenticate: `clasp login`
4. Push changes: `clasp push`

In the Apps Script Triggers UI, create an installable **On edit** trigger pointing to `onPilotNumberEdit` to enable auto-population of contract dates when a pilot number is entered.

## Documentation

See [DOCUMENTATION.md](DOCUMENTATION.md) for the full function reference, sheet structure, data flow diagram, and renewal status values.
