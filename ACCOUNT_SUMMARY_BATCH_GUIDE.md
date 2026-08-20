# Overnight Account Summary Database

This project can now analyse every Excel workbook in
`C:\Users\admin\Desktop\bank_trails` one at a time through the existing
`app_account.py` account-summary code.

## Start everything

Double-click:

`START_OVERNIGHT_AND_DASHBOARD.bat`

It opens:

1. The resumable overnight worker.
2. The local dashboard at `http://127.0.0.1:5002`.

The worker prevents Windows system sleep while it is active. The monitor may
turn off normally. Keep the computer connected to power, and do not close a
laptop lid if Windows is configured to sleep when the lid closes. Do not shut
down or restart Windows during the run. If the worker is stopped, start the
launcher again and it will continue from the saved queue.

## What is saved

- SQLite database: `data\account_summaries.sqlite`
- Worker log: `data\account_summary_worker.log`

The database keeps:

- Account Wise Summary
- Bank Wise Summary
- Partial Bank Wise Summary
- Per-file status, duration, attempts, and error details
- Duplicate-removal audit counts for transactions, recovery sheets, summary
  rows, and exact copied workbooks

Completed files are skipped on later runs. New files are discovered
automatically. Changed files are recalculated and replaced in one database
transaction, so incomplete results are not shown.

The normal launcher uses a fast duplicate audit instead of recalculating all
29,000+ files. It reads only the ACK, credited account, and credited transaction
ID cells directly from each XLSX, caches that audit against the file
fingerprint, and queues only affected or unreadable workbooks. Queued workbooks
are analysed by four independent processes while SQLite writes remain atomic.
Clean existing summaries remain available throughout.

Large workbooks are also accelerated inside `app_account.py`: duplicate credit
keys are identified once for each ACK instead of rebuilding and scanning the
complete ACK once for every account. This removes the main quadratic bottleneck
from the previous multi-day run.

## Credit-only duplicate protection

The first matching credit contributes to **Total Credited Amount**. Every later
duplicate credit is excluded from that total only. Its source row remains
available for **Total Debited Amount**, **Found in Other Sheets**, recovery
breakdowns, status/flow calculations, and transaction lookup. A duplicate
requires exactly the same ACK, credited transaction ID, and last four
characters of the credited account number. All other columns are ignored for
this comparison. A row with a blank credited transaction ID or blank credited
account is retained because it cannot be matched safely.

The **Duplicate Entry Info** column identifies the credited transaction ID,
number of later credited amounts not counted, and the excluded credited amount.
The note is shown only on the affected credited account row and explicitly says
that debit and other-sheet matching remain counted.
Exact workbook copies are skipped by SHA-256, and SQLite unique indexes provide
a final guard against repeated summary rows.

The dashboard's **Duplicate credits excluded** counter shows the recorded
credit-only audit total. Hover over it to see legacy audit fields and exact
workbook-copy information.

To perform a one-time rebuild of all existing summaries, double-click:

`REPROCESS_ALL_STRICT_ONCE.bat`

Only one worker can use the database at a time.

## Dashboard

The dashboard opens with **All ACKs** selected. It supports:

- All ACKs or one acknowledgement number
- Account, bank, and partial-bank views
- Status and text filters
- Processing progress and failed-file details
- Excel download for all ACKs or the selected ACK

The dashboard binds only to `127.0.0.1`, so the financial data is not published
to the internet.

## Command-line options

Run one pass and stop:

```powershell
python batch_account_summaries.py --input "C:\Users\admin\Desktop\bank_trails"
```

Retry files that reached the failure limit:

```powershell
python batch_account_summaries.py --retry-failed
```

Recalculate every discovered file:

```powershell
python batch_account_summaries.py --reprocess-all
```

Recalculate only files where the previous report detected duplicate
transactions:

```powershell
python batch_account_summaries.py --reprocess-detected-duplicates
```

Fast-audit all files and recalculate only those matching the current duplicate
rule (recommended):

```powershell
python batch_account_summaries.py --fast-reprocess-duplicates --audit-workers 16 --process-workers 4
```

Run only the dashboard:

```powershell
python account_summary_dashboard.py
```
