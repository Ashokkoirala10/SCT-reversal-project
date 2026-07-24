# IBFT &rarr; Reversal Generator (Django)

A small Django web app that takes the daily `ibft-transaction_*.xlsx` export and
produces a `need_to_reversal_*.xlsx` file, exactly like the manual process it
replaces:

- **failed** sheet: `FAILED` transactions, excluding "Insufficient funds"
- **coop** sheet: manual-reversal transactions for every aggregator other
  than IME REMIT / CITY REMIT, grouped by Member Name (alphabetical), each group's
  "S NO" restarting at 1
- **imeremit** sheet: manual-reversal transactions where Aggregator = `IME REMIT`
- **cityremit** sheet: manual-reversal transactions where Aggregator = `CITY REMIT`
- **timeout** sheet: any `TIMEOUT` transactions, unchanged

For every kept reversal row, three fields are (re)built:
- **Debit Account Number** — the fixed clearing account `0002335524115`,
  **except** when the original transaction's **Debtor Bank is Prabhu Bank**,
  in which case the dedicated account `99901170130555` is used instead
  (Prabhu Bank reversals are rare but need their own account).
- **Credit Account Number** — the *original* row's Debit Account Number (the
  money goes back to wherever it was originally pulled from)
- **narration** (in the duplicated "Member Transaction Id" column) —
  `REV` + the member transaction id with its leading zeros and leading
  prefix/letters stripped + `-` + Session Id
  (e.g. `REQ_1783249894525` &rarr; `REV1783249894525-100168`,
  `IME-00113275637-572` &rarr; `REV00113275637-572-100168`)

All id-like columns (Member Transaction Id, Network Reference Id, Session Id,
Debit/Credit Account Number) are written as **text**, not numbers, in the
output workbook. Long numeric ids (17+ digits) silently lose trailing
precision if Excel is allowed to store them as a number — this was traced
back as the likely cause of a past mismatch between a source file's Member
Transaction Id and what showed up in the generated reversal file (the
narration still matched because it was already being built as a string, not
a number). Forcing text format everywhere makes that class of bug
impossible going forward. If it ever happens again, please save the exact
source value + what showed up in the generated file, so the real cause can
be confirmed and fixed for good.

## Double-reversal prevention

Before writing a new reversal file, the app looks at the **most recently
generated file that has already been marked "passed"** (see below — not
just any generated file) and collects every Network Reference Id that
already appears in its `coop` / `imeremit` / `cityremit` sheets. Any row in
the new upload whose Network Reference Id is already in that set is skipped
(and counted as `duplicate_skipped` in the audit log / result page / day
reconciliation popup) instead of being reversed a second time.

## Login

There is no self-registration — two hardcoded accounts are seeded by a data
migration (`core/migrations/0003_seed_users.py`):

| Username        | Password    | Role                                             |
|-----------------|-------------|---------------------------------------------------|
| `ashok.koirala` | `ashok@123` | Regular operator — can upload/generate            |
| `admin`         | `admin@123` | Admin — also sees the central Audit Log & Dashboard |

Change these passwords (via `/admin/` or the Django shell) before using this
anywhere beyond your own machine — they're intentionally simple/hardcoded
per the current requirements.

## Passed workflow (review before it's shared)

A newly generated file only shows up in **your own "My activity" list** —
not the central shared list, and not the analytics dashboard — until you
(or an admin) review it and click **"Mark as passed"** (on the result page,
or right from the activity table). Once passed:

- it moves into the **"Shared passed reports"** central list, visible to
  everyone
- it's included in the analytics dashboard
- it becomes the baseline the *next* upload's double-reversal check compares
  against (see below)

Anyone can unmark their own passed report; admins can mark/unmark anyone's.

## UI

- **Generate page** (`/`): a 3-column layout — upload form on the left,
  the central **"Shared passed reports"** list in the middle (bigger, since
  it's the shared source of truth), and your own **"My activity"** on the
  right. Both activity tables are paginated 10 rows at a time, and file
  columns are icon-only download buttons (hover for the filename).
- **Audit Log** (`/audit-log/`, admin only): every file from every user,
  passed or not, with Prabhu-reroute / unrecognized-bank / duplicate-skip
  counts per row.
- **Dashboard** (`/dashboard/`, admin only): all-time totals (including
  amount + charge for success, failed, manual reversal, and system
  reversal) plus a day-by-day breakdown. **Click any day** to open a full
  reconciliation popup: manual reversal vs. system reversal vs. failed vs.
  success, each with count/amount/charge, the day's failure-reason
  breakdown, and the usual double-reversal / Prabhu-reroute / unrecognized-
  bank checks. Charge follows the same business rule as the generated file
  itself — NCHL-routed transactions carry their charge inside the amount,
  so their charge is treated as 0 everywhere charge is summed.
- "Manual reversal" vs. "system reversal": a `REVERSAL` row is a **manual**
  reversal if its Source Message says so (these are the ones written into
  the generated file); any other `REVERSAL` row is a **system** reversal
  (the switch already reversed it automatically) — not written to the
  file, but still counted for reconciliation.

## File naming

- Every uploaded source file is renamed/stored as `ibft_txn_data_<date>.xlsx`
  (date parsed from the original filename if it has one, otherwise today's
  date) — whatever it was originally called.
- Every generated file is `need_to_reversal_<date>.xlsx`. Downloads always
  use this clean name (via a dedicated download view), even though the
  file actually stored on disk may have a suffix Django added to avoid a
  naming collision.

## Setup

```bash
python -m venv venv
source venv/bin/activate        # venv\Scripts\activate on Windows
pip install -r requirements.txt

python manage.py migrate        # also seeds the two hardcoded logins
python manage.py runserver
```

Then open http://127.0.0.1:8000/, log in with either account above, upload
a file, and (as `admin`) check http://127.0.0.1:8000/dashboard/ and
http://127.0.0.1:8000/audit-log/.

## Project layout

```
reversal_project/
├── manage.py
├── requirements.txt
├── reversal_project/        # Django project settings/urls
└── core/                    # the one app
    ├── models.py            # ProcessingLog (audit trail)
    ├── forms.py              # upload form
    ├── services.py           # <-- all the business logic lives here
    ├── views.py
    ├── urls.py
    ├── admin.py
    └── templates/core/
        ├── base.html
        ├── login.html
        ├── upload.html
        ├── result.html
        ├── audit_log.html
        └── dashboard.html
```

## Notes / things worth knowing

- The app reads the source workbook's **"Transactions"** sheet (or the first
  sheet if that name isn't found) and auto-detects the header row, so a stray
  "Period: ..." banner row above the header (as in the sample file) is
  handled automatically.
- If the uploaded file is missing an expected column, the app shows a clear
  error on the upload page and records the failure in the audit log — it
  won't silently produce a bad file.
- `DATA_UPLOAD_MAX_MEMORY_SIZE` / `FILE_UPLOAD_MAX_MEMORY_SIZE` are raised to
  50 MB in `settings.py` since these exports can be large (the sample file
  was ~4.5 MB / ~29,000 rows). Raise further if needed.
- This ships with `DEBUG = True` and a placeholder `SECRET_KEY` — fine for
  running locally / on a trusted internal network, but change both (and set
  `DEBUG = False`, configure `ALLOWED_HOSTS` properly, rotate the two
  hardcoded passwords, and put uploaded/media files somewhere
  access-controlled) before exposing this beyond your own machine, since
  transaction files are sensitive.
- The "Prabhu Bank" check matches the **Debtor Bank** column case-
  insensitively containing "PRABHU", so "Prabhu Bank", "Prabhu Bank
  Limited", etc. all match.
- The double-reversal check only looks at the single most recent prior
  successful file, matching the "one step behind" comparison that was
  asked for. If you skip a day (or need to check further back), the
  `extract_reversal_network_reference_ids()` helper in `core/services.py`
  can be pointed at any older generated file too.
