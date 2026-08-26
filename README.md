# Ailo — Alvys TMS connection

Pulls Loads and Invoices data out of [Alvys](https://alvys.com) (via their
public API) and writes it to an Excel report.

## Setup

1. In Alvys: **Profile → API** to open the Public API page, and create an
   application to get a `client_id` and `client_secret`. Make sure the app
   is granted the scopes you need (e.g. `load:read`, `invoice:read`).
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Copy `.env.example` to `.env` and fill in your real credentials:
   ```
   cp .env.example .env
   ```
   `.env` is gitignored — your credentials never get committed.

## Usage

```
python generate_report.py --start-date 2026-07-01 --end-date 2026-07-31 --output july_report.xlsx
```

All arguments are optional; by default it pulls the last 30 days into
`alvys_report.xlsx` in the current directory.

## Statement invoicing

Some customers (e.g. Glovis, Jess Smith and Sons, Berg Mill) are invoiced by
vessel name or by booking/PO # instead of load-by-load. `statement_export.py`
pulls Alvys loads for those customers and writes a CSV shaped for the
**Statement Tracker** artifact, which groups loads into statements, lets you
save your external invoice number against each one, and tracks paid history:

```
python statement_export.py --start-date 2026-07-01 --end-date 2026-08-31 --output statement_loads.csv
```

Paste the resulting CSV (or upload the file) into the tracker's Import tab.
`--customers` accepts a comma-separated list of case-insensitive substrings
to match against the load's customer name (defaults to `glovis,jess
smith,berg mill`) if you need to add or change which customers it pulls.

## Notes / things to verify

- Auth uses Alvys's OAuth2 client-credentials flow
  (`POST https://auth.alvys.com/oauth/token`).
- The exact `/loads/search` and `/invoices/search` paths and query
  parameters in `alvys_client.py` are based on Alvys's public docs
  (docs.alvys.com) but haven't been verified against a live call yet —
  run a small test once real credentials are in place and adjust field
  names if the response shape differs.
- Currently the Alvys public API is read-only, so this only pulls data;
  it doesn't write anything back to Alvys.
- The field names `statement_export.py` looks for on each load record
  (`FIELD_CANDIDATES`) are also best-effort guesses — once you've made a
  live call, check the real field names and adjust the candidate lists if
  needed.
