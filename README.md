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
