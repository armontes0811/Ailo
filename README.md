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
to match against the load's customer name (defaults to `mobis,imports/
exports,jess smith`) if you need to add or change which customers it pulls.
These defaults are deliberately specific: Alvys has multiple divisions with
overlapping names (e.g. `GLOVIS America, Inc.` vs `Glovis (Mobis Division)
America, Inc`, or `Berg Mill Supply` vs `Berg Mill Supply
(Imports/Exports)`) — a looser match like `glovis` or `berg mill` would
pull in the wrong division's loads.

## Notes / things to verify

- Auth uses Alvys's OAuth2 client-credentials flow
  (`POST https://auth.alvys.com/oauth/token`).
- **Verified against a live account:** the search endpoints are
  `POST https://integrations.alvys.com/api/p/v1/loads/search` and
  `.../invoices/search` (JSON body: `page`, `pageSize`, `dateRange` /
  `invoicedDateRange`, `status`) — not the GET `/public/...` paths this
  originally assumed. Alvys requires at least one non-date filter
  (`status`, `PONumbers`, `CustomerId`, `LoadNumbers`, `OrderNumbers`, ...),
  so `alvys_client.py` defaults `status` to the full enum
  (`ALL_LOAD_STATUSES` / `ALL_INVOICE_STATUSES`) when a caller doesn't pass
  one. Responses are shaped `{Page, PageSize, Total, Items: [...]}`.
- On a load record, `VesselName` isn't a top-level field — it's an entry in
  `References` where `Name == "Vessel Name"`. Origin/destination come from
  `Stops` (`StopType == "Pickup"` / `"Delivery"`), and the billable amount
  is `CustomerRate.Amount`. See `statement_export.py` for the extraction
  logic.
- `generate_report.py` and `statement_export.py` are otherwise both
  single-page pulls / paginated pulls respectively — `generate_report.py`
  doesn't loop over pages yet, so long date ranges are capped at the first
  200 rows per endpoint.
- Currently the Alvys public API is read-only, so this only pulls data;
  it doesn't write anything back to Alvys.
