# Ailo — Alvys TMS connection

Pulls Loads and Invoices data out of [Alvys](https://alvys.com) (via their
public API) and writes it to a multi-sheet Excel report:

- **Summary** — headline KPIs for the date range (load/invoice counts, total
  invoiced revenue, total accessorial charges).
- **Loads** — load/shipment activity.
- **Invoices** — financial/billing data.
- **Accessorials** — invoice line items classified as accessorial charges
  (detention, lumper, stop-off, TONU, fuel surcharge, etc.), flattened out of
  the invoice records.

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
- The **Accessorials** sheet in `generate_report.py` is derived from each
  invoice's charge line items. Since the exact field name Alvys uses for
  that array, and the exact charge-type values, aren't confirmed yet, the
  code checks several plausible key names (`lineItems`, `charges`,
  `chargeLines`, `invoiceLines`) and classifies a line item as accessorial
  by matching keywords (detention, lumper, stop off, TONU, fuel surcharge,
  etc.) against its `type`/`chargeType`/`category`/`description` fields —
  see `LINE_ITEM_KEYS` and `ACCESSORIAL_KEYWORDS` at the top of
  `generate_report.py`. **Run it against a real invoice response and adjust
  both lists to match Alvys's actual field/value names** — if none of the
  keys match, the Accessorials sheet will just come back empty rather than
  erroring.
- Currently the Alvys public API is read-only, so this only pulls data;
  it doesn't write anything back to Alvys.
