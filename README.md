# Ailo — Alvys TMS connection

Pulls Loads and Invoices data out of [Alvys](https://alvys.com) (via their
public API) and writes it to a multi-sheet Excel report:

- **Summary** — headline KPIs for the date range, including the unbilled
  totals below.
- **Unbilled Loads** — loads that have been delivered but are sitting in any
  status other than "Invoiced" or "Complete" (i.e. not yet billed to the
  customer or paid to the carrier), with days-since-delivery added. Loads
  that haven't been delivered yet (open/covered/dispatched/in transit) are
  excluded entirely.
- **Unbilled by Carrier** / **Unbilled by Customer** — the above, grouped
  with load counts, aging (avg/max days since delivery), and total rate
  value.
- **Loads** — full load/shipment activity (not filtered).
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
`alvys_report.xlsx` in the current directory. This also writes a JSON export
(`july_report.json` in the example above) with the same data, for the
dashboard below.

## Dashboard

`dashboard.html` is a self-contained visual dashboard focused on unbilled
loads: KPIs (unbilled count, avg/oldest days since delivery, unbilled value),
a breakdown by status, an aging chart (days since delivery, bucketed), and
breakdowns by carrier and customer — all filterable by status via the chips
above the KPIs (e.g. view only "Delivered" vs. only "POD Received"). Meant to
be published as a Claude Artifact; it reads report data from a
`#report-data` script tag embedded in the page, with a manual
file-load/drag-and-drop fallback for viewing an export ad hoc.

To refresh the published dashboard with new data:

```
python generate_report.py --start-date ... --end-date ...
python build_dashboard.py alvys_report.json
```

Then ask Claude to republish `dashboard.html` to the same Artifact URL. This
only works from a session that can actually reach Alvys's API (some sandboxed
environments block outbound access to `auth.alvys.com` / `api.alvys.com` by
policy) — "refresh the dashboard" is exactly this three-step loop asked of
Claude in chat.

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
- The **unbilled loads** logic assumes a single `status` field (or one of
  `STATUS_KEYS`' fallbacks) on each load, with pre-delivery values in
  `PRE_DELIVERY_STATUSES` (open/covered/dispatched/in transit) and billed
  values in `BILLED_STATUSES` (invoiced/complete) — anything else is treated
  as unbilled. Carrier (`CARRIER_KEYS`), customer (`LOAD_CUSTOMER_KEYS`),
  delivery date (`DELIVERY_DATE_KEYS`), and rate (`LOAD_RATE_KEYS`) field
  names are similarly best-effort. **Run this against real Alvys data and
  adjust those lists in `generate_report.py`** to match the actual status
  values and field names — the dashboard will show "field not found" on any
  chart it can't resolve.
