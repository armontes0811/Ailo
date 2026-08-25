# Ailo — Alvys TMS connection

Pulls Loads, Trips, Carriers, and Invoices out of [Alvys](https://alvys.com)
(via their public API) and writes a multi-sheet Excel report centered on
**unbilled loads** — brokerage-fleet loads that are either not yet invoiced
to the customer, not yet paid to the carrier, or both.

- **Summary** — headline KPIs: load/invoice counts, total invoiced revenue,
  total accessorial charges, unbilled load counts by stage (uninvoiced / not
  yet invoiced to carrier / awaiting carrier payment / overdue), aging, and
  $ at risk on each side.
- **Unbilled Loads** — one row per in-scope load, with `IsUninvoiced`,
  `IsNotYetInvoicedToCarrier`, `IsAwaitingCarrierPayment`,
  `IsOverdueToCarrier`, resolved `CarrierName`, and `daysSinceDelivery`.
- **Unbilled by Carrier** / **Unbilled by Customer** — the above, grouped
  with load counts, aging (avg/max days since delivery), and total value.
- **Loads** — full load/shipment activity (not filtered).
- **Invoices** — customer invoices (financial/billing data).
- **Accessorials** — invoice line items classified as accessorial charges
  (detention, lumper, layover, TONU, fuel surcharge), flattened out of the
  invoice records.

## Unbilled load logic (confirmed live against the API)

A load is in scope only if its `Fleet.Name` is `"brokerage"` and its status
isn't `Cancelled`. From there:

- **Pre-delivery loads are dropped entirely** — anything still `Quoted`,
  `Open`, `Reserved`, `Covered`, `Dispatched`, `En-Route`, or `In Transit`.
- **Uninvoiced**: the Load's own `Status` is anything other than `Invoiced`
  or `Completed`. `TONU` loads are a special case — `TONU` alone doesn't
  tell you whether it was invoiced, so those are checked directly against
  `/invoices/search` by load number instead of by status.
- **Carrier payment has two distinct stages**, both joined via the matching
  Trip (by `LoadNumber`) since Alvys tracks carrier settlement on the Trip,
  not the Load:
  - **Not yet invoiced to carrier**: `Carrier.CarrierInvoiceNumber` is null
    — the settlement is still "Open"; nothing has been created/sent to
    Triumph Pay yet. This is an internal processing backlog, not a Triumph
    delay.
  - **Awaiting carrier payment**: `Carrier.CarrierInvoiceNumber` is set
    (sent to Triumph) but `CarrierPaidAt` is still null. **Overdue to
    carrier** narrows this to loads whose Trip `DueDate` has passed —
    confirmed live to equal `ReleasedAt` + the carrier's actual term (30
    days standard, ~2 days for quickpay carriers).
  - Getting this distinction right required a full pass with the account
    owner: an earlier version conflated "not yet invoiced" with "overdue"
    because `DueDate` gets pre-populated as an *estimate* even before an
    invoice exists — checking `DueDate` alone, without confirming
    `CarrierInvoiceNumber` is actually set, wrongly flagged loads that
    hadn't even reached Triumph yet as "overdue."
  - Carrier identity/name comes from the Trip (`Trip.Carrier.Id`), not the
    Load — Alvys only exposes a carrier GUID there, so the name is resolved
    via a `/carriers/search` lookup.
- All of these flags are **independent, not combined** — a load can be
  invoiced to the customer already but still stuck at any carrier-payment
  stage, or vice versa. The dashboard filters on each separately.
- **Timestamp parsing**: Alvys's timestamps have inconsistent fractional-
  second precision across records (0, 6, or 7 digits). `pandas.to_datetime`'s
  automatic format *inference* silently turns non-matching rows into `NaT`
  under `errors="coerce"` instead of raising — confirmed live, this dropped
  a real chunk of valid dates. Every date parse in `generate_report.py` uses
  `format="ISO8601"` to avoid this; don't remove it.
- `Admin`, `In Review`, and `Financed` statuses haven't been seen in
  practice; they're not specially classified (so they'd currently show up
  as uninvoiced by default) — adjust `BILLED_STATUSES`/`PRE_DELIVERY_STATUSES`
  in `generate_report.py` if that turns out wrong.

## Setup

1. In Alvys: **Profile → API** to open the Public API page, and create an
   application to get a `client_id` and `client_secret`.
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Copy `.env.example` to `.env` and fill in your real credentials:
   ```
   cp .env.example .env
   ```
   `.env` is gitignored — your credentials never get committed. Generated
   report exports (`*.xlsx`, `*.json`) are gitignored too, since they contain
   real customer/carrier/financial data.

## Usage

```
python generate_report.py --start-date 2026-07-01 --end-date 2026-07-31 --output july_report.xlsx
```

All arguments are optional; by default it pulls the last 30 days into
`alvys_report.xlsx` in the current directory. This also writes a JSON export
(`july_report.json` in the example above) with the same data, for the
dashboard below.

## Dashboard

`dashboard.html` is a self-contained visual dashboard: KPIs (unbilled count,
uninvoiced revenue at risk, unpaid carrier cost, oldest/avg days since
delivery), a status breakdown, an aging chart (days since delivery,
bucketed), and breakdowns by carrier and customer. Three independent filters
sit above the KPIs — **View** (all unbilled / uninvoiced only / unpaid-to-
carrier only), **Status**, and **Carrier** — and every KPI/chart/table
recomputes from whatever's currently selected.

Meant to be published as a Claude Artifact; it reads report data from a
`#report-data` script tag embedded in the page, with a manual
file-load/drag-and-drop fallback for viewing an export ad hoc.

To refresh the published dashboard with new data:

```
python generate_report.py --start-date ... --end-date ...
python build_dashboard.py alvys_report.json
```

Then ask Claude to republish `dashboard.html` to the same Artifact URL. This
only works from a session that can actually reach Alvys's API (some sandboxed
environments block outbound access to `auth.alvys.com` /
`integrations.alvys.com` by policy) — "refresh the dashboard" is exactly this
three-step loop asked of Claude in chat.

## API notes

- Auth: OAuth2 client-credentials (`POST https://auth.alvys.com/oauth/token`).
- Base URL: `https://integrations.alvys.com/api/p/v1` — search endpoints are
  **POST** requests with a JSON body (`page`/`pageSize` plus filters), not
  GET with query params. Results come back paginated under an `Items` key.
- Every search endpoint requires at least one filter beyond pagination
  (e.g. `status`, or an ID/number list) — date range alone isn't enough.
- `loadNumbers` filters (on `/trips/search` and `/invoices/search`) are
  capped at 50 per request; `ids` (on `/carriers/search`) tops out around
  the 200 page-size ceiling. `alvys_client.py` batches automatically.
- Confirmed `Status` enums (via the API's own validation errors): Loads —
  `TONU, Dispatched, In Transit, Queued, Paid, Open, Completed, Financed,
  Cancelled, Covered, In Review, Released-Carrier Paid, Reserved, Trip
  Completed, En-Route, Quoted, Admin, Invoiced, Carrier Paid, Released,
  Delivered`. Invoices — `Draft, AwaitingPayment, Paid`.
- Response fields are PascalCase with nested objects (e.g. `CustomerRate.Amount`,
  `Fleet.Name`); `pandas.json_normalize` flattens them to dotted column names.
- The Alvys public API is read-only from this integration's side — nothing
  is written back to Alvys.
