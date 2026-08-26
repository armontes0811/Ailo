# Ailo — Alvys TMS connection

Pulls Loads (and, best-effort, Invoices) out of [Alvys](https://alvys.com)
via their API, and turns Loads/Stops data into a per-stop **detention
report** — flagging any stop where the driver was there more than 2 hours
past the scheduled appointment.

## Setup

1. In Alvys: **Profile → API** to open the Public API page, and create an
   application to get a `client_id` and `client_secret`.
2. Install dependencies:
   ```
   python3 -m venv .venv && source .venv/bin/activate
   pip install -r requirements.txt
   ```
3. Copy `.env.example` to `.env` and fill in your real credentials:
   ```
   cp .env.example .env
   ```
   `.env` is gitignored — your credentials never get committed.

## Detention report (CLI)

```
python generate_detention_report.py --start-date 2026-07-01 --end-date 2026-07-31 --output july_detention.xlsx
```

Pulls every Load in Alvys, flattens each Stop into one row (customer, stop
type, location, appointment, arrival, departure), and flags any stop where
departure was more than 2 hours after the appointment (for stops with a
scheduled-window instead of a fixed appointment, the end of that window is
used) -- **and only if the driver arrived on time** (at or before the
appointment). A late arrival never counts as detention, no matter how long
the stop took, since the driver didn't hold up their end of the appointment
either. If arrival isn't known, it doesn't qualify (can't confirm on-time).
Add `--customer "Some Customer"` to scope to one customer.

Columns worth knowing:
- `on_time` — whether the driver arrived at or before the appointment.
  Detention only applies when this is true.
- `hours_from_appointment` — the full span from appointment to departure
  (includes the free 2-hour window), regardless of `on_time`.
- `detention_hours` — the billable excess past that free window, `0` unless
  `on_time` is true (and the stop still ran over 2h). This is what "By
  Customer" totals and the dashboard's "Total detention" stat add up.

Writes:
- `july_detention.xlsx` — a "Detention Detail" sheet (filterable, flagged
  rows highlighted red) and a "By Customer" summary sheet.
- `july_detention.json` — the same rows as JSON, in the shape
  `dashboard_template.html` expects.

**Note on dates:** Alvys's `/loads/search` accepts a `dateRange` filter but
it does not actually narrow results server-side (confirmed against a live
call — `Total` was identical across wildly different ranges). So this pulls
every load in the "executed" statuses (dispatched through paid/invoiced —
see `EXECUTED_STATUSES` in `alvys_client.py`) and filters by
appointment/arrival/departure date client-side. That's a few thousand loads
per pull on a busy account; expect it to take a bit.

## Detention Tracker dashboard (live Artifact)

`dashboard_template.html` is a self-contained dashboard (filter by customer,
sort by any column, flag stops over 2h) published as a Claude Artifact.
Artifacts can't hold API credentials or call Alvys directly from the
browser, so data is baked into the page at publish time instead:

```
python refresh_dashboard.py
```

pulls the last 30 days (via the same logic as the CLI report above),
writes `detention_report.xlsx`, and writes `dashboard.html` — a copy of
`dashboard_template.html` with the pulled rows embedded. Publishing
`dashboard.html` with Claude's Artifact tool updates the live dashboard.

**Daily refresh:** a scheduled Routine ("Detention Tracker daily refresh",
~7am ET) re-runs this and republishes the same Artifact automatically, so
the dashboard reflects roughly the last day's activity without anyone
running anything by hand. Ask Claude to refresh it early, change the
schedule, or turn it off at any time — it just runs `refresh_dashboard.py`
and republishes.

`dashboard.html` and the `.xlsx`/`.json` outputs are generated, gitignored
files — only `dashboard_template.html` (the template) is committed.

## Notes / things to verify

- Auth: `POST https://auth.alvys.com/oauth/token` (client-credentials).
  Verified against a live call.
- Loads: `POST https://integrations.alvys.com/api/p/v1/loads/search`.
  Verified against a live call on 2026-08-26 — field names in
  `detention.py` (`AppointmentDate`, `ArrivedAt`, `DepartedAt`,
  `CustomerName`, etc.) come from a real response, not docs guesses.
- Invoices: `search_invoices` in `alvys_client.py` is still an unverified
  guess (only the base URL has been corrected) — check a live response
  before relying on it.
