"""CLI: pull Alvys loads for statement-invoiced customers and write a CSV
formatted for the Statement Tracker artifact's Import tab.

Statement-invoiced customers are billed on a booking or vessel basis
instead of load-by-load, so this filters loads down to the configured
customers and exports the columns the tracker groups on (Customer, Order #,
Load #, PO #, Vessel Name, Status, Date, Origin, Destination, Amount).

Field mapping was verified against a live /api/p/v1/loads/search response:
Vessel Name isn't a top-level field -- it's an entry in each load's
`References` list (Name == "Vessel Name"). Origin/destination come from the
`Stops` list (StopType == "Pickup" / "Delivery").
"""

import argparse
import csv
import sys

from dotenv import load_dotenv

from alvys_client import AlvysAuthError, AlvysClient

# Substring match against the load's CustomerName (case-insensitive). These
# need to be specific: Alvys has multiple divisions with overlapping names,
# e.g. "GLOVIS America, Inc." and "GLOVIS America, Inc. (Drayage)" are NOT
# the same account as "Glovis (Mobis Division) America, Inc", and
# "Berg Mill Supply" is not the same account as
# "Berg Mill Supply (Imports/Exports)". Keep these in sync with the customer
# settings in the Statement Tracker artifact.
DEFAULT_CUSTOMERS = ["mobis", "imports/exports", "jess smith"]

# Statuses relevant to statement invoicing -- loads still in earlier stages
# (Open, Dispatched, In Transit, ...) aren't ready to bill yet. Pass
# --statuses to override.
DEFAULT_STATUSES = ["Delivered", "Invoiced", "Completed", "Paid"]

CSV_HEADERS = [
    "Customer",
    "Order #",
    "Load #",
    "PO #",
    "Vessel Name",
    "Status",
    "Date",
    "Origin",
    "Destination",
    "Amount",
]

PAGE_SIZE = 200


def parse_args():
    parser = argparse.ArgumentParser(
        description="Export Alvys loads for statement-invoiced customers to a CSV for the Statement Tracker."
    )
    parser.add_argument("--start-date", help="ISO-8601 datetime, e.g. 2026-07-01T00:00:00.000Z")
    parser.add_argument("--end-date", help="ISO-8601 datetime, e.g. 2026-08-31T23:59:59.000Z")
    parser.add_argument(
        "--customers",
        default=",".join(DEFAULT_CUSTOMERS),
        help="Comma-separated substrings to match against each load's customer name (case-insensitive).",
    )
    parser.add_argument(
        "--statuses",
        default=",".join(DEFAULT_STATUSES),
        help="Comma-separated Alvys load statuses to search (see Alvys for the full enum).",
    )
    parser.add_argument("--output", default="statement_loads.csv")
    return parser.parse_args()


def reference_value(record, name):
    for ref in record.get("References") or []:
        if str(ref.get("Name", "")).strip().lower() == name.lower():
            return ref.get("Value", "")
    return ""


def stop_location(record, stop_type):
    for stop in record.get("Stops") or []:
        if str(stop.get("StopType", "")).strip().lower() == stop_type.lower():
            addr = stop.get("Address") or {}
            city, state = addr.get("City"), addr.get("State")
            if city and state:
                return f"{city}, {state}"
            return city or state or ""
    return ""


def amount_value(record):
    for path in ("CustomerRate", "Linehaul"):
        amount = (record.get(path) or {}).get("Amount")
        if amount not in (None, ""):
            return amount
    return ""


def matches_customer(customer_name, needles):
    lowered = str(customer_name or "").lower()
    return any(needle in lowered for needle in needles)


def to_row(record):
    return [
        record.get("CustomerName", ""),
        record.get("OrderNumber", ""),
        record.get("LoadNumber", ""),
        record.get("PONumber", ""),
        reference_value(record, "Vessel Name"),
        record.get("Status", ""),
        record.get("ScheduledDeliveryAt") or record.get("ScheduledPickupAt") or record.get("CreatedAt", ""),
        stop_location(record, "Pickup"),
        stop_location(record, "Delivery"),
        amount_value(record),
    ]


def fetch_all_loads(client, statuses, start_date, end_date):
    records = []
    page = 0
    while True:
        payload = client.search_loads(
            start_date=start_date, end_date=end_date, page=page, page_size=PAGE_SIZE, status=statuses
        )
        items = payload.get("Items", [])
        records.extend(items)
        total = payload.get("Total", len(records))
        if not items or len(records) >= total:
            break
        page += 1
    return records


def main():
    load_dotenv()
    args = parse_args()
    needles = [c.strip().lower() for c in args.customers.split(",") if c.strip()]
    statuses = [s.strip() for s in args.statuses.split(",") if s.strip()]

    try:
        client = AlvysClient()
    except AlvysAuthError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)

    try:
        records = fetch_all_loads(client, statuses, args.start_date, args.end_date)
    except Exception as exc:
        print(f"Error calling Alvys: {exc}", file=sys.stderr)
        sys.exit(1)

    rows = [to_row(r) for r in records if matches_customer(r.get("CustomerName"), needles)]

    with open(args.output, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(CSV_HEADERS)
        writer.writerows(rows)

    print(f"Searched {len(records)} loads across statuses {statuses}, matched {len(rows)} for customers {needles}.")
    print(f"Wrote {args.output}")
    print("Paste this file's contents (or upload it) into the Statement Tracker artifact's Import tab.")


if __name__ == "__main__":
    main()
