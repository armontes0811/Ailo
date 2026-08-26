"""CLI: pull Alvys loads for statement-invoiced customers and write a CSV
formatted for the Statement Tracker artifact's Import tab.

Statement-invoiced customers are billed on a booking or vessel basis
instead of load-by-load, so this filters loads down to the configured
customers and exports the columns the tracker groups on (Customer, Order #,
Load #, PO #, Vessel Name, Status, Date, Origin, Destination, Amount).
"""

import argparse
import csv
import sys

from dotenv import load_dotenv

from alvys_client import AlvysAuthError, AlvysClient

# Substring match against the load's customer name (case-insensitive) --
# keep in sync with the customer settings in the Statement Tracker artifact.
DEFAULT_CUSTOMERS = ["glovis", "jess smith", "berg mill"]

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

# NOTE: Alvys field names below are best-effort guesses (see README caveats
# on alvys_client.py) -- once you have a live response, check these against
# the actual payload and adjust as needed. Each entry lists candidate field
# names, in priority order, tried against a load record.
FIELD_CANDIDATES = {
    "customer": ["customerName", "customer.name", "billToName", "customer"],
    "orderNumber": ["orderNumber", "referenceNumber", "orderId", "orderNo"],
    "loadNumber": ["loadNumber", "loadId", "loadNo", "id"],
    "poNumber": ["poNumber", "bookingNumber", "customerReferenceNumber", "po"],
    "vesselName": ["vesselName", "vessel", "vesselVoyage"],
    "status": ["status", "loadStatus"],
    "date": ["deliveryDate", "dropDate", "pickupDate", "date"],
    "origin": ["originCity", "pickupCity", "origin"],
    "destination": ["destinationCity", "dropCity", "destination"],
    "amount": ["totalRate", "customerRate", "rate", "amount"],
}


def parse_args():
    parser = argparse.ArgumentParser(
        description="Export Alvys loads for statement-invoiced customers to a CSV for the Statement Tracker."
    )
    parser.add_argument("--start-date")
    parser.add_argument("--end-date")
    parser.add_argument(
        "--customers",
        default=",".join(DEFAULT_CUSTOMERS),
        help="Comma-separated substrings to match against each load's customer name (case-insensitive).",
    )
    parser.add_argument("--output", default="statement_loads.csv")
    return parser.parse_args()


def dig(record, dotted_path):
    value = record
    for part in dotted_path.split("."):
        if not isinstance(value, dict):
            return None
        value = value.get(part)
    return value


def field(record, candidates):
    for candidate in candidates:
        value = dig(record, candidate)
        if value not in (None, ""):
            return value
    return ""


def matches_customer(customer_name, needles):
    lowered = str(customer_name or "").lower()
    return any(needle in lowered for needle in needles)


def to_row(record):
    return [
        field(record, FIELD_CANDIDATES["customer"]),
        field(record, FIELD_CANDIDATES["orderNumber"]),
        field(record, FIELD_CANDIDATES["loadNumber"]),
        field(record, FIELD_CANDIDATES["poNumber"]),
        field(record, FIELD_CANDIDATES["vesselName"]),
        field(record, FIELD_CANDIDATES["status"]),
        field(record, FIELD_CANDIDATES["date"]),
        field(record, FIELD_CANDIDATES["origin"]),
        field(record, FIELD_CANDIDATES["destination"]),
        field(record, FIELD_CANDIDATES["amount"]),
    ]


def main():
    load_dotenv()
    args = parse_args()
    needles = [c.strip().lower() for c in args.customers.split(",") if c.strip()]

    try:
        client = AlvysClient()
    except AlvysAuthError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)

    payload = client.search_loads(start_date=args.start_date, end_date=args.end_date)
    records = payload.get("data", payload) if isinstance(payload, dict) else payload

    rows = [
        to_row(record)
        for record in records
        if matches_customer(field(record, FIELD_CANDIDATES["customer"]), needles)
    ]

    with open(args.output, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(CSV_HEADERS)
        writer.writerows(rows)

    print(f"Wrote {len(rows)} statement loads to {args.output}")
    print("Paste this file's contents (or upload it) into the Statement Tracker artifact's Import tab.")


if __name__ == "__main__":
    main()
