"""CLI: pull Loads + Invoices from Alvys and write a multi-sheet Excel report
covering load/shipment activity, financial/billing, and accessorial charges."""

import argparse
import json
import sys
from datetime import date, timedelta

import pandas as pd
from dotenv import load_dotenv

from alvys_client import AlvysAuthError, AlvysClient

# NOTE: Alvys's exact invoice line-item field names aren't confirmed against a
# live response yet (see README caveats). We check each of these keys, in
# order, for a list of charge line items on an invoice record.
LINE_ITEM_KEYS = ("lineItems", "charges", "chargeLines", "invoiceLines")

# Substrings (case-insensitive) used to classify a line item as accessorial
# rather than linehaul/freight revenue. Adjust once real charge-type values
# from Alvys are confirmed.
ACCESSORIAL_KEYWORDS = (
    "detention", "lumper", "stop off", "stopoff", "layover", "tonu",
    "truck order not used", "fuel surcharge", "fsc", "storage", "escort",
    "permit", "tarp", "pallet exchange", "reconsign", "after hours",
    "holiday", "weekend", "hazmat", "oversize", "overweight",
)


def parse_args():
    parser = argparse.ArgumentParser(
        description="Generate an Alvys TMS report (loads, invoices, accessorials, summary)."
    )
    parser.add_argument("--start-date", default=(date.today() - timedelta(days=30)).isoformat())
    parser.add_argument("--end-date", default=date.today().isoformat())
    parser.add_argument("--output", default="alvys_report.xlsx")
    parser.add_argument(
        "--json-output",
        default=None,
        help="Path for the JSON export used by the dashboard artifact "
        "(defaults to --output with a .json extension).",
    )
    return parser.parse_args()


def to_dataframe(payload):
    rows = payload.get("data", payload) if isinstance(payload, dict) else payload
    return pd.json_normalize(rows)


def to_records(df):
    """DataFrame -> list of JSON-safe dicts (NaN becomes null)."""
    if df.empty:
        return []
    return json.loads(df.to_json(orient="records", date_format="iso"))


def _find_line_items(invoice):
    for key in LINE_ITEM_KEYS:
        items = invoice.get(key)
        if isinstance(items, list) and items:
            return items
    return []


def _is_accessorial(line_item):
    text = " ".join(
        str(line_item.get(field, ""))
        for field in ("type", "chargeType", "category", "description")
    ).lower()
    return any(keyword in text for keyword in ACCESSORIAL_KEYWORDS)


def extract_line_items(invoices_payload):
    """Flatten invoice charge line items into one row per line item, tagged
    with the parent invoice's number/load id/customer for traceability."""
    invoices = invoices_payload.get("data", invoices_payload) if isinstance(invoices_payload, dict) else invoices_payload
    rows = []
    for invoice in invoices:
        if not isinstance(invoice, dict):
            continue
        context = {
            "invoiceNumber": invoice.get("invoiceNumber") or invoice.get("invoiceId"),
            "loadId": invoice.get("loadId"),
            "customerName": invoice.get("customerName"),
        }
        for item in _find_line_items(invoice):
            if isinstance(item, dict):
                rows.append({**context, **item})
    return pd.json_normalize(rows)


def build_accessorials_df(invoices_payload):
    line_items_df = extract_line_items(invoices_payload)
    if line_items_df.empty:
        return line_items_df
    mask = line_items_df.apply(
        lambda row: _is_accessorial({
            "type": row.get("type", ""),
            "chargeType": row.get("chargeType", ""),
            "category": row.get("category", ""),
            "description": row.get("description", ""),
        }),
        axis=1,
    )
    return line_items_df[mask].reset_index(drop=True)


def build_summary_df(loads_df, invoices_df, accessorials_df, start_date, end_date):
    total_revenue = None
    for col in ("totalAmount", "amount", "total"):
        if col in invoices_df.columns:
            total_revenue = pd.to_numeric(invoices_df[col], errors="coerce").sum()
            break

    accessorial_total = None
    for col in ("amount", "total"):
        if col in accessorials_df.columns:
            accessorial_total = pd.to_numeric(accessorials_df[col], errors="coerce").sum()
            break

    return pd.DataFrame([
        {"metric": "Start date", "value": start_date},
        {"metric": "End date", "value": end_date},
        {"metric": "Loads", "value": len(loads_df)},
        {"metric": "Invoices", "value": len(invoices_df)},
        {"metric": "Total invoiced revenue", "value": total_revenue},
        {"metric": "Accessorial line items", "value": len(accessorials_df)},
        {"metric": "Total accessorial charges", "value": accessorial_total},
    ])


def main():
    load_dotenv()
    args = parse_args()

    try:
        client = AlvysClient()
    except AlvysAuthError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)

    loads_payload = client.search_loads(start_date=args.start_date, end_date=args.end_date)
    invoices_payload = client.search_invoices(start_date=args.start_date, end_date=args.end_date)

    loads_df = to_dataframe(loads_payload)
    invoices_df = to_dataframe(invoices_payload)
    accessorials_df = build_accessorials_df(invoices_payload)
    summary_df = build_summary_df(loads_df, invoices_df, accessorials_df, args.start_date, args.end_date)

    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        loads_df.to_excel(writer, sheet_name="Loads", index=False)
        invoices_df.to_excel(writer, sheet_name="Invoices", index=False)
        accessorials_df.to_excel(writer, sheet_name="Accessorials", index=False)

    json_output = args.json_output or (
        args.output[:-5] + ".json" if args.output.endswith(".xlsx") else args.output + ".json"
    )
    with open(json_output, "w") as f:
        json.dump(
            {
                "meta": {
                    "startDate": args.start_date,
                    "endDate": args.end_date,
                    "generatedAt": date.today().isoformat(),
                },
                "summary": to_records(summary_df),
                "loads": to_records(loads_df),
                "invoices": to_records(invoices_df),
                "accessorials": to_records(accessorials_df),
            },
            f,
            indent=2,
        )

    print(f"Report written to {args.output}")
    print(f"JSON export (for the dashboard artifact) written to {json_output}")
    print(f"  Loads: {len(loads_df)} rows")
    print(f"  Invoices: {len(invoices_df)} rows")
    print(f"  Accessorial line items: {len(accessorials_df)} rows")


if __name__ == "__main__":
    main()
