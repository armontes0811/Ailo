"""CLI: pull Loads + Invoices from Alvys and write them to an Excel report."""

import argparse
import sys
from datetime import date, timedelta

import pandas as pd
from dotenv import load_dotenv

from alvys_client import AlvysAuthError, AlvysClient


def parse_args():
    parser = argparse.ArgumentParser(description="Generate an Alvys TMS report (loads + invoices).")
    parser.add_argument("--start-date", default=(date.today() - timedelta(days=30)).isoformat())
    parser.add_argument("--end-date", default=date.today().isoformat())
    parser.add_argument("--output", default="alvys_report.xlsx")
    return parser.parse_args()


def to_dataframe(payload):
    rows = payload.get("Items", payload) if isinstance(payload, dict) else payload
    return pd.json_normalize(rows)


def main():
    load_dotenv()
    args = parse_args()

    try:
        client = AlvysClient()
    except AlvysAuthError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)

    loads_df = to_dataframe(client.search_loads(start_date=args.start_date, end_date=args.end_date))
    invoices_df = to_dataframe(client.search_invoices(start_date=args.start_date, end_date=args.end_date))

    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        loads_df.to_excel(writer, sheet_name="Loads", index=False)
        invoices_df.to_excel(writer, sheet_name="Invoices", index=False)

    print(f"Report written to {args.output}")
    print(f"  Loads: {len(loads_df)} rows")
    print(f"  Invoices: {len(invoices_df)} rows")


if __name__ == "__main__":
    main()
