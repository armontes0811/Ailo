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
    rows = payload.get("data", payload) if isinstance(payload, dict) else payload
    return pd.json_normalize(rows)


def main():
    load_dotenv()
    args = parse_args()

    try:
        client = AlvysClient()
    except AlvysAuthError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)

    # NOTE: /loads/search doesn't actually filter by date server-side (see
    # alvys_client.py) -- this pulls every load in the "executed" statuses
    # and doesn't narrow by --start-date/--end-date. Use
    # generate_detention_report.py if you need a date-scoped, stop-level view.
    loads_df = pd.json_normalize(client.search_all_loads())
    invoices_df = to_dataframe(client.search_invoices())

    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        loads_df.to_excel(writer, sheet_name="Loads", index=False)
        invoices_df.to_excel(writer, sheet_name="Invoices", index=False)

    print(f"Report written to {args.output}")
    print(f"  Loads: {len(loads_df)} rows")
    print(f"  Invoices: {len(invoices_df)} rows")


if __name__ == "__main__":
    main()
