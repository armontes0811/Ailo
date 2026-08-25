"""CLI: pull Loads, Trips, Carriers, and Invoices from Alvys and write a
multi-sheet Excel report focused on unbilled loads (delivered but not yet
invoiced to the customer, and/or not yet paid to the carrier), plus the
underlying load/shipment activity, financial/billing, and accessorial data.

Confirmed live against the Alvys public API (see alvys_client.py and
README.md for endpoint details). Two independent flags drive the unbilled
report -- a load can be uninvoiced, unpaid-to-carrier, or both:

- Uninvoiced: the Load's own Status is anything other than "Invoiced" or
  "Completed" (loads still in a pre-delivery status -- Quoted/Open/
  Reserved/Covered/Dispatched/En-Route/In Transit -- are dropped entirely,
  as are Cancelled loads). TONU loads are handled separately since "TONU"
  doesn't tell you invoiced-or-not by itself -- we cross-reference
  invoices/search by LoadNumbers to check directly.
- Unpaid to carrier: the matching Trip's CarrierPaidAt is null. Carrier
  identity/name comes from the Trip (Trip.Carrier.Id -> carriers/search),
  not the Load.
"""

import argparse
import json
import sys
from datetime import date, timedelta

import pandas as pd
from dotenv import load_dotenv

from alvys_client import AlvysAuthError, AlvysClient

# Confirmed via a live 400 response from /loads/search (and matches the
# documented /trips/search enum, which is a 16-value subset of this list).
ALL_LOAD_STATUSES = [
    "TONU", "Dispatched", "In Transit", "Queued", "Paid", "Open", "Completed",
    "Financed", "Cancelled", "Covered", "In Review", "Released-Carrier Paid",
    "Reserved", "Trip Completed", "En-Route", "Quoted", "Admin", "Invoiced",
    "Carrier Paid", "Released", "Delivered",
]

# Loads in these statuses haven't been dispatched/delivered yet -- excluded
# from the unbilled report entirely (confirmed with the report requester).
PRE_DELIVERY_STATUSES = {"Quoted", "Open", "Reserved", "Covered", "Dispatched", "En-Route", "In Transit"}
# Dead loads -- not "unbilled", just irrelevant to this report.
DROPPED_STATUSES = {"Cancelled"}
# Only these two count as "already invoiced" to the customer (confirmed).
BILLED_STATUSES = {"Invoiced", "Completed"}
TONU_STATUS = "TONU"

# Confirmed via a live 400 response from /invoices/search.
ALL_INVOICE_STATUSES = ["Draft", "AwaitingPayment", "Paid"]

# Only loads on this Fleet belong on this report (confirmed with the report
# requester -- other Fleet values represent other business entities).
IN_SCOPE_FLEET_NAME = "brokerage"

# NOTE: Alvys's invoice line-item charge-type values aren't confirmed beyond
# "Linehaul" and "Fuel Surcharge" from a live sample. Adjust as more appear.
ACCESSORIAL_KEYWORDS = (
    "detention", "lumper", "stop off", "stopoff", "layover", "tonu",
    "truck order not used", "fuel surcharge", "fsc", "storage", "escort",
    "permit", "tarp", "pallet exchange", "reconsign", "after hours",
    "holiday", "weekend", "hazmat", "oversize", "overweight",
)


def parse_args():
    parser = argparse.ArgumentParser(
        description="Generate an Alvys TMS report (unbilled loads, loads, invoices, accessorials)."
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


def filter_brokerage_loads(loads_df):
    if "Fleet.Name" not in loads_df.columns:
        return loads_df.iloc[0:0].copy()
    mask = loads_df["Fleet.Name"].astype(str).str.strip().str.lower() == IN_SCOPE_FLEET_NAME
    return loads_df[mask].copy()


def find_tonu_invoiced_load_numbers(client, tonu_load_numbers):
    """Which TONU loads already have an invoice, determined directly by
    cross-referencing invoices/search rather than guessing from Status."""
    if not tonu_load_numbers:
        return set()
    invoices_df = to_dataframe(client.search_invoices_by_load_numbers(tonu_load_numbers))
    if invoices_df.empty or "Loads" not in invoices_df.columns:
        return set()
    invoiced = set()
    for loads in invoices_df["Loads"]:
        if isinstance(loads, list):
            invoiced.update(l.get("LoadNumber") for l in loads if isinstance(l, dict) and l.get("LoadNumber"))
    return invoiced


def build_candidate_unbilled_df(client, loads_df):
    """Non-cancelled, brokerage-fleet loads that are either post-delivery
    (and not yet invoiced) or TONU (and not yet invoiced), with an
    `IsUninvoiced` column. Pre-delivery and Cancelled loads are dropped."""
    brokerage_df = filter_brokerage_loads(loads_df)
    if brokerage_df.empty or "Status" not in brokerage_df.columns:
        return brokerage_df.assign(IsUninvoiced=pd.Series(dtype=bool))

    non_cancelled = brokerage_df[~brokerage_df["Status"].isin(DROPPED_STATUSES)].copy()

    is_tonu = non_cancelled["Status"] == TONU_STATUS
    tonu_df = non_cancelled[is_tonu].copy()
    rest_df = non_cancelled[~is_tonu].copy()

    rest_df = rest_df[~rest_df["Status"].isin(PRE_DELIVERY_STATUSES)].copy()
    rest_df["IsUninvoiced"] = ~rest_df["Status"].isin(BILLED_STATUSES)

    if not tonu_df.empty:
        invoiced_load_numbers = find_tonu_invoiced_load_numbers(client, tonu_df["LoadNumber"].tolist())
        tonu_df["IsUninvoiced"] = ~tonu_df["LoadNumber"].isin(invoiced_load_numbers)

    return pd.concat([rest_df, tonu_df], ignore_index=True)


def attach_trip_and_carrier_info(client, candidate_df):
    """Join Trip (carrier identity, invoice/payment state, DueDate) and
    resolved carrier Name onto each candidate load, by LoadNumber ->
    Trip.LoadNumber.

    Three-stage carrier-payment model, confirmed live against real Trip
    data (see README):
    - Not yet invoiced to carrier: Carrier.CarrierInvoiceNumber is null --
      the settlement is still "Open," nothing has been sent to Triumph Pay
      yet. DueDate can already be populated here as an *estimate* -- it
      does NOT mean an invoice exists, so it must never be used alone to
      decide "overdue."
    - Awaiting carrier payment: CarrierInvoiceNumber is set (sent to
      Triumph) but CarrierPaidAt is still null.
    - Overdue to carrier: awaiting carrier payment AND DueDate has passed --
      DueDate = ReleasedAt + the carrier's actual term (30 days standard,
      ~2 days for quickpay, confirmed against 200+ live trips).
    CarrierPaidAt set at all means paid -- excluded from every bucket above.
    """
    if candidate_df.empty or "LoadNumber" not in candidate_df.columns:
        return candidate_df.assign(CarrierName=None, CarrierPaidAt=None, DueDate=None,
                                    HasCarrierInvoice=False, IsUnpaidToCarrier=False,
                                    IsNotYetInvoicedToCarrier=False, IsAwaitingCarrierPayment=False,
                                    IsOverdueToCarrier=False)

    trips_df = to_dataframe(client.search_trips_by_load_numbers(candidate_df["LoadNumber"].tolist()))
    trip_cols = ["LoadNumber", "Carrier.Id", "Carrier.CarrierInvoiceNumber", "CarrierPaidAt", "DueDate",
                 "Carrier.TotalPayable.Amount"]
    for col in trip_cols:
        if col not in trips_df.columns:
            trips_df[col] = pd.NA
    trips_df = trips_df[trip_cols].drop_duplicates(subset="LoadNumber", keep="last")

    merged = candidate_df.merge(trips_df, on="LoadNumber", how="left")

    carrier_ids = merged["Carrier.Id"].dropna().unique().tolist()
    if carrier_ids:
        carriers_df = to_dataframe(client.search_carriers_by_ids(carrier_ids))
        name_map = dict(zip(carriers_df.get("Id", []), carriers_df.get("Name", [])))
    else:
        name_map = {}
    merged["CarrierName"] = merged["Carrier.Id"].map(name_map)

    merged["HasCarrierInvoice"] = merged["Carrier.CarrierInvoiceNumber"].notna()
    merged["IsUnpaidToCarrier"] = merged["Carrier.Id"].notna() & merged["CarrierPaidAt"].isna()
    merged["IsNotYetInvoicedToCarrier"] = merged["IsUnpaidToCarrier"] & ~merged["HasCarrierInvoice"]
    merged["IsAwaitingCarrierPayment"] = merged["IsUnpaidToCarrier"] & merged["HasCarrierInvoice"]

    # format="ISO8601" is required here: Alvys's timestamps have inconsistent
    # fractional-second precision (0/6/7 digits) across records, and pandas's
    # format *inference* silently coerces non-matching rows to NaT otherwise --
    # confirmed live, this dropped a real chunk of valid overdue dates.
    due_date = pd.to_datetime(merged["DueDate"], errors="coerce", utc=True, format="ISO8601")
    merged["IsOverdueToCarrier"] = merged["IsAwaitingCarrierPayment"] & due_date.notna() & (due_date < pd.Timestamp.now(tz="UTC"))
    return merged


def build_unbilled_df(client, loads_df):
    candidate_df = build_candidate_unbilled_df(client, loads_df)
    if candidate_df.empty:
        return candidate_df.assign(CarrierName=None, IsUnpaidToCarrier=False, daysSinceDelivery=pd.NA)

    enriched_df = attach_trip_and_carrier_info(client, candidate_df)
    unbilled_df = enriched_df[enriched_df["IsUninvoiced"] | enriched_df["IsUnpaidToCarrier"]].copy()

    if "DeliveredAt" in unbilled_df.columns:
        delivered_at = pd.to_datetime(unbilled_df["DeliveredAt"], errors="coerce", utc=True, format="ISO8601")
        now = pd.Timestamp.now(tz="UTC")
        unbilled_df["daysSinceDelivery"] = (now - delivered_at).dt.days
    else:
        unbilled_df["daysSinceDelivery"] = pd.NA

    sort_col = "daysSinceDelivery" if "daysSinceDelivery" in unbilled_df.columns else None
    if sort_col:
        unbilled_df = unbilled_df.sort_values(sort_col, ascending=False, na_position="last")
    return unbilled_df.reset_index(drop=True)


def summarize_unbilled_by(unbilled_df, group_col, value_col=None):
    columns = ["group", "loadCount", "avgDaysSinceDelivery", "maxDaysSinceDelivery", "totalValue"]
    if unbilled_df.empty or group_col not in unbilled_df.columns:
        return pd.DataFrame(columns=columns)

    working = unbilled_df.copy()
    working["_group"] = working[group_col].where(working[group_col].notna() & (working[group_col] != ""), "(unspecified)")

    summary = working.groupby("_group").size().rename("loadCount").reset_index().rename(columns={"_group": "group"})

    if "daysSinceDelivery" in working.columns:
        days = working.groupby("_group")["daysSinceDelivery"].agg(["mean", "max"]).reset_index().rename(
            columns={"_group": "group", "mean": "avgDaysSinceDelivery", "max": "maxDaysSinceDelivery"}
        )
        days["avgDaysSinceDelivery"] = days["avgDaysSinceDelivery"].round(1)
        summary = summary.merge(days, on="group", how="left")
    else:
        summary["avgDaysSinceDelivery"] = pd.NA
        summary["maxDaysSinceDelivery"] = pd.NA

    if value_col and value_col in working.columns:
        totals = working.groupby("_group")[value_col].apply(
            lambda s: pd.to_numeric(s, errors="coerce").sum()
        ).rename("totalValue").reset_index().rename(columns={"_group": "group"})
        summary = summary.merge(totals, on="group", how="left")
    else:
        summary["totalValue"] = pd.NA

    return summary.sort_values("loadCount", ascending=False).reset_index(drop=True)


def _find_line_items(invoice):
    items = invoice.get("LineItems")
    return items if isinstance(items, list) else []


def _is_accessorial(name):
    text = str(name or "").lower()
    return any(keyword in text for keyword in ACCESSORIAL_KEYWORDS)


def extract_line_items(invoices_payload):
    """Flatten invoice charge line items into one row per line item, tagged
    with the parent invoice's number/customer for traceability."""
    invoices = invoices_payload.get("data", invoices_payload) if isinstance(invoices_payload, dict) else invoices_payload
    rows = []
    for invoice in invoices:
        if not isinstance(invoice, dict):
            continue
        context = {
            "InvoiceNumber": invoice.get("Number"),
            "CustomerName": (invoice.get("Customer") or {}).get("Name"),
        }
        for item in _find_line_items(invoice):
            if isinstance(item, dict):
                rows.append({**context, **item})
    return pd.json_normalize(rows)


def build_accessorials_df(invoices_payload):
    line_items_df = extract_line_items(invoices_payload)
    if line_items_df.empty or "Name" not in line_items_df.columns:
        return line_items_df
    mask = line_items_df["Name"].apply(_is_accessorial)
    return line_items_df[mask].reset_index(drop=True)


def build_summary_df(loads_df, invoices_df, accessorials_df, unbilled_df, start_date, end_date):
    total_revenue = pd.to_numeric(invoices_df.get("Total.Amount"), errors="coerce").sum() if "Total.Amount" in invoices_df.columns else None
    accessorial_total = pd.to_numeric(accessorials_df.get("Amount.Amount"), errors="coerce").sum() if "Amount.Amount" in accessorials_df.columns else None

    def subset(flag_col):
        return unbilled_df[unbilled_df.get(flag_col, False)] if flag_col in unbilled_df.columns else unbilled_df.iloc[0:0]

    uninvoiced_df = subset("IsUninvoiced")
    not_yet_invoiced_carrier_df = subset("IsNotYetInvoicedToCarrier")
    awaiting_carrier_df = subset("IsAwaitingCarrierPayment")
    overdue_df = subset("IsOverdueToCarrier")

    avg_days = max_days = None
    if "daysSinceDelivery" in unbilled_df.columns:
        days = pd.to_numeric(unbilled_df["daysSinceDelivery"], errors="coerce")
        if days.notna().any():
            avg_days = round(days.mean(), 1)
            max_days = int(days.max())

    def value_sum(df):
        return pd.to_numeric(df.get("Carrier.TotalPayable.Amount"), errors="coerce").sum() if "Carrier.TotalPayable.Amount" in df.columns else None

    uninvoiced_value = pd.to_numeric(uninvoiced_df.get("CustomerRate.Amount"), errors="coerce").sum() if "CustomerRate.Amount" in uninvoiced_df.columns else None
    not_yet_invoiced_carrier_value = value_sum(not_yet_invoiced_carrier_df)
    awaiting_carrier_value = value_sum(awaiting_carrier_df)
    overdue_value = value_sum(overdue_df)

    return pd.DataFrame([
        {"metric": "Start date", "value": start_date},
        {"metric": "End date", "value": end_date},
        {"metric": "Loads", "value": len(loads_df)},
        {"metric": "Invoices", "value": len(invoices_df)},
        {"metric": "Total invoiced revenue", "value": total_revenue},
        {"metric": "Accessorial line items", "value": len(accessorials_df)},
        {"metric": "Total accessorial charges", "value": accessorial_total},
        {"metric": "Unbilled loads (uninvoiced or unpaid-to-carrier)", "value": len(unbilled_df)},
        {"metric": "Uninvoiced loads", "value": len(uninvoiced_df)},
        {"metric": "Not yet invoiced to carrier (settlement still Open)", "value": len(not_yet_invoiced_carrier_df)},
        {"metric": "Awaiting carrier payment (invoiced, Triumph hasn't paid)", "value": len(awaiting_carrier_df)},
        {"metric": "Overdue to carrier (awaiting payment, past DueDate)", "value": len(overdue_df)},
        {"metric": "Avg days since delivery (unbilled)", "value": avg_days},
        {"metric": "Oldest unbilled (days since delivery)", "value": max_days},
        {"metric": "Uninvoiced revenue at risk", "value": uninvoiced_value},
        {"metric": "Not-yet-invoiced-to-carrier cost", "value": not_yet_invoiced_carrier_value},
        {"metric": "Awaiting carrier payment cost", "value": awaiting_carrier_value},
        {"metric": "Overdue carrier cost", "value": overdue_value},
    ])


def main():
    load_dotenv()
    args = parse_args()

    try:
        client = AlvysClient()
    except AlvysAuthError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)

    loads_payload = client.search_loads(start_date=args.start_date, end_date=args.end_date, status=ALL_LOAD_STATUSES)
    invoices_payload = client.search_invoices(start_date=args.start_date, end_date=args.end_date, status=ALL_INVOICE_STATUSES)

    loads_df = to_dataframe(loads_payload)
    invoices_df = to_dataframe(invoices_payload)
    accessorials_df = build_accessorials_df(invoices_payload)
    unbilled_df = build_unbilled_df(client, loads_df)
    unbilled_by_carrier_df = summarize_unbilled_by(unbilled_df, "CarrierName", "Carrier.TotalPayable.Amount")
    unbilled_by_customer_df = summarize_unbilled_by(unbilled_df, "CustomerName", "CustomerRate.Amount")
    summary_df = build_summary_df(loads_df, invoices_df, accessorials_df, unbilled_df, args.start_date, args.end_date)

    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        unbilled_df.to_excel(writer, sheet_name="Unbilled Loads", index=False)
        unbilled_by_carrier_df.to_excel(writer, sheet_name="Unbilled by Carrier", index=False)
        unbilled_by_customer_df.to_excel(writer, sheet_name="Unbilled by Customer", index=False)
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
                "unbilled": to_records(unbilled_df),
                "unbilledByCarrier": to_records(unbilled_by_carrier_df),
                "unbilledByCustomer": to_records(unbilled_by_customer_df),
            },
            f,
            indent=2,
        )

    print(f"Report written to {args.output}")
    print(f"JSON export (for the dashboard artifact) written to {json_output}")
    print(f"  Loads: {len(loads_df)} rows")
    print(f"  Invoices: {len(invoices_df)} rows")
    print(f"  Accessorial line items: {len(accessorials_df)} rows")
    print(f"  Unbilled loads: {len(unbilled_df)} rows")


if __name__ == "__main__":
    main()
