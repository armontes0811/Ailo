"""Pull fresh detention data from Alvys and bake it into dashboard.html.

Run this, then publish/redeploy dashboard.html as the Detention Tracker
Artifact. This is the script the daily refresh (and any on-demand refresh)
runs -- see README.md.
"""

import argparse
import json
import sys
from datetime import date, datetime, timedelta

from dotenv import load_dotenv

from alvys_client import AlvysAuthError
from generate_detention_report import build_rows, write_excel
import pandas as pd
from generate_detention_report import DETAIL_COLUMNS

TEMPLATE_PATH = "dashboard_template.html"


def parse_args():
    parser = argparse.ArgumentParser(description="Refresh the Detention Tracker dashboard from Alvys.")
    parser.add_argument("--start-date", default=(date.today() - timedelta(days=30)).isoformat())
    parser.add_argument("--end-date", default=date.today().isoformat())
    parser.add_argument("--xlsx-output", default="detention_report.xlsx")
    parser.add_argument("--html-output", default="dashboard.html")
    return parser.parse_args()


def bake_dashboard(rows, generated_at_label, template_path=TEMPLATE_PATH):
    with open(template_path, "r") as f:
        html = f.read()

    # Escape "</script>" inside any string field so it can't break out of
    # the inline <script> block it's embedded in.
    data_json = json.dumps(rows).replace("</script", "<\\/script")

    html = html.replace("/*__EMBEDDED_ROWS__*/ []", data_json)
    html = html.replace("/*__GENERATED_AT__*/", generated_at_label)
    return html


def main():
    load_dotenv()
    args = parse_args()

    try:
        rows = build_rows(args.start_date, args.end_date)
    except AlvysAuthError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)

    if not rows:
        print("No stops found for the given date range -- not overwriting dashboard.html.", file=sys.stderr)
        sys.exit(1)

    rows_df = pd.DataFrame(rows, columns=DETAIL_COLUMNS)
    write_excel(rows_df, args.xlsx_output)

    generated_at_label = datetime.now().strftime("%b %-d, %Y %-I:%M %p")
    html = bake_dashboard(rows, generated_at_label)
    with open(args.html_output, "w") as f:
        f.write(html)

    flagged = sum(1 for r in rows if r["detention_flag"])
    print(f"Wrote {args.html_output} ({len(rows)} stops, {flagged} flagged) and {args.xlsx_output}")
    print(f"Publish {args.html_output} with the Artifact tool to update the live dashboard.")


if __name__ == "__main__":
    main()
