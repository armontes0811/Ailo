"""Embed a generate_report.py JSON export into dashboard.html for publishing.

Usage:
    python generate_report.py --start-date 2026-07-01 --end-date 2026-08-25
    python build_dashboard.py alvys_report.json

This rewrites dashboard.html in place with the report data embedded in its
`#report-data` script tag, ready to publish (or re-publish) as an Artifact.
"""

import argparse
import json
import re

PLACEHOLDER_PATTERN = re.compile(
    r'(<script id="report-data" type="application/json">)(.*?)(</script>)',
    re.DOTALL,
)


def parse_args():
    parser = argparse.ArgumentParser(description="Embed a report JSON export into dashboard.html.")
    parser.add_argument("json_path", help="Path to the JSON export from generate_report.py")
    parser.add_argument("--dashboard", default="dashboard.html")
    return parser.parse_args()


def main():
    args = parse_args()

    with open(args.json_path) as f:
        data = json.load(f)

    with open(args.dashboard) as f:
        html = f.read()

    new_html, count = PLACEHOLDER_PATTERN.subn(
        lambda m: m.group(1) + json.dumps(data) + m.group(3), html, count=1
    )
    if count == 0:
        raise SystemExit(f"Could not find the #report-data placeholder in {args.dashboard}")

    with open(args.dashboard, "w") as f:
        f.write(new_html)

    print(f"Embedded {args.json_path} into {args.dashboard}")


if __name__ == "__main__":
    main()
