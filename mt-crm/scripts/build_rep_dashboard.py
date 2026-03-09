#!/usr/bin/env python3
"""
Build the Missing Thorn Rep Dashboard.

Reads Excel source files, generates JavaScript data objects, and injects them
into the HTML template between __DATA_START__ and __DATA_END__ markers.

Usage:
    python3 build_rep_dashboard.py --data ~/OneDrive/MT\ Dashboard\ Data/ \
        --template templates/rep_dashboard_template.html \
        --output dist/index_rep.html

If --data is not provided, defaults to ./data/ in the repo root.
If --template is not provided, defaults to templates/rep_dashboard_template.html.
If --output is not provided, defaults to dist/index_rep.html.
"""

import argparse
import json
import os
import sys
from datetime import datetime
from pathlib import Path

# Add scripts directory to path for shared modules
sys.path.insert(0, os.path.dirname(__file__))


def find_excel_files(data_dir):
    """Find all Excel files in the data directory."""
    data_path = Path(data_dir)
    if not data_path.exists():
        print(f"ERROR: Data directory not found: {data_dir}")
        print("Provide --data flag or create a data/ directory in repo root.")
        sys.exit(1)

    xlsx_files = list(data_path.glob("*.xlsx")) + list(data_path.glob("*.xls"))
    if not xlsx_files:
        print(f"WARNING: No Excel files found in {data_dir}")

    return xlsx_files


def generate_data_block(data_dir):
    """
    Generate the JavaScript data block from Excel source files.

    This function should be customized to match your specific Excel file
    structure and data processing logic.

    Expected Excel files in data_dir:
    - VIP distributor portal exports (depletions, inventory, placements)
    - QuickBooks exports (orders, revenue)

    Returns a string of JavaScript variable declarations.
    """
    # TODO: Implement Excel parsing logic
    # This is a placeholder that should be replaced with actual data processing.
    #
    # The build script should produce these JavaScript variables:
    #   stateNames, regionMap, distScorecard, accountsTop, inventoryData,
    #   distHealth, depletionsData (opportunities), reEngagementData, newWins,
    #   reorderData, qbDistOrders, warehouseInventory, classicTracker,
    #   distDetail, placementSummary, acctConcentration, sampleSummary,
    #   buildDate, dataThrough
    #
    # For now, try to import from an existing build module if available.

    try:
        from build_rep_data import generate_all_data
        return generate_all_data(data_dir)
    except ImportError:
        pass

    print("WARNING: No data generation module found.")
    print("The template will be built with placeholder data markers only.")
    print("To build with real data, implement generate_data_block() in this script")
    print(f"or create scripts/build_rep_data.py with a generate_all_data(data_dir) function.")

    now = datetime.now()
    lines = [
        f'const buildDate = "{now.strftime("%Y-%m-%d %H:%M")}";',
        f'const dataThrough = "{now.strftime("%Y-%m-%d")}";',
        '// WARNING: No data loaded - implement build_rep_data.py',
    ]
    return "\n".join(lines)


def inject_data(template_path, data_block, output_path):
    """Inject the data block into the template between markers."""
    with open(template_path, "r", encoding="utf-8") as f:
        template = f.read()

    start_marker = "// __DATA_START__"
    end_marker = "// __DATA_END__"

    start_idx = template.find(start_marker)
    end_idx = template.find(end_marker)

    if start_idx == -1 or end_idx == -1:
        print("ERROR: Could not find __DATA_START__ / __DATA_END__ markers in template.")
        sys.exit(1)

    output = (
        template[: start_idx + len(start_marker)]
        + "\n"
        + data_block
        + "\n"
        + template[end_idx:]
    )

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(output)

    print(f"Built: {output_path} ({len(output):,} bytes)")


def main():
    parser = argparse.ArgumentParser(description="Build Missing Thorn Rep Dashboard")
    parser.add_argument(
        "--data",
        default="data/",
        help="Path to directory containing Excel source files (default: data/)",
    )
    parser.add_argument(
        "--template",
        default="templates/rep_dashboard_template.html",
        help="Path to HTML template (default: templates/rep_dashboard_template.html)",
    )
    parser.add_argument(
        "--output",
        default="dist/index_rep.html",
        help="Output path for built dashboard (default: dist/index_rep.html)",
    )
    args = parser.parse_args()

    print(f"Data dir:  {args.data}")
    print(f"Template:  {args.template}")
    print(f"Output:    {args.output}")

    data_block = generate_data_block(args.data)
    inject_data(args.template, data_block, args.output)


if __name__ == "__main__":
    main()
