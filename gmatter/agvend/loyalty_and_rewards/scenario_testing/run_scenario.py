"""
Reads a transactions CSV (the kind of thing you'd enter in Excel), runs it
through the earnings engine, and writes the SAME rows back out with the
computed columns appended -- so each transaction sits right next to its own
result. Row order matches your input file, not date order (the engine still
processes in date order internally; it just reports back in the order you
gave it).

INPUT CSV columns (header row required):
    date, category, amount, txn_id (txn_id is optional), plus optionally
    expected_earnings and expected_level (your own manual calculation per
    row, checked against the engine's actual result -- see below). Any other
    columns you include, e.g. participant name, are preserved and passed
    through too.

    category must be one of: fert, chem, seed, corn, beans_wheat,
    refined_fuels, propane, lubes
    date format: YYYY-MM-DD preferred, but common Excel variants (9/20/26,
    9/20/2026, etc.) are tolerated too.

OUTPUT CSV columns:
    <all of your original columns, unchanged> + cum_after, earning_quantity,
    actual_earnings, effective_rate, actual_level, category_mode, then a
    running_<category> column for EVERY category -- a full snapshot of
    cumulative spend across every category as of that transaction. If you
    filled in expected_earnings/expected_level, you also get earnings_passed
    and level_passed (TRUE/FALSE, blank if you left the expected value empty).

HOW TO RUN FROM VS CODE:
    Edit the three settings under "SCENARIO SETTINGS" below, then hit Run.
    OUTPUT_CSV can equal INPUT_CSV to overwrite in place.

    NOTE: if this file is open in Excel while you run the script, close it
    first -- Excel won't auto-refresh to show the script's changes, and
    saving from Excel afterward can silently overwrite the script's output.

For a version that reads/writes directly to a Google Sheet instead of a
local CSV (so there's no open-file/stale-view issue), see
run_scenario_sheets.py.
"""

import csv
from datetime import datetime

from engine import Transaction, simulate_scenario
import report


# =============================================================================
# SCENARIO SETTINGS — edit these, then Run
# =============================================================================
INPUT_CSV = "working_scenario.csv"    # your active test transactions — edit this file directly
OUTPUT_CSV = "working_scenario.csv"   # overwrites in place so running totals stay next to your data
ENROLL_DATE = "2026-09-15"            # YYYY-MM-DD, or None to treat everything as enrolled
# =============================================================================


def load_transactions(path: str):
    """Returns (transactions, raw_rows, fieldnames) so original columns/order
    can be preserved when writing results back out."""
    txns = []
    raw_rows = []
    with open(path, newline="") as f:
        reader = csv.DictReader(f)
        fieldnames = list(reader.fieldnames or [])
        for i, row in enumerate(reader):
            raw_rows.append(row)
            txns.append(
                Transaction(
                    date=report.parse_date(row["date"], row_num=i + 2),  # +2: header + 1-indexed
                    category=row["category"].strip(),
                    amount=report.parse_amount(row["amount"], row_num=i + 2),
                    txn_id=row.get("txn_id") or f"row{i+1}",
                )
            )
    return txns, raw_rows, fieldnames


def write_results(raw_rows, fieldnames, results, path: str) -> None:
    out_fieldnames, out_rows = report.build_output_rows(raw_rows, fieldnames, results)
    with open(path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=out_fieldnames)
        writer.writeheader()
        writer.writerows(out_rows)


def main():
    enroll_date = None
    if ENROLL_DATE:
        enroll_date = datetime.strptime(ENROLL_DATE, "%Y-%m-%d").date()

    txns, raw_rows, fieldnames = load_transactions(INPUT_CSV)
    results = simulate_scenario(txns, enroll_date=enroll_date)  # returned in original row order
    write_results(raw_rows, fieldnames, results, OUTPUT_CSV)
    report.print_summary(raw_rows, results, destination_label=OUTPUT_CSV)


if __name__ == "__main__":
    main()