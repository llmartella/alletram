#!/usr/bin/env python3
"""
payload_to_csv.py

Reads a JSON payload (with a top-level "data" array, where each item has
"potential_earnings" and "qualification_results" arrays) and writes two
CSV files:
    - potential_earnings.csv
    - qualification_results.csv

Each object within those arrays becomes one row; each key becomes a column.
The "errors" and "meta" keys at the top level are ignored.

HOW TO USE:
    Just edit the two paths below (INPUT_FILE and OUTPUT_DIR), then hit Run.
"""

import csv
import json
import sys
from pathlib import Path

# ============================================================
# EDIT THESE TWO PATHS, THEN RUN THE SCRIPT.
# ============================================================

INPUT_FILE = "/Users/lorimartella/Documents/gmatter/agvend/support/CALC troubleshooting/CALC response input/basf_response.json"

OUTPUT_DIR = "/Users/lorimartella/Documents/gmatter/agvend/support/CALC troubleshooting/CALC response output"

# ============================================================


def flatten_records(payload: dict, array_key: str) -> list[dict]:
    """
    Pull every object out of payload["data"][i][array_key] into a single
    flat list of dicts, ready to write as CSV rows.
    """
    records = []
    for participant in payload.get("data", []):
        for record in participant.get(array_key, []):
            records.append(record)
    return records


def write_csv(records: list[dict], output_path: Path) -> None:
    """
    Write a list of dicts to a CSV file. Column headers are the union of
    all keys across all records (in first-seen order), so this is safe
    even if some records have extra/missing keys.
    """
    if not records:
        print(f"No records found for {output_path.name} — skipping.")
        return

    fieldnames = []
    seen = set()
    for record in records:
        for key in record.keys():
            if key not in seen:
                seen.add(key)
                fieldnames.append(key)

    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for record in records:
            # Convert any list/dict values (e.g. "properties": []) to a
            # JSON string so they fit cleanly in a single CSV cell.
            row = {
                k: (json.dumps(v) if isinstance(v, (list, dict)) else v)
                for k, v in record.items()
            }
            writer.writerow(row)

    print(f"Wrote {len(records)} rows to {output_path}")


def main():
    input_path = Path(INPUT_FILE)
    output_dir = Path(OUTPUT_DIR)

    if not input_path.exists():
        sys.exit(f"Input file not found: {input_path}")

    with open(input_path, "r", encoding="utf-8") as f:
        payload = json.load(f)

    potential_earnings = flatten_records(payload, "potential_earnings")
    qualification_results = flatten_records(payload, "qualification_results")

    write_csv(potential_earnings, output_dir / "potential_earnings.csv")
    write_csv(qualification_results, output_dir / "qualification_results.csv")


if __name__ == "__main__":
    main()