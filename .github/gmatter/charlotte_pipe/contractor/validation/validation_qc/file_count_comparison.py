# copy/paste the file names and transaction count for a payment run
# make sure the counts do not have commas in the numbers

import duckdb
import re
import pandas as pd
import sys

def compare_file_names_with_counts(db_path, table_name, output_csv):
    con = duckdb.connect(db_path)

    # Query DB: get archive_file_name and row counts
    query = f"""
        SELECT archive_file_name, COUNT(*) AS row_count
        FROM {table_name}
        GROUP BY archive_file_name
    """
    db_counts = {row[0].strip(): row[1] for row in con.execute(query).fetchall() if row[0]}
    archive_set = set(db_counts.keys())

    print("Paste your file names and counts below (one per line).")
    print("Example: my_file.csv 120")
    print("Press Ctrl+D (Mac/Linux) or Ctrl+Z then Enter (Windows) when done.\n")

    # Read pasted input from terminal
    pasted_lines = sys.stdin.read().strip().splitlines()

    # Parse input
    input_file_names = []
    for line in pasted_lines:
        line = line.strip()
        if not line:
            continue
        match = re.match(r"(.+?)\s+(\d+)$", line)
        if match:
            fname, count = match.groups()
            input_file_names.append((fname.strip(), int(count)))
        else:
            input_file_names.append((line, None))

    input_set = {fname for fname, _ in input_file_names}

    results = []

    # Input → DB
    for fname, count in input_file_names:
        if fname in archive_set:
            db_count = db_counts[fname]
            if count is None:
                status = "found (no input count)"
            elif count == db_count:
                status = "found (count matches)"
            else:
                status = f"found (count mismatch: input={count}, db={db_count})"
            results.append((fname, count, db_count, status))
        else:
            results.append((fname, count, None, "not found"))

    # DB → Input
    for fname, db_count in db_counts.items():
        if fname not in input_set:
            results.append((fname, None, db_count, "MISSING"))

    # Save results
    df = pd.DataFrame(results, columns=["file_name", "input_count", "db_count", "status"])
    df.to_csv(output_csv, index=False)

    # Summary
    found_match = (df["status"] == "found (count matches)").sum()
    found_mismatch = df["status"].str.contains("count mismatch").sum()
    not_found_count = (df["status"] == "not found").sum()
    missing_count = (df["status"] == "MISSING").sum()

    print(f"\nResults written to: {output_csv}")
    print(f"\nSummary:")
    print(f"- Input files checked: {len(input_file_names)}")
    print(f"- Found with matching count: {found_match}")
    print(f"- Found with count mismatch: {found_mismatch}")
    print(f"- Not found in DB: {not_found_count}")
    print(f"- Missing from input: {missing_count}")

    return df

if __name__ == "__main__":
    db_path = "/Users/lorimartella/Documents/gmatter/charlotte_pipe/charlotte_pipe.duckdb"
    table_name = "contractor_transactions"
    output_csv = "/Users/lorimartella/Documents/gmatter/charlotte_pipe/cpf_python_scripts/outputs/file_count_results.csv"

    compare_file_names_with_counts(db_path, table_name, output_csv)
