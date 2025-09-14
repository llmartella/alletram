# deletes all rows in all tables related to the summary file (exposure)

import duckdb

# Path to your DuckDB database
duckdb_file = "/Users/lorimartella/Documents/gmatter/charlotte_pipe/charlotte_pipe.duckdb"

# Tables you want to clear
tables = [
    "by_contractor",
    "exposure_summary",
    "program_summary",
    "offer_summary",
    "program_rate_summary",
    "by_item"
]

# Connect to DuckDB
conn = duckdb.connect(database=duckdb_file, read_only=False)

# Delete all rows from each table
for table in tables:
    conn.execute(f'DELETE FROM "{table}"')
    # Alternative: conn.execute(f'TRUNCATE "{table}"')  # also works, faster if supported

conn.close()

print("All specified tables have been cleared.")
