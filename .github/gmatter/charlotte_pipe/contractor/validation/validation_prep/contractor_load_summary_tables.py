# run contractor_clear_summary_tables first
# this loads data from all summary sheets to duckdb

import pandas as pd
import duckdb

excel_file = "/Users/lorimartella/Downloads/20250901_summaries.xlsx"
duckdb_file = "/Users/lorimartella/Documents/gmatter/charlotte_pipe/charlotte_pipe.duckdb"

# Mapping for exceptions where sheet name ≠ table name
sheet_to_table = {
    "By Contractor": "by_contractor",
    "By Item": "by_item"
}

conn = duckdb.connect(database=duckdb_file, read_only=False)
xls = pd.ExcelFile(excel_file)

for sheet_name in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet_name)

    # Use mapped table name if available, else use sheet name
    table_name = sheet_to_table.get(sheet_name, sheet_name)

    # Register DataFrame as a DuckDB temp view
    conn.register("temp_df", df)

    # Create the table if it doesn’t exist, otherwise append
    conn.execute(f'''
        CREATE TABLE IF NOT EXISTS "{table_name}" AS SELECT * FROM temp_df LIMIT 0
    ''')
    conn.execute(f'INSERT INTO "{table_name}" SELECT * FROM temp_df')

    conn.unregister("temp_df")

conn.close()
print("All sheets loaded successfully!")
