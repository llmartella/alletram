# drop contractor_transactions before running this
# this script loads all contractor transactions to duckdb in contracto_transactions table

import duckdb
import pandas as pd

def load_excel_to_duckdb(excel_file, duckdb_file="/Users/lorimartella/Documents/gmatter/charlotte_pipe/charlotte_pipe.duckdb"):
    # Read Excel into a pandas DataFrame
    df = pd.read_excel(excel_file)

    # Open a connection to DuckDB (creates file if it doesn't exist)
    conn = duckdb.connect(duckdb_file)

    # Create or replace the contractor_transactions table with DataFrame schema
    conn.execute("DROP TABLE IF EXISTS contractor_transactions")
    conn.register("df_view", df)
    conn.execute("CREATE TABLE contractor_transactions AS SELECT * FROM df_view")

    # Verify columns
    print("Table created with columns:")
    print(conn.execute("PRAGMA table_info('contractor_transactions')").fetchall())

    conn.close()


if __name__ == "__main__":
    # Example usage
    excel_file = "/Users/lorimartella/Downloads/transactions_20250901.xlsx"  # <-- replace with your Excel file path
    load_excel_to_duckdb(excel_file)
    print("Data successfully loaded into contractor_transactions table.")
