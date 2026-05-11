import duckdb
import pandas as pd

def load_file_to_duckdb(
    input_file,
    duckdb_file="/Users/lorimartella/Documents/gmatter/charlotte_pipe/charlotte_pipe.duckdb"
):
    if input_file.lower().endswith(".csv"):
        df = pd.read_csv(input_file)
    else:
        df = pd.read_excel(input_file)

    conn = duckdb.connect(duckdb_file)

    conn.execute("DROP TABLE IF EXISTS contractor_transactions")
    conn.register("df_view", df)
    conn.execute("""
        CREATE TABLE contractor_transactions AS
        SELECT * FROM df_view
    """)

    print("Table created with columns:")
    print(conn.execute(
        "PRAGMA table_info('contractor_transactions')"
    ).fetchall())

    conn.close()

if __name__ == "__main__":
    input_file = "/Users/lorimartella/Downloads/transactions_20260427.xlsx"
    load_file_to_duckdb(input_file)
    print("Data successfully loaded into contractor_transactions table.")
