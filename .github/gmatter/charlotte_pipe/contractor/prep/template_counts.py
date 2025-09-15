# place all Template files in /Users/lorimartella/Documents/gmatter/charlotte_pipe/templates
# be sure to remove previous payment run files
# then run this to get the count of transactions in each file

import duckdb
import pandas as pd

def get_column_type(con, table, column):
    """Return the DuckDB logical type of a column"""
    row = con.execute(f"PRAGMA table_info({table})").fetchdf()
    dtype = row.loc[row['name'] == column, 'type']
    return dtype.iloc[0] if not dtype.empty else None


def validate_total_counts(con):
    # Check column type for # Transactions
    col_type = get_column_type(con, "by_contractor", "# Transactions")

    if col_type and "VARCHAR" in col_type.upper():
        col_expr = 'CAST("# Transactions" AS BIGINT)'
    else:
        col_expr = '"# Transactions"'

    contractor_total = con.execute(f"""
        SELECT SUM({col_expr}) as contractor_total
        FROM by_contractor
    """).fetchone()[0] or 0

    transactions_total = con.execute("""
        SELECT COUNT(*) as transactions_total
        FROM contractor_transactions
    """).fetchone()[0]

    results = pd.DataFrame([
        ["By Contractor", contractor_total],
        ["Contractor Transactions", transactions_total]
    ], columns=["Source", "Transaction Count"])

    results["Match?"] = "YES" if contractor_total == transactions_total else "NO"
    results.to_csv("outputs/01_total_count_validation.csv", index=False)
    return results


def validate_contractor_counts(con):
    # Check column type for # Transactions
    col_type = get_column_type(con, "by_contractor", "# Transactions")

    if col_type and "VARCHAR" in col_type.upper():
        col_expr = 'CAST("# Transactions" AS BIGINT)'
    else:
        col_expr = '"# Transactions"'

    contractor_counts = con.execute(f"""
        SELECT Contractor, {col_expr} as contractor_count
        FROM by_contractor
    """).df()

    transaction_counts = con.execute("""
        SELECT contractor_name as Contractor,
               COUNT(*) as transactions_count
        FROM contractor_transactions
        GROUP BY contractor_name
    """).df()

    merged = pd.merge(
        contractor_counts,
        transaction_counts,
        on="Contractor",
        how="outer"
    ).fillna(0)

    merged["contractor_count"] = merged["contractor_count"].astype(int)
    merged["transactions_count"] = merged["transactions_count"].astype(int)
    merged["Match?"] = merged["contractor_count"] == merged["transactions_count"]
    merged["Match?"] = merged["Match?"].map({True: "YES", False: "NO"})

    merged.to_csv("outputs/02_contractor_count_validation.csv", index=False)
    return merged


def run_validations(db_path="mydata.duckdb"):
    con = duckdb.connect(db_path)

    print("Running total counts validation...")
    total_results = validate_total_counts(con)
    print(total_results)

    print("\nRunning contractor counts validation...")
    contractor_results = validate_contractor_counts(con)
    print(contractor_results)

    con.close()


if __name__ == "__main__":
    run_validations("/Users/lorimartella/Documents/gmatter/charlotte_pipe/charlotte_pipe.duckdb")
