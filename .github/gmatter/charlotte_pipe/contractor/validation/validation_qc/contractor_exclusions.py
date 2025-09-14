# checks for contractor transactions that should be excluded based on keywords (competitors, non-covered products)
# requires the transactions data to be loaded to duckdb in contractor_transactions table

import duckdb
import pandas as pd

def create_exclusions_report(db_path="mydata.duckdb", output_csv="03_exclusions.csv"):
    """
    Creates exclusions report by validating contractor_transactions
    against exclusion terms.
    """

    exclusion_terms = [
        'copper',
        'spears',
        'tyler',
        'JM eagle',
        'lasco',
        'ipex',
        'nibco'
    ]

    con = duckdb.connect(db_path)

    # Load data from contractor_transactions
    df = con.execute("""
        SELECT contractor_name,
               archive_file_name,
               item_description,
               exclude,
               potential_earnings
        FROM contractor_transactions
    """).df()

    results = []
    for _, row in df.iterrows():
        item_desc = str(row["item_description"] or "").lower()
        exclude_flag = str(row["exclude"] or "").strip()
        potential_earnings = row["potential_earnings"]

        for term in exclusion_terms:
            if term.lower() in item_desc:
                # Issue: Found exclusion term but exclude != "Y"
                if exclude_flag != "Y":
                    results.append([
                        row["contractor_name"],
                        row["archive_file_name"],
                        row["item_description"],
                        exclude_flag,
                        f'Found "{term}" but exclude is not "Y"'
                    ])
                # Issue: Found exclusion term but potential_earnings is not empty
                if pd.notna(potential_earnings) and str(potential_earnings).strip() != "":
                    results.append([
                        row["contractor_name"],
                        row["archive_file_name"],
                        row["item_description"],
                        exclude_flag,
                        f'Found "{term}" but potential_earnings is not empty'
                    ])
                break  # stop checking after first term match

    if results:
        report_df = pd.DataFrame(results, columns=[
            "contractor_name",
            "archive_file_name",
            "item_description",
            "exclude",
            "issue_type"
        ])
    else:
        # Empty report but with headers
        report_df = pd.DataFrame(columns=[
            "contractor_name",
            "archive_file_name",
            "item_description",
            "exclude",
            "issue_type"
        ])

    # Save to CSV
    report_df.to_csv(output_csv, index=False)
    print(f"Exclusions report written to {output_csv} with {len(report_df)} issues.")

    con.close()


if __name__ == "__main__":
    create_exclusions_report(
        db_path="/Users/lorimartella/Documents/gmatter/charlotte_pipe/charlotte_pipe.duckdb",
        output_csv="outputs/03_exclusions.csv"
    )
