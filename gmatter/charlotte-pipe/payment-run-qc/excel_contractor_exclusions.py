import pandas as pd


def create_exclusions_report(excel_path, output_csv, sheet_name="Sheet1"):
    """
    Creates exclusions report by validating item descriptions in an Excel
    sheet against exclusion terms.
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

    # Load data from Excel
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    df = df.reset_index(drop=True)

    results = []
    for _, row in df.iterrows():
        item_desc = str(row["item_description"] or "").lower()
        exclude_flag = str(row["exclude"] or "").strip()

        for term in exclusion_terms:
            if term.lower() in item_desc:
                # Issue: Found exclusion term but exclude != Y
                if exclude_flag.lower() != "y":
                    results.append([
                        row["item_description"],
                        exclude_flag,
                        f'Found "{term}" but exclude is not Y'
                    ])
                break  # stop checking after first term match.

    if results:
        report_df = pd.DataFrame(results, columns=[
            "item_description",
            "exclude",
            "issue_type"
        ])
    else:
        # Empty report but with headers
        report_df = pd.DataFrame(columns=[
            "item_description",
            "exclude",
            "issue_type"
        ])

    # Save to CSV
    report_df.to_csv(output_csv, index=False)
    print(f"Exclusions report written to {output_csv} with {len(report_df)} issues.")



if __name__ == "__main__":
    create_exclusions_report(
        excel_path="/Users/lorimartella/Downloads/20260810_summaries.xlsx",
        output_csv="/Users/lorimartella/Documents/gmatter/charlotte_pipe/cpf_python_scripts/outputs/03_exclusions_excel.csv",
        sheet_name="item_descriptions"
    )