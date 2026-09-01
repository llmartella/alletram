import pandas as pd

# --- Configuration ---
excel_file = "/Users/lorimartella/Downloads/20260810_summaries.xlsx"
sheet_name = "item_descriptions"  # change to a sheet name (string) if not the first sheet
output_csv = "/Users/lorimartella/Documents/gmatter/charlotte_pipe/cpf_python_scripts/outputs/pipe_fitting_discrepancies_excel.csv"

# Keywords with priority (keyword, category, priority)
keywords = [
    ('tubing', 'pipe', 1), ('conduit', 'pipe', 1), ('hose', 'pipe', 1),
    ('tube', 'pipe', 1), ('solid', 'pipe', 2),
    ('elbow', 'fitting', 1), ('tee', 'fitting', 1), ('coupling', 'fitting', 1),
    ('union', 'fitting', 1), ('valve', 'fitting', 1), ('connector', 'fitting', 1),
    ('adapter', 'fitting', 1), ('reducer', 'fitting', 1), ('cap', 'fitting', 1),
    ('plug', 'fitting', 1), ('flange', 'fitting', 1), ('fitting', 'fitting', 1),
    ('inc/red', 'fitting', 1), ('bushing', 'fitting', 1), ('bush', 'fitting', 1),
    ('wyes', 'fitting', 1), ('increaser', 'fitting', 1)
]

# Columns to carry through into the output, in addition to item_description/pipe/fittings
# NOTE: 'competitor' removed - not present in current file structure (as of 2026-08 layout)
original_cols = [
    'cast_iron', 'plastic', 'pvc', 'dwv', 'cpvc', 'cts', 'abs',
    'exclude', 'segment'
]

# --- Load Excel into a DataFrame ---
df = pd.read_excel(excel_file, sheet_name=sheet_name)
df = df.reset_index(drop=True)
df.insert(0, 'row_number', df.index + 1)


def find_matches(description):
    """Return list of (keyword, category, priority) tuples found in the description."""
    if not isinstance(description, str):
        return []
    desc_lower = description.lower()
    return [(kw, cat, pr) for kw, cat, pr in keywords if kw.lower() in desc_lower]


def classify_row(description):
    matches = find_matches(description)
    if not matches:
        return pd.Series({'keywords_found': None, 'final_category': None})
    # dedupe keywords while preserving order
    seen = []
    for kw, _, _ in matches:
        if kw not in seen:
            seen.append(kw)
    keywords_found = ', '.join(seen)
    # pick the category of the highest-priority match (ties -> first found)
    final_category = max(matches, key=lambda m: m[2])[1]
    return pd.Series({'keywords_found': keywords_found, 'final_category': final_category})


classified = df['item_description'].apply(classify_row)
df = pd.concat([df, classified], axis=1)


def discrepancy_type(row):
    pipe_val = str(row['pipe']).strip().lower()
    fittings_val = str(row['fittings']).strip().lower()
    truthy = {'y', '1', 'true'}
    if row['final_category'] == 'pipe' and pipe_val not in truthy:
        return 'Should be marked as pipe but pipe column is not true'
    if row['final_category'] == 'fitting' and fittings_val not in truthy:
        return 'Should be marked as fitting but fittings column is not true'
    return None


df['discrepancy_type'] = df.apply(discrepancy_type, axis=1)

output_cols = (
    ['row_number', 'item_description', 'keywords_found', 'final_category', 'pipe', 'fittings']
    + original_cols
    + ['discrepancy_type']
)

discrepancies_df = df[df['discrepancy_type'].notna()][output_cols].copy()
discrepancies_df = discrepancies_df.rename(columns={'pipe': 'pipe_column_value', 'fittings': 'fittings_column_value'})

# --- Save to CSV ---
discrepancies_df.to_csv(output_csv, index=False)

print(f"Validation complete. Found {len(discrepancies_df)} discrepancies.")