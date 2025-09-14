import duckdb
import pandas as pd

# --- Configuration ---
database_file = "/Users/lorimartella/Documents/gmatter/charlotte_pipe/charlotte_pipe.duckdb"
table_name = "contractor_transactions"

# Keywords with priority
keywords = [
    ('tubing','pipe',1), ('conduit','pipe',1), ('hose','pipe',1),
    ('tube','pipe',1), ('solid','pipe',2),
    ('elbow','fitting',1), ('tee','fitting',1), ('coupling','fitting',1),
    ('union','fitting',1), ('valve','fitting',1), ('connector','fitting',1),
    ('adapter','fitting',1), ('reducer','fitting',1), ('cap','fitting',1),
    ('plug','fitting',1), ('flange','fitting',1), ('fitting','fitting',1),
    ('inc/red','fitting',1), ('bushing','fitting',1), ('bush','fitting',1),
    ('wyes','fitting',1), ('increaser','fitting',1)
]

# Columns to include in CSV
original_cols = [
    'cast_iron', 'plastic', 'pipe', 'fittings', 'pvc', 'abs', 'dwv',
    'cpvc', 'cts', 'neither', 'questionable', 'exclude'
]

# --- Connect to DuckDB ---
con = duckdb.connect(database_file)

# --- Create a temporary table for keywords ---
con.execute("CREATE TEMPORARY TABLE keyword_table(keyword VARCHAR, category VARCHAR, priority INTEGER)")
con.executemany("INSERT INTO keyword_table VALUES (?, ?, ?)", keywords)

# --- SQL query: collect all matching keywords per row ---
sql = f"""
WITH matches AS (
    SELECT 
        t.rowid AS row_number,
        t.contractor_name,
        t.item_description,
        t.pipe,
        t.fittings,
        {', '.join('t.' + col for col in original_cols)},
        k.keyword,
        k.category AS expected_category,
        k.priority,
        POSITION(LOWER(k.keyword) IN LOWER(t.item_description)) AS pos
    FROM {table_name} t
    JOIN keyword_table k
    ON POSITION(LOWER(k.keyword) IN LOWER(t.item_description)) > 0
),
agg_keywords AS (
    SELECT
        row_number,
        contractor_name,
        item_description,
        pipe,
        fittings,
        {', '.join(original_cols)},
        STRING_AGG(keyword, ', ') AS keywords_found,
        MAX(priority) AS highest_priority
    FROM matches
    GROUP BY row_number, contractor_name, item_description, pipe, fittings, {', '.join(original_cols)}
),
highest_category AS (
    SELECT
        m.row_number,
        m.contractor_name,
        m.item_description,
        m.pipe,
        m.fittings,
        {', '.join('m.' + col for col in original_cols)},
        m.keywords_found,
        FIRST_VALUE(k.category) OVER (PARTITION BY m.row_number ORDER BY k.priority DESC) AS final_category
    FROM agg_keywords m
    LEFT JOIN keyword_table k
    ON POSITION(LOWER(k.keyword) IN LOWER(m.keywords_found)) > 0
)
SELECT
    row_number,
    contractor_name,
    item_description,
    keywords_found,
    final_category,
    pipe AS pipe_column_value,
    fittings AS fittings_column_value,
    {', '.join(original_cols)},
    CASE
        WHEN final_category='pipe' AND (LOWER(pipe) NOT IN ('y','1','true')) 
            THEN 'Should be marked as pipe but pipe column is not true'
        WHEN final_category='fitting' AND (LOWER(fittings) NOT IN ('y','1','true')) 
            THEN 'Should be marked as fitting but fittings column is not true'
    END AS discrepancy_type
FROM highest_category
WHERE discrepancy_type IS NOT NULL
ORDER BY row_number;
"""

# --- Execute query ---
discrepancies_df = con.execute(sql).fetchdf()

# --- Save to CSV ---
discrepancies_df.to_csv("outputs/pipe_fitting_discrepancies_duckdb.csv", index=False)

print(f"Validation complete. Found {len(discrepancies_df)} discrepancies.")
