#!/usr/bin/env python3
"""
run_sql_credit.py
-----------------
Runs all generated SQL files against the transaction_mapping_credit table
in DuckDB, with post-insert data quality validation.

Validation checks (per file loaded):
  HARD FAIL  - Empty rows (all data fields NULL)
  HARD FAIL  - Missing item_description (required field)
  HARD FAIL  - Junk-only values (dashes, equals, asterisks) in any column
  WARNING    - Sparse columns (NULL in some rows but not all)
  WARNING    - Rows skipped due to blank/filter (raw range vs loaded count)

Rows are kept on failure so you can inspect them in DataGrip.
"""

import argparse
import sys
from pathlib import Path

import duckdb

# ---------------------------------------------------------------------------
# Default paths — edit these to match your environment
# ---------------------------------------------------------------------------
DEFAULT_DB  = "/Users/lorimartella/Documents/gmatter/charlotte_pipe/charlotte_pipe.duckdb"
DEFAULT_SQL = "/Users/lorimartella/Documents/gmatter/charlotte_pipe/cpf_python_scripts/sql_output_credit"

# ---------------------------------------------------------------------------
# Columns included in quality checks
# ---------------------------------------------------------------------------
DATA_COLUMNS = [
    "customer_name", "vendor", "sales_order_number", "ordered_on",
    "item_sku", "item_sku_alt", "item_sku_category", "item_upc",
    "material_group_number", "item_description",
    "unit_price", "ship_quantity", "uom", "extended_price",
]

TEXT_COLUMNS = [
    "customer_name", "vendor", "sales_order_number",
    "item_sku", "item_sku_alt", "item_sku_category", "item_upc",
    "material_group_number", "item_description", "uom",
]

JUNK_PATTERN = r"^[\-=\*\s]+$"
TARGET_TABLE = "transaction_mapping_credit"


def get_meta(con):
    """Read and drop the _meta temp table left by the SQL, if present."""
    try:
        meta_tables = con.execute("""
            SELECT table_name FROM duckdb_tables()
            WHERE temporary = true AND table_name LIKE '%_meta'
            ORDER BY table_name
        """).fetchall()
        if not meta_tables:
            return None
        meta_table = meta_tables[-1][0]
        row = con.execute(f"""
            SELECT archive_file_name, sheet_name, raw_rows, loaded_rows, has_explicit_filter
            FROM {meta_table} LIMIT 1
        """).fetchone()
        con.execute(f"DROP TABLE IF EXISTS {meta_table}")
        return row
    except Exception:
        return None


def validate(con, archive_file_name):
    """Run quality checks on rows just inserted. Returns (failures, warnings)."""
    failures = []
    warnings = []

    # 1. Empty rows
    null_checks = " AND ".join(f"{c} IS NULL" for c in DATA_COLUMNS)
    empty_rows = con.execute(f"""
        SELECT row_id FROM {TARGET_TABLE}
        WHERE archive_file_name = ? AND ({null_checks})
        ORDER BY row_id
    """, [archive_file_name]).fetchall()
    if empty_rows:
        ids = [str(r[0]) for r in empty_rows]
        failures.append({
            "check":  "Empty rows",
            "detail": f"{len(ids)} row(s) with all data columns NULL",
            "rows":   ids,
        })

    # 2. Missing item_description
    missing_desc = con.execute(f"""
        SELECT row_id FROM {TARGET_TABLE}
        WHERE archive_file_name = ?
          AND (item_description IS NULL OR trim(item_description) = '')
        ORDER BY row_id
    """, [archive_file_name]).fetchall()
    if missing_desc:
        ids = [str(r[0]) for r in missing_desc]
        failures.append({
            "check":  "Missing item_description",
            "detail": f"{len(ids)} row(s) with NULL or blank item_description",
            "rows":   ids,
        })

    # 3. Junk-only values
    junk_hits = []
    for col in TEXT_COLUMNS:
        rows = con.execute(f"""
            SELECT row_id, {col} FROM {TARGET_TABLE}
            WHERE archive_file_name = ?
              AND {col} IS NOT NULL
              AND regexp_matches({col}, '{JUNK_PATTERN}')
            ORDER BY row_id
        """, [archive_file_name]).fetchall()
        for row_id, val in rows:
            junk_hits.append({"row_id": row_id, "column": col, "value": val})
    if junk_hits:
        by_col = {}
        for h in junk_hits:
            by_col.setdefault(h["column"], []).append(f"row {h['row_id']}={h['value']!r}")
        failures.append({
            "check":  "Junk values",
            "detail": f"{len(junk_hits)} junk value(s) in {len(by_col)} column(s)",
            "rows":   [f"{col}: {', '.join(vals[:5])}" for col, vals in by_col.items()],
        })

    # 4. Sparse columns (warning only)
    total = con.execute(f"""
        SELECT COUNT(*) FROM {TARGET_TABLE} WHERE archive_file_name = ?
    """, [archive_file_name]).fetchone()[0]
    if total > 0:
        for col in DATA_COLUMNS:
            null_count = con.execute(f"""
                SELECT COUNT(*) FROM {TARGET_TABLE}
                WHERE archive_file_name = ? AND {col} IS NULL
            """, [archive_file_name]).fetchone()[0]
            non_null = total - null_count
            if 0 < non_null < total:
                pct_null = (null_count / total) * 100
                warnings.append({
                    "check":  "Sparse column",
                    "detail": f"'{col}': {non_null:,} rows have a value, "
                              f"{null_count:,} are NULL ({pct_null:.0f}% NULL)",
                })

    return failures, warnings


def print_issue(issue, symbol):
    print(f"      {symbol} [{issue['check']}] {issue['detail']}")
    for r in issue.get("rows", [])[:5]:
        print(f"          → {r}")
    excess = len(issue.get("rows", [])) - 5
    if excess > 0:
        print(f"          → ... and {excess} more")


def main():
    ap = argparse.ArgumentParser(description=f"Run generated SQL files into {TARGET_TABLE}")
    ap.add_argument("--db",  default=DEFAULT_DB,  help="Path to your DuckDB database file")
    ap.add_argument("--sql", default=DEFAULT_SQL, help="Folder containing the .sql files")
    args = ap.parse_args()

    db_path = Path(args.db)
    sql_dir = Path(args.sql)

    if not sql_dir.exists():
        print(f"ERROR: sql folder not found: {sql_dir}", file=sys.stderr)
        sys.exit(1)

    sql_files = sorted(
        f for f in sql_dir.glob("*.sql")
        if f.name not in ("00_create_table_credit.sql",)
    )

    if not sql_files:
        print(f"No SQL files found in {sql_dir}")
        sys.exit(0)

    print(f"\nConnecting to {db_path} ...")
    con = duckdb.connect(str(db_path))

    before = con.execute(f"SELECT COUNT(*) FROM {TARGET_TABLE}").fetchone()[0]
    print(f"Rows before run : {before:,}")
    print(f"Files to process: {len(sql_files)}\n")

    load_failed    = []
    quality_failed = []
    all_warnings   = []
    succeeded      = []

    for sql_file in sql_files:
        print(f"  {'─'*56}")
        print(f"  File: {sql_file.name}")

        # ── Step 1: Load ──────────────────────────────────────────────
        sql = sql_file.read_text(encoding="utf-8")

        # Split into individual statements and execute one at a time.
        # This avoids DuckDB re-encoding special characters (e.g. & → &amp;)
        # that can occur when passing a large multi-statement string.
        statements = [s.strip() for s in sql.split(";") if s.strip()]

        try:
            count_before = con.execute(
                f"SELECT COUNT(*) FROM {TARGET_TABLE}"
            ).fetchone()[0]
            for stmt in statements:
                con.execute(stmt)
            count_after = con.execute(
                f"SELECT COUNT(*) FROM {TARGET_TABLE}"
            ).fetchone()[0]
            rows_added = count_after - count_before
            print(f"  ✓ Loaded  ({rows_added:,} rows inserted)")
        except Exception as e:
            print(f"  ✗ LOAD FAILED")
            for line in str(e).splitlines():
                print(f"      {line}")
            load_failed.append((sql_file.name, str(e)))
            get_meta(con)
            continue

        if rows_added == 0:
            print(f"  ? No rows inserted — skipping validation")
            get_meta(con)
            succeeded.append(sql_file.name)
            continue

        # ── Step 2: Check raw vs loaded row counts ────────────────────
        file_warnings = []
        meta = get_meta(con)
        if meta:
            _, sheet_name, raw_rows, loaded_rows, has_explicit_filter = meta
            skipped = raw_rows - loaded_rows
            if skipped > 0 and not has_explicit_filter:
                msg = (f"'{sheet_name}': {skipped} row(s) skipped "
                       f"({loaded_rows:,} loaded from {raw_rows:,} rows in range)")
                print(f"  ⚠  Skipped rows: {msg}")
                file_warnings.append({"check": "Skipped rows", "detail": msg})

        # ── Step 3: Identify archive_file_name ────────────────────────
        archive_file_name = con.execute(f"""
            SELECT DISTINCT archive_file_name
            FROM {TARGET_TABLE}
            ORDER BY rowid DESC
            LIMIT 1
        """).fetchone()[0]

        # ── Step 4: Validate data quality ─────────────────────────────
        failures, warnings = validate(con, archive_file_name)
        file_warnings.extend(warnings)

        if file_warnings:
            all_warnings.append((sql_file.name, file_warnings))
            for w in file_warnings:
                if w["check"] != "Skipped rows":
                    print_issue(w, "⚠")

        if failures:
            print(f"  ✗ QUALITY CHECKS FAILED  (rows kept for inspection)")
            for f in failures:
                print_issue(f, "✗")
            quality_failed.append((sql_file.name, failures))
        else:
            print(f"  ✓ Quality checks passed")
            succeeded.append(sql_file.name)

    # ── Summary ───────────────────────────────────────────────────────
    after = con.execute(f"SELECT COUNT(*) FROM {TARGET_TABLE}").fetchone()[0]
    con.close()

    print(f"\n{'='*60}")
    print(f"  SUMMARY")
    print(f"{'='*60}")
    print(f"  Files run        : {len(sql_files)}")
    print(f"  Load failures    : {len(load_failed)}")
    print(f"  Quality failures : {len(quality_failed)}")
    print(f"  Warnings         : {len(all_warnings)}")
    print(f"  Rows added       : {after - before:,}  ({before:,} → {after:,})")

    if load_failed:
        print(f"\n  Load failures:")
        for name, err in load_failed:
            print(f"    ✗ {name}")
            print(f"      {err.splitlines()[0]}")

    if quality_failed:
        print(f"\n  Quality failures (data kept — inspect in DataGrip):")
        for name, failures in quality_failed:
            print(f"    ✗ {name}")
            for f in failures:
                print(f"      [{f['check']}] {f['detail']}")

    if all_warnings:
        print(f"\n  Warnings:")
        for name, warnings in all_warnings:
            print(f"    ⚠  {name}")
            for w in warnings:
                print(f"      {w['detail']}")

    if not load_failed and not quality_failed:
        print(f"\n  All files loaded and validated successfully.")
    else:
        sys.exit(1)


if __name__ == "__main__":
    main()
