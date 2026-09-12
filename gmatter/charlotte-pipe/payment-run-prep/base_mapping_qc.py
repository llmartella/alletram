#!/usr/bin/env python3
"""
base_mapping_qc.py
-------------------
Combined generate + load + validate script for the base/unspecified pipeline.

  1. Auto-detects this month's mapping doc from Drive (searching the month
     folder and its subfolders), or uses VSCODE_SHEET_ID as a manual override.
  2. Generates one DuckDB SQL file per Excel file referenced in the mapping,
     into a sql_output folder next to this script.
  3. Immediately loads each generated SQL file into DuckDB and runs
     data-quality validation on the inserted rows.
  4. Prints a SUMMARY block.

Doing generate + run in one script (rather than two separate scripts and a
shared output folder) means there's no path to get out of sync between them
— whatever this run generates is exactly what this run then loads.

Mapping sheet columns (order-independent):
    payment_run, file_name, sheet_name, structure_type, headers,
    data_range, header_range, use_header, rows_count, process,
    contractor_number, contractor_name, wholesaler_hq,
    wholesaler_branch_number, wholesaler_branch_name,
    vendor, customer_name, sales_order_number, ordered_on,
    item_sku, item_sku_alt, item_sku_category, item_upc, item_description,
    unit_price, ship_quantity, uom, extended_price,
    filter, segment_hint

Column value semantics
----------------------
  "Column Name"   double-quoted  → Excel column ref (header mode)
  B / AB          bare letter(s) → Excel column ref (no-header mode)
  NULL / empty                   → NULL::<type>
  $$literal$$     already dollar-quoted → used as-is, never re-wrapped
  anything else (static fields)  → dollar-quoted literal

Rows where process != TRUE are skipped.

Validation checks (per file loaded):
  HARD FAIL  - Empty rows (all data fields NULL)
  HARD FAIL  - Missing item_description (required field)
  HARD FAIL  - Junk-only values (dashes, equals, asterisks) in any column
  WARNING    - Sparse columns (NULL in some rows but not all)
  WARNING    - Rows skipped due to blank/filter (raw range vs loaded count)

Rows are kept on failure so you can inspect them in DataGrip.

Usage:
    python base_mapping_qc.py [--db /path/to/your/database.duckdb]
"""

import argparse
import csv
import html
import io
import os
import pickle
import re
import sys
import textwrap
from collections import defaultdict
from pathlib import Path

import duckdb
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from pipeline_helpers import find_month_folder, find_mapping_sheet_id, sanitize_ampersand_sheets

# =============================================================================
# VS CODE CONFIG
# =============================================================================
# Leave VSCODE_SHEET_ID blank to auto-detect this month's mapping doc from
# Drive. Set it to a specific sheet ID to bypass auto-detection entirely
# (e.g. to re-run an older month, or as a fallback if Drive lookup breaks).
VSCODE_SHEET_ID     = ""  ## leave blank for auto-detect; set to override
VSCODE_DATA_DIR     = "/Users/lorimartella/Documents/gmatter/charlotte_pipe/unspecified/done"
# Written next to this script itself (wherever that currently lives) — this
# is also where this script reads its own generated .sql files back from a
# few lines later, so the two can never get out of sync.
VSCODE_OUTPUT_DIR   = str(Path(__file__).parent / "sql_output")
VSCODE_CREATE_TABLE = False

# DuckDB database this pipeline loads into.
DEFAULT_DB = "/Users/lorimartella/Documents/gmatter/charlotte_pipe/charlotte_pipe.duckdb"

# Drive folder containing the YYYY-MM month folders, each holding this
# month's base + credit mapping docs (searched recursively, since the docs
# may be nested a level deeper, e.g. under a per-payment-run subfolder).
MAPPING_ROOT_FOLDER_ID  = "1pkSdlq7gYLPvaJIWHCjI7kpmOhTcJsEt"
# Filename must contain this keyword to be picked as the BASE mapping doc...
MAPPING_INCLUDE_KEYWORD = "unspecified"
# ...and must NOT contain this keyword (so the credit doc is never mistaken
# for the base doc if both happen to mention "unspecified" somewhere).
MAPPING_EXCLUDE_KEYWORD = "credit"
# =============================================================================

# ---------------------------------------------------------------------------
# Schema: output field → (duckdb_type, macro_or_None)
# ---------------------------------------------------------------------------
FIELD_SCHEMA: dict[str, tuple[str, str | None]] = {
    "transaction_id":           ("uuid",            None),
    "wholesaler_hq":            ("text",            None),
    "wholesaler_branch_name":   ("text",            None),
    "wholesaler_branch_number": ("text",            None),
    "contractor_name":          ("text",            None),
    "contractor_number":        ("text",            None),
    "customer_name":            ("text",            None),
    "vendor":                   ("text",            None),
    "sales_order_number":       ("text",            None),
    "ordered_on":               ("date",            "tinderfy"),
    "item_sku":                 ("text",            None),
    "item_sku_alt":             ("text",            None),
    "item_sku_category":        ("text",            None),
    "item_upc":                 ("text",            None),
    "item_description":         ("text",            None),
    "unit_price":               ("decimal(10, 2)",  "sanitize_amount"),
    "ship_quantity":            ("int",             "sanitize_quantity"),
    "uom":                      ("text",            None),
    "extended_price":           ("decimal(10, 2)",  "sanitize_amount"),
    "row_id":                   ("int",             None),
    "type":                     ("text",            None),
    "sheet_name":               ("text",            None),
    "archive_file_name":        ("text",            None),
}

OUTPUT_FIELDS = list(FIELD_SCHEMA.keys())

ALWAYS_STATIC = {
    "wholesaler_hq", "wholesaler_branch_name", "wholesaler_branch_number",
    "contractor_name", "contractor_number", "type",
}

AUTO_FIELDS = {"transaction_id", "row_id", "type", "sheet_name", "archive_file_name"}

# ---------------------------------------------------------------------------
# Value helpers
# ---------------------------------------------------------------------------

def is_null(val: str) -> bool:
    return val is None or val.strip().upper() in ("", "NULL", "N/A", "NA")


def is_quoted_col(val: str) -> bool:
    v = val.strip()
    return v.startswith('"') and v.endswith('"') and len(v) >= 3


def is_column_letter(val: str) -> bool:
    return bool(re.fullmatch(r"[A-Z]{1,2}", val.strip().upper()))


def col_ref(val: str) -> str:
    v = val.strip()
    if is_quoted_col(v):
        return v
    if is_column_letter(v):
        return f'"{v.upper()}"'
    return f'"{v}"'


def dq(val: str) -> str:
    return f"$${val}$$"


def is_dollar_quoted(val: str) -> bool:
    v = val.strip()
    return v.startswith("$$") and v.endswith("$$") and len(v) >= 4


def safe_dq(val: str) -> str:
    """Like dq(), but passes through values that are already dollar-quoted
    (e.g. a mapping cell that already reads '$$Some Literal$$') instead of
    wrapping them again — double-wrapping produces four dollar signs in a
    row, which DuckDB parses as an empty string followed by bare, invalid
    SQL tokens."""
    return val.strip() if is_dollar_quoted(val.strip()) else dq(val)


def clean_filter(raw: str) -> str:
    if is_null(raw):
        return ""
    v = raw.strip()
    v = re.sub(r"(?i)^\s*WHERE\s+", "", v).strip()
    return v


# ---------------------------------------------------------------------------
# SQL rendering
# ---------------------------------------------------------------------------

ALIGN = 68

def render_field(field: str, mapping_val: str, db_type: str,
                 macro: str | None) -> str:
    if field in AUTO_FIELDS:
        expr = mapping_val
        return f"  {expr}  AS {field}"
    elif field in ALWAYS_STATIC:
        if is_null(mapping_val):
            expr = f"NULL::{db_type}"
        else:
            expr = f"{safe_dq(mapping_val)}::{db_type}"
    else:
        if is_null(mapping_val):
            expr = f"NULL::{db_type}"
        elif is_dollar_quoted(mapping_val):
            # Hardcoded literal already wrapped in $$...$$ in the mapping doc
            # (e.g. $$Acme Corp$$) — use as-is, don't wrap again.
            expr = f"{mapping_val.strip()}::{db_type}"
        elif is_quoted_col(mapping_val) or is_column_letter(mapping_val):
            ref = col_ref(mapping_val)
            inner = f"trim({ref})"
            expr = f"{macro}({inner})::{db_type}" if macro else f"{inner}::{db_type}"
        else:
            expr = f"{safe_dq(mapping_val)}::{db_type}" if not macro \
                   else f"{macro}({safe_dq(mapping_val)})::{db_type}"

    return f"  {expr:<{ALIGN}}AS {field}"


def build_read_xlsx(file_path: str, sheet: str, data_range: str,
                    use_header: bool) -> str:
    lines = [
        f"FROM read_xlsx(",
        f"   {dq(file_path)}",
        f"  ,sheet={dq(sheet)}",
        f"  ,stop_at_empty=false",
        f"  ,header={'TRUE' if use_header else 'FALSE'}",
        f"  ,all_varchar=true",
    ]
    rng = data_range.strip() if data_range else ""
    if rng and rng.upper() not in ("", "NULL"):
        lines.append(f"  ,range={dq(rng)}")
    lines.append(")")
    return "\n".join(lines)


def generate_block(row: dict, file_path: str, row_num: int,
                    read_sheet_name: str | None = None) -> str:
    """Build one CREATE TEMP TABLE + INSERT block for a single mapping row.

    read_sheet_name, if given, is the (possibly renamed) sheet name to pass
    to read_xlsx() — used when the workbook had to be sanitized for a "&" in
    the sheet name (see pipeline_helpers.sanitize_ampersand_sheets). The DB's
    sheet_name column always records the true original name from the
    mapping doc, regardless.
    """
    file_name_only = os.path.basename(file_path)
    # Unescape again defensively — if the mapping cell was double-escaped
    # at the source (e.g. "L&amp;amp;S ..."), one unescape pass in
    # load_mapping() only removes one layer. This second pass is a no-op
    # on an already-clean value.
    sheet        = html.unescape(row.get("sheet_name", "").strip())
    data_range   = row.get("data_range", "").strip()
    header_range = row.get("header_range", "").strip()
    use_header   = row.get("use_header", "TRUE").strip().upper() not in ("FALSE","0","NO","F")
    where_pred   = clean_filter(row.get("filter", ""))
    tmp          = f"_staging_{row_num}"

    read_range = data_range
    if use_header and header_range and data_range:
        m_hdr = re.match(r"[A-Z]+(\d+)", header_range.strip())
        m_dat = re.match(r"([A-Z]+)\d+:(.+)", data_range.strip())
        if m_hdr and m_dat:
            hdr_row   = m_hdr.group(1)
            start_col = m_dat.group(1)
            end_part  = m_dat.group(2)
            read_range = f"{start_col}{hdr_row}:{end_part}"

    lines = []
    excel_sourced = []
    for field in OUTPUT_FIELDS:
        db_type, macro = FIELD_SCHEMA[field]

        if field == "transaction_id":
            mv = "uuid()"
        elif field == "row_id":
            mv = "ROW_NUMBER() OVER ()"
        elif field == "type":
            mv = f"$$unspecified$$::{db_type}"
        elif field == "sheet_name":
            mv = f"{dq(sheet)}::{db_type}"
        elif field == "archive_file_name":
            mv = f"{dq(file_name_only)}::{db_type}"
        else:
            mv = row.get(field, "") or ""
            if field not in ALWAYS_STATIC and not is_null(mv) and not is_dollar_quoted(mv):
                excel_sourced.append(field)

        lines.append(render_field(field, mv, db_type, macro))

    BLANK_ROW_EXCLUDE = {"item_description"}
    blank_row_cols = [c for c in excel_sourced if c not in BLANK_ROW_EXCLUDE]
    if blank_row_cols:
        blank_row_filter = "NOT (" + " AND ".join(
            f"trim({col}::varchar) = ''" if col in ("unit_price", "ship_quantity", "extended_price", "ordered_on")
            else f"{col} IS NULL"
            for col in blank_row_cols
        ) + ")"
    else:
        blank_row_filter = ""

    if where_pred and blank_row_filter:
        full_where = f"\n        WHERE ({where_pred})\n          AND {blank_row_filter}"
    elif where_pred:
        full_where = f"\n        WHERE {where_pred}"
    elif blank_row_filter:
        full_where = f"\n        WHERE {blank_row_filter}"
    else:
        full_where = ""

    select_body = "\n".join(
        (f"   {l.lstrip()}" if i == 0 else f"  ,{l.lstrip()}")
        for i, l in enumerate(lines)
    )

    read_call = build_read_xlsx(
        file_path,
        read_sheet_name if read_sheet_name is not None else sheet,
        read_range, use_header,
    )

    return textwrap.dedent(f"""\
        -- -----------------------------------------------------------------------
        -- Row {row_num}: {file_name_only}
        -- Sheet: {sheet}  |  range: {data_range}  |  header: {use_header}
        -- -----------------------------------------------------------------------

        CREATE OR REPLACE TEMP TABLE {tmp}_raw_count AS
        SELECT COUNT(*) AS raw_count
        {read_call}
        ;

        CREATE OR REPLACE TEMP TABLE {tmp} AS
        SELECT
{select_body}
        {read_call}{full_where}
        ;

        CREATE OR REPLACE TEMP TABLE {tmp}_meta AS
        SELECT
           {dq(file_name_only)}                    AS archive_file_name
          ,{dq(sheet)}                             AS sheet_name
          ,(SELECT raw_count FROM {tmp}_raw_count) AS raw_rows
          ,(SELECT COUNT(*) FROM {tmp})            AS loaded_rows
          ,{str(bool(where_pred)).upper()}::boolean AS has_explicit_filter
        ;

        INSERT INTO transaction_mapping_base
        SELECT * FROM {tmp}
        ;

        DROP TABLE {tmp};
        DROP TABLE {tmp}_raw_count;

    """)


# ---------------------------------------------------------------------------
# SQL preamble
# ---------------------------------------------------------------------------

PREAMBLE = textwrap.dedent("""\
    INSTALL excel;
    LOAD excel;

    CREATE OR REPLACE MACRO tinderfy(_str) AS
    CASE
      WHEN nullif(trim(_str), '') IS NULL
        THEN NULL::date
      WHEN try(trim(_str)::int) IS NOT NULL
        THEN excel_text(trim(_str)::int, 'yyyy-mm-dd')::date
      WHEN try(strptime(trim(_str), '%m/%d/%y')) IS NOT NULL
        THEN strptime(trim(_str), '%m/%d/%y')::date
      WHEN try(strptime(trim(_str), '%m/%d/%Y')) IS NOT NULL
        THEN strptime(trim(_str), '%m/%d/%Y')::date
      WHEN try(trim(_str)::date) IS NOT NULL
        THEN trim(_str)::date
      ELSE NULL
    END
    ;

    CREATE OR REPLACE MACRO sanitize_amount(_str) AS
    CASE
      WHEN nullif(trim(_str), '') IS NULL
        THEN NULL::decimal(10, 2)
      WHEN try(_str::decimal(10, 2)) IS NOT NULL
        THEN _str::decimal(10, 2)
      WHEN try(regexp_replace(_str, '(\\*|,|#VALUE!|#DIV/0!)', '0', 'g')::decimal(10, 2)) IS NOT NULL
        THEN nullif(regexp_replace(_str, '(\\*|,|#VALUE!|#DIV/0!)', '', 'g'), '')::decimal(10, 2)
      ELSE 'unknown format'
    END
    ;

    CREATE OR REPLACE MACRO sanitize_quantity(_str) AS
    CASE
      WHEN nullif(trim(_str), '') IS NULL
        THEN NULL::int
      WHEN try(_str::int) IS NOT NULL
        THEN _str::int
      WHEN try(regexp_replace(_str, '(\\*|,|ea|ft|pc)', '', 'gi')::int) IS NOT NULL
        THEN regexp_replace(_str, '(\\*|,|ea|ft|pc)', '', 'gi')::int
      ELSE 'unknown format'
    END
    ;

""")

CREATE_TABLE_SQL = textwrap.dedent("""\
    CREATE TABLE IF NOT EXISTS transaction_mapping_base (
        transaction_id           UUID              NOT NULL,
        wholesaler_hq            TEXT,
        wholesaler_branch_name   TEXT,
        wholesaler_branch_number TEXT,
        contractor_name          TEXT,
        contractor_number        TEXT,
        customer_name            TEXT,
        vendor                   TEXT,
        sales_order_number       TEXT,
        ordered_on               DATE,
        item_sku                 TEXT,
        item_sku_alt             TEXT,
        item_sku_category        TEXT,
        item_upc                 TEXT,
        item_description         TEXT,
        unit_price               DECIMAL(10, 2),
        ship_quantity            INTEGER,
        uom                      TEXT,
        extended_price           DECIMAL(10, 2),
        row_id                   INTEGER,
        type                     TEXT,
        sheet_name               TEXT,
        archive_file_name        TEXT,
        PRIMARY KEY (transaction_id)
    );
""")

BASE_DIR    = Path(__file__).parent
MAPPING_TAB = "info"
# drive.readonly is required to look up the month folder / mapping doc by
# name; spreadsheets.readonly is required to read the mapping doc's values.
SCOPES      = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]


def get_credentials() -> Credentials:
    creds      = None
    # Scoped to this script specifically so it never collides with
    # credit_mapping_qc.py's cached token if both live in the same folder.
    token_path = BASE_DIR / "token_base.pkl"
    creds_path = BASE_DIR / "oauth_desktop_app.json"

    if token_path.exists():
        with open(token_path, "rb") as f:
            creds = pickle.load(f)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow  = InstalledAppFlow.from_client_secrets_file(str(creds_path), SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, "wb") as f:
            pickle.dump(creds, f)

    return creds


def load_mapping(sheet_id: str) -> list[dict]:
    print(f"  Reading sheet {sheet_id!r}, tab {MAPPING_TAB!r} ...")
    creds   = get_credentials()
    service = build("sheets", "v4", credentials=creds)
    result  = (
        service.spreadsheets().values()
        .get(spreadsheetId=sheet_id, range=f"{MAPPING_TAB}")
        .execute()
    )
    rows = result.get("values", [])
    if not rows:
        raise ValueError(f"Tab '{MAPPING_TAB}' in sheet {sheet_id!r} is empty.")

    headers = [h.strip() for h in rows[0]]
    return [
        {headers[i]: (html.unescape(cell.strip()) if cell else "")
         for i, cell in enumerate(row)}
        for row in rows[1:]
        if any(cell.strip() for cell in row)
    ]


def resolve_sheet_id() -> str:
    """Return the mapping sheet ID to use: the manual override if set,
    otherwise this month's mapping doc auto-detected from Drive."""
    if VSCODE_SHEET_ID:
        print(f"DEBUG: VSCODE_SHEET_ID override set — using {VSCODE_SHEET_ID!r} "
              f"instead of auto-detecting from Drive")
        return VSCODE_SHEET_ID

    creds         = get_credentials()
    drive_service = build("drive", "v3", credentials=creds)

    folder_id, folder_name = find_month_folder(drive_service, MAPPING_ROOT_FOLDER_ID)
    sheet_id, sheet_name = find_mapping_sheet_id(
        drive_service, folder_id,
        include_keyword=MAPPING_INCLUDE_KEYWORD,
        exclude_keyword=MAPPING_EXCLUDE_KEYWORD,
    )
    print(f"DEBUG: auto-detected mapping doc for '{folder_name}': "
          f"'{sheet_name}' ({sheet_id})")
    return sheet_id


# ---------------------------------------------------------------------------
# Validation (used against the loaded rows, after generate + load)
# ---------------------------------------------------------------------------

DATA_COLUMNS = [
    "customer_name", "vendor", "sales_order_number", "ordered_on",
    "item_sku", "item_sku_alt", "item_sku_category", "item_upc",
    "item_description", "unit_price", "ship_quantity", "uom", "extended_price",
]

TEXT_COLUMNS = [
    "customer_name", "vendor", "sales_order_number",
    "item_sku", "item_sku_alt", "item_sku_category", "item_upc",
    "item_description", "uom",
]

JUNK_PATTERN = r"^[\-=\*\s]+$"


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
        SELECT row_id FROM transaction_mapping_base
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
    missing_desc = con.execute("""
        SELECT row_id FROM transaction_mapping_base
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
            SELECT row_id, {col} FROM transaction_mapping_base
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
    total = con.execute("""
        SELECT COUNT(*) FROM transaction_mapping_base WHERE archive_file_name = ?
    """, [archive_file_name]).fetchone()[0]
    if total > 0:
        for col in DATA_COLUMNS:
            null_count = con.execute(f"""
                SELECT COUNT(*) FROM transaction_mapping_base
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


# ---------------------------------------------------------------------------
# Main: generate, then immediately load + validate
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="Generate + load + validate the base pipeline")
    ap.add_argument("--db", default=DEFAULT_DB, help="Path to your DuckDB database file")
    args, _unused = ap.parse_known_args()

    print("DEBUG: script started")

    # ------------------------------------------------------------------
    # GENERATE
    # ------------------------------------------------------------------
    sheet_id     = resolve_sheet_id()
    data_dir     = Path(VSCODE_DATA_DIR).resolve()
    output_dir   = Path(VSCODE_OUTPUT_DIR) if VSCODE_OUTPUT_DIR else Path(__file__).parent / "sql_output"
    create_table = VSCODE_CREATE_TABLE

    print(f"DEBUG: sheet_id={sheet_id}")
    print(f"DEBUG: data_dir={data_dir}")
    print(f"DEBUG: output_dir={output_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"DEBUG: output dir created/confirmed")

    print(f"DEBUG: calling load_mapping...")
    rows = load_mapping(sheet_id)
    print(f"DEBUG: load_mapping returned {len(rows)} rows")
    if not rows:
        print("ERROR: mapping file is empty.", file=sys.stderr)
        sys.exit(1)

    if create_table:
        ct = output_dir / "00_create_table.sql"
        ct.write_text(CREATE_TABLE_SQL, encoding="utf-8")
        print(f"  Wrote {ct}")

    by_file: dict[str, list[tuple[int, dict]]] = defaultdict(list)
    for i, row in enumerate(rows, start=2):
        print(f"DEBUG: row {i} — process={row.get('process','MISSING')!r}  file={row.get('file_name','MISSING')[:50]!r}")
        process = row.get("process", "TRUE").strip().upper()
        if process not in ("TRUE", "1", "YES", "Y"):
            print(f"  SKIP row {i}: process={process!r} — {row.get('file_name','')[:60]}")
            continue

        fname = row.get("file_name", "").strip()
        if not fname:
            print(f"  WARNING row {i}: no file_name, skipping.", file=sys.stderr)
            continue

        by_file[fname].append((i, row))
    print(f"DEBUG: {len(by_file)} unique file(s) found")

    written = 0
    for fname, row_pairs in sorted(by_file.items()):
        file_path = data_dir / fname

        if not file_path.exists():
            print(f"  WARNING: file not found (SQL still generated): {file_path}",
                  file=sys.stderr)

        # Work around DuckDB being unable to match sheet names containing
        # "&" — sanitize once per source file and reuse for all its rows.
        read_file_path = file_path
        rename_map: dict[str, str] = {}
        if file_path.exists():
            sanitized_cache_dir = output_dir / "_sanitized_xlsx_cache"
            read_file_path, rename_map = sanitize_ampersand_sheets(
                file_path, sanitized_cache_dir
            )

        parts = [
            PREAMBLE,
            f"-- =======================================================================\n",
            f"-- Source file : {fname}\n",
            f"-- Sheet(s)    : {', '.join(r.get('sheet_name','') for _, r in row_pairs)}\n",
            f"-- Generated by: base_mapping_qc.py\n",
            f"-- =======================================================================\n\n",
        ]

        for row_num, row in row_pairs:
            original_sheet = html.unescape(row.get("sheet_name", "").strip())
            read_sheet = rename_map.get(original_sheet, original_sheet)
            parts.append(generate_block(row, str(read_file_path), row_num,
                                          read_sheet_name=read_sheet))

        sql_content = "".join(parts)

        safe = re.sub(r"[^\w.\-]", "_", Path(fname).stem)
        out  = output_dir / f"{safe}.sql"
        out.write_text(sql_content, encoding="utf-8")
        sheets = [r.get("sheet_name","") for _, r in row_pairs]
        print(f"  Wrote {out.name}  ({len(row_pairs)} sheet(s): {', '.join(sheets)})")
        written += 1

    print(f"\nDone. {written} SQL file(s) → {output_dir}/")

    # ------------------------------------------------------------------
    # LOAD + VALIDATE
    # ------------------------------------------------------------------
    db_path = Path(args.db)
    sql_dir = output_dir

    sql_files = sorted(
        f for f in sql_dir.glob("*.sql")
        if f.name != "00_create_table.sql"
    )

    if not sql_files:
        print(f"No SQL files found in {sql_dir}")
        return

    print(f"\nConnecting to {db_path} ...")
    con = duckdb.connect(str(db_path))

    before = con.execute("SELECT COUNT(*) FROM transaction_mapping_base").fetchone()[0]
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
        statements = [s.strip() for s in sql.split(";") if s.strip()]
        try:
            count_before = con.execute(
                "SELECT COUNT(*) FROM transaction_mapping_base"
            ).fetchone()[0]
            for stmt in statements:
                con.execute(stmt)
            count_after = con.execute(
                "SELECT COUNT(*) FROM transaction_mapping_base"
            ).fetchone()[0]
            rows_added = count_after - count_before
            print(f"  ✓ Loaded  ({rows_added:,} rows inserted)")
        except Exception as e:
            print(f"  ✗ LOAD FAILED")
            for line in str(e).splitlines():
                print(f"      {line}")
            load_failed.append((sql_file.name, str(e)))
            get_meta(con)  # clean up meta table even on failure
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
            # Only warn about skipped rows when there is no explicit filter in
            # the mapping doc — if a filter exists, skipped rows are intentional
            if skipped > 0 and not has_explicit_filter:
                msg = (f"'{sheet_name}': {skipped} row(s) skipped "
                       f"({loaded_rows:,} loaded from {raw_rows:,} rows in range)")
                print(f"  ⚠  Skipped rows: {msg}")
                file_warnings.append({"check": "Skipped rows", "detail": msg})

        # ── Step 3: Identify archive_file_name ────────────────────────
        archive_file_name = con.execute("""
            SELECT DISTINCT archive_file_name
            FROM transaction_mapping_base
            ORDER BY rowid DESC
            LIMIT 1
        """).fetchone()[0]

        # ── Step 4: Validate data quality ─────────────────────────────
        failures, warnings = validate(con, archive_file_name)
        file_warnings.extend(warnings)

        if file_warnings:
            all_warnings.append((sql_file.name, file_warnings))
            if not any(w["check"] == "Skipped rows" for w in file_warnings) or len(file_warnings) > 1:
                for w in file_warnings:
                    if w["check"] != "Skipped rows":  # already printed above
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
    after = con.execute("SELECT COUNT(*) FROM transaction_mapping_base").fetchone()[0]
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