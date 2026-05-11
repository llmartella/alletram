#!/usr/bin/env python3
"""
mapping_qc_generate_sql.py  v2
-------------------
Reads a mapping TSV/CSV (exported from Google Sheets) and generates one DuckDB
SQL file per Excel file referenced in the mapping.

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
  anything else (static fields)  → dollar-quoted literal

Rows where process != TRUE are skipped.
"""

import argparse
import csv
import io
import os
import pickle
import re
import sys
import textwrap
from collections import defaultdict
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# =============================================================================
# VS CODE CONFIG
# =============================================================================
VSCODE_SHEET_ID     = "1riKNLl_6Uqaoed2V1HKHn168HmAvcXc4V1uvGRe15-E"
VSCODE_DATA_DIR     = "/Users/lorimartella/Documents/gmatter/charlotte_pipe/unspecified/done"
VSCODE_OUTPUT_DIR   = "/Users/lorimartella/Documents/gmatter/charlotte_pipe/cpf_python_scripts/sql_output"
VSCODE_CREATE_TABLE = False
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
            expr = f"{dq(mapping_val)}::{db_type}"
    else:
        if is_null(mapping_val):
            expr = f"NULL::{db_type}"
        elif is_quoted_col(mapping_val) or is_column_letter(mapping_val):
            ref = col_ref(mapping_val)
            inner = f"trim({ref})"
            expr = f"{macro}({inner})::{db_type}" if macro else f"{inner}::{db_type}"
        else:
            expr = f"{dq(mapping_val)}::{db_type}" if not macro \
                   else f"{macro}({dq(mapping_val)})::{db_type}"

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


def generate_block(row: dict, file_path: str, row_num: int) -> str:
    file_name_only = os.path.basename(file_path)
    sheet        = row.get("sheet_name", "").strip()
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
            if field not in ALWAYS_STATIC and not is_null(mv):
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

    read_call = build_read_xlsx(file_path, sheet, read_range, use_header)

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
SCOPES      = ["https://www.googleapis.com/auth/spreadsheets.readonly"]


def get_credentials() -> Credentials:
    creds      = None
    token_path = BASE_DIR / "token.pkl"
    creds_path = BASE_DIR / "credentials.json"

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
        {headers[i]: (cell.strip() if cell else "")
         for i, cell in enumerate(row)}
        for row in rows[1:]
        if any(cell.strip() for cell in row)
    ]


def main():
    print("DEBUG: script started")
    print(f"DEBUG VSCODE_SHEET_ID repr: {VSCODE_SHEET_ID!r}")

    # Use VS Code config directly — no argparse needed
    sheet_id     = VSCODE_SHEET_ID
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

        parts = [
            PREAMBLE,
            f"-- =======================================================================\n",
            f"-- Source file : {fname}\n",
            f"-- Sheet(s)    : {', '.join(r.get('sheet_name','') for _, r in row_pairs)}\n",
            f"-- Generated by: mapping_qc_generate_sql.py\n",
            f"-- =======================================================================\n\n",
        ]

        for row_num, row in row_pairs:
            parts.append(generate_block(row, str(file_path), row_num))

        sql_content = "".join(parts)

        safe = re.sub(r"[^\w.\-]", "_", Path(fname).stem)
        out  = output_dir / f"{safe}.sql"
        out.write_text(sql_content, encoding="utf-8")
        sheets = [r.get("sheet_name","") for _, r in row_pairs]
        print(f"  Wrote {out.name}  ({len(row_pairs)} sheet(s): {', '.join(sheets)})")
        written += 1

    print(f"\nDone. {written} SQL file(s) → {output_dir}/")


if __name__ == "__main__":
    main()
