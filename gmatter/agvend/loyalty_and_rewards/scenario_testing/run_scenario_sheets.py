"""
Same engine, same rules, same output columns as run_scenario.py -- but reads
transactions from, and writes results directly back into, a live Google
Sheet via the Sheets API (using a service account), instead of a local CSV
file. This avoids the "I had the file open in Excel and it didn't refresh /
I overwrote the results by saving" problem entirely, since you're always
looking at the live document, not a local copy.

ONE-TIME SETUP (you've already done steps 1-3):
    1. A Google Cloud service account exists, with a credentials JSON key
       file downloaded locally (SERVICE_ACCOUNT_FILE below).
    2. The target Google Sheet has been shared with that service account's
       client_email (found inside the credentials JSON) as an Editor.
    3. `pip install gspread google-auth --break-system-packages`

EVERY TIME YOU RUN:
    1. Add/edit transaction rows directly in the Google Sheet, in your
       browser: date, category, amount, scenario (REQUIRED -- see below),
       txn_id (optional), expected_earnings (optional), expected_level
       (optional), plus any extra columns you want.
    2. Run this script (Run button in VS Code, or python3 run_scenario_sheets.py).
    3. Look at the Sheet in your browser -- the computed columns are filled
       in, right next to your data. No file to close/reopen.

MULTIPLE SCENARIOS IN ONE SHEET:
    Add a 'scenario' column and give every row a scenario value (e.g. 1, 2,
    3 -- text or numbers, doesn't matter). All rows sharing a scenario value
    are simulated together, completely independently of every other
    scenario's rows (fresh running totals, fresh tier, starting from zero --
    as if each scenario were its own separate sheet). This lets you compare
    several what-if timelines side by side in one place.

    REQUIREMENT: each scenario's rows must be CONTIGUOUS -- all of scenario
    1's rows together, then all of scenario 2's rows, etc. If a scenario
    value reappears after a different one came in between (rows out of
    order/interleaved), the script stops with an error rather than guessing
    what you meant. Every row must have a scenario value -- a blank one also
    stops the script with an error naming the row.

    The console output breaks totals out per scenario (points earned, final
    tier, earnings/level check pass rates for that scenario only).

IMPORTANT: this script only touches the columns it computes (see
report.py's RESULT_COLUMNS) -- your input columns, and anything you've
added yourself (extra columns, formulas), are never cleared or rewritten,
as long as they sit to the LEFT of the computed columns. Don't insert a new
column in the middle of the computed block later.
"""

from datetime import datetime

import gspread
from google.oauth2.service_account import Credentials

from engine import Transaction, simulate_scenario
import report


# =============================================================================
# SCENARIO SETTINGS — edit these, then Run
# =============================================================================
SERVICE_ACCOUNT_FILE = "/Users/lorimartella/alletram/gmatter/agvend/loyalty_and_rewards/scenario_testing/credentials.json"
SHEET_ID = "15qTVB9N0CjvxwOvOkX8bYc2Kh0m6OjhnpJ8PfhDgfrY"
SHEET_NAME = "working_scenario"            # tab name within that Sheet
ENROLL_DATE = "2026-09-15"                 # YYYY-MM-DD, or None to treat everything as enrolled
# =============================================================================

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def get_worksheet():
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SHEET_ID)
    return sh.worksheet(SHEET_NAME)


def load_transactions(ws):
    """Returns (transactions, raw_rows, fieldnames), mirroring the CSV
    version's shape so report.py's shared logic works unchanged."""
    all_values = ws.get_all_values()
    if not all_values:
        raise ValueError(
            f"Sheet tab {SHEET_NAME!r} is completely empty -- add a header "
            "row (date, category, amount, txn_id, ...) first."
        )

    fieldnames = all_values[0]
    data_rows = all_values[1:]

    if "date" not in fieldnames or "category" not in fieldnames or "amount" not in fieldnames:
        raise ValueError(
            f"Header row is {fieldnames!r} -- expected at least "
            "'date', 'category', and 'amount' columns."
        )
    if "scenario" not in fieldnames:
        raise ValueError(
            f"Header row is {fieldnames!r} -- expected a 'scenario' column "
            "so each group of transactions can be simulated independently. "
            "Add a 'scenario' column and give every row a scenario value "
            "(e.g. 1, 2, 3)."
        )

    txns = []
    raw_rows = []
    for i, row in enumerate(data_rows):
        # Sheets can return short rows if trailing cells are blank -- pad.
        padded = row + [""] * (len(fieldnames) - len(row))
        row_dict = dict(zip(fieldnames, padded))

        if row_dict.get("date", "").strip() == "" and row_dict.get("category", "").strip() == "":
            continue  # skip fully-blank rows (common at the bottom of a sheet)

        raw_rows.append(row_dict)
        row_num = i + 2  # +2: header row + 1-indexed

        scenario = row_dict.get("scenario", "").strip()
        if scenario == "":
            raise ValueError(
                f"Row {row_num}: 'scenario' is blank. Every transaction row "
                "needs a scenario value (e.g. 1, 2, 3) so it's simulated as "
                "part of the right group."
            )

        txns.append(
            Transaction(
                date=report.parse_date(row_dict["date"], row_num=row_num),
                category=row_dict["category"].strip(),
                amount=report.parse_amount(row_dict["amount"], row_num=row_num),
                txn_id=row_dict.get("txn_id") or f"row{i+1}",
            )
        )
    return txns, raw_rows, fieldnames, len(all_values)


def group_by_scenario(raw_rows, txns):
    """
    Splits raw_rows/txns into independent scenario groups, in the order each
    scenario first appears. Requires each scenario's rows to be CONTIGUOUS
    (all together, not interleaved with another scenario) -- raises a clear
    error if that's violated, since silently reordering could hide a mistake
    (e.g. a stray row typed with the wrong scenario number).

    Returns a list of (scenario_value, indices, group_raw_rows, group_txns)
    tuples, in first-appearance order. `indices` are positions into the
    original raw_rows/txns lists, used to scatter each group's results back
    into overall row order afterward.
    """
    scenario_values = [rr["scenario"].strip() for rr in raw_rows]

    groups: dict[str, list[int]] = {}
    order: list[str] = []
    finished: set[str] = set()
    current = None

    for i, sv in enumerate(scenario_values):
        if sv != current:
            if sv in finished:
                raise ValueError(
                    f"Scenario {sv!r} reappears at row {i + 2} after another "
                    f"scenario's rows came in between. Keep each scenario's "
                    f"rows grouped together (contiguous), not interleaved."
                )
            if current is not None:
                finished.add(current)
            current = sv
            if sv not in groups:
                groups[sv] = []
                order.append(sv)
        groups[sv].append(i)

    return [
        (sv, groups[sv], [raw_rows[i] for i in groups[sv]], [txns[i] for i in groups[sv]])
        for sv in order
    ]


def write_results(ws, out_fieldnames, out_rows, prior_total_rows: int) -> None:
    """
    Only ever touches the columns the script itself computes (report.RESULT_COLUMNS)
    -- never the original input columns. That means anything you've added to
    the left of those (extra columns, manual formulas like a scratch
    cum_after check, etc.) is left completely alone: not cleared, not
    rewritten, not even re-typed with the same value.

    "Protected" columns = whatever's in the header that ISN'T one of the
    script's own result columns. Those are assumed to always come first,
    with the script's managed block appended contiguously after them (which
    is exactly what happens naturally: the first time this runs against a
    fresh sheet, the result columns get appended at the end; every run after
    that, they're already sitting in that same spot).

    prior_total_rows = how many rows (including header) the sheet had BEFORE
    this run, so that if you deleted a transaction row since last time, any
    stale computed values left over below the new data get cleared too --
    without ever touching the protected columns to the left.

    Takes already-built out_fieldnames/out_rows (see main() -- these are
    assembled PER SCENARIO GROUP via report.build_output_rows, then merged
    back into overall row order, so each scenario's final_level is computed
    against only its own rows).
    """
    protected_fieldnames = [c for c in out_fieldnames if c not in report.RESULT_COLUMNS]
    managed_start_col = len(protected_fieldnames) + 1  # 1-indexed
    managed_end_col = managed_start_col + len(report.RESULT_COLUMNS) - 1

    start_a1 = gspread.utils.rowcol_to_a1(1, managed_start_col)
    start_col_letter = start_a1.rstrip("0123456789")
    end_col_letter = gspread.utils.rowcol_to_a1(1, managed_end_col).rstrip("0123456789")

    new_row_count = len(out_rows) + 1  # +1 for header
    clear_through_row = max(prior_total_rows, new_row_count)
    ws.batch_clear([f"{start_col_letter}1:{end_col_letter}{clear_through_row}"])

    values = [report.RESULT_COLUMNS]
    for out_row in out_rows:
        values.append([out_row.get(col, "") for col in report.RESULT_COLUMNS])

    ws.update(
        values=values,
        range_name=start_a1,
        value_input_option=gspread.utils.ValueInputOption.user_entered,
    )


def main():
    enroll_date = None
    if ENROLL_DATE:
        enroll_date = datetime.strptime(ENROLL_DATE, "%Y-%m-%d").date()

    ws = get_worksheet()
    txns, raw_rows, fieldnames, prior_total_rows = load_transactions(ws)

    if not txns:
        print(f"No transaction rows found in {SHEET_NAME!r} -- nothing to do.")
        return

    groups = group_by_scenario(raw_rows, txns)

    # Simulate each scenario independently (simulate_scenario() carries no
    # state across calls, and applies the settled-tier drop rule PER
    # SCENARIO -- a drop in scenario 1 never affects scenario 2's math).
    # Build each scenario's output rows SEPARATELY too (report.build_output_rows
    # computes one final_level per call, from whatever results it's given --
    # calling it once per scenario keeps that final_level scenario-scoped,
    # rather than one final_level bleeding across every scenario in the sheet).
    results = [None] * len(raw_rows)
    out_fieldnames = None
    out_rows_by_index: dict[int, dict] = {}
    for scenario_value, indices, group_raw_rows, group_txns in groups:
        group_results = simulate_scenario(group_txns, enroll_date=enroll_date)
        for idx, r in zip(indices, group_results):
            results[idx] = r
        group_out_fieldnames, group_out_rows = report.build_output_rows(
            group_raw_rows, fieldnames, group_results
        )
        out_fieldnames = group_out_fieldnames  # identical across groups
        for idx, out_row in zip(indices, group_out_rows):
            out_rows_by_index[idx] = out_row

    out_rows = [out_rows_by_index[i] for i in range(len(raw_rows))]
    write_results(ws, out_fieldnames, out_rows, prior_total_rows)

    print(f"Wrote {len(results)} rows to Google Sheet tab '{SHEET_NAME}' "
          f"across {len(groups)} scenario(s)")
    for scenario_value, indices, group_raw_rows, _ in groups:
        group_results = [results[i] for i in indices]
        print(f"\nScenario {scenario_value}:")
        for line in report.scenario_summary_lines(group_raw_rows, group_results):
            print(line)


if __name__ == "__main__":
    main()