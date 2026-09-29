"""
Storage-agnostic logic shared by run_scenario.py (CSV) and
run_scenario_sheets.py (Google Sheets), so both back ends compute and label
results identically -- only how rows are READ and WRITTEN differs between
the two scripts.
"""

from datetime import datetime

from config import CATEGORIES

RUNNING_TOTAL_COLUMNS = [f"running_{cat}" for cat in CATEGORIES]

RESULT_COLUMNS = [
    "cum_after", "earning_quantity", "actual_earnings", "effective_rate",
    "actual_level", "final_level", "category_mode", "earnings_passed", "level_passed",
] + RUNNING_TOTAL_COLUMNS

EARNINGS_TOLERANCE = 0.01  # allow a penny of rounding slack when comparing

DATE_FORMATS_TO_TRY = [
    "%Y-%m-%d",   # 2026-09-20  (what the script writes / prefers)
    "%m/%d/%Y",   # 9/20/2026
    "%m/%d/%y",   # 9/20/26     (Excel's usual auto-reformat)
    "%Y/%m/%d",   # 2026/09/20
    "%m-%d-%Y",   # 9-20-2026
    "%m-%d-%y",   # 9-20-26
]


def parse_date(raw_value: str, row_num: int):
    """Excel/Sheets like to silently reformat a 'date-looking' column (e.g.
    2026-09-20 -> 9/20/26) -- even for rows you didn't touch. Rather than
    break on that, try the formats they commonly produce before giving up."""
    raw_value = (raw_value or "").strip()
    for fmt in DATE_FORMATS_TO_TRY:
        try:
            return datetime.strptime(raw_value, fmt).date()
        except ValueError:
            continue
    raise ValueError(
        f"Row {row_num}: couldn't parse date {raw_value!r}. "
        f"Tried formats: {DATE_FORMATS_TO_TRY}. "
        "Tip: format the date column as Text (or yyyy-mm-dd) to stop "
        "Excel/Sheets from auto-reformatting it on save."
    )


def parse_amount(raw_value, row_num: int) -> float:
    """Tolerates currency formatting a person might type/paste in directly:
    '$50,000', '50000', ' 1200.50 ', etc."""
    s = str(raw_value).strip()
    cleaned = s.replace("$", "").replace(",", "").strip()
    if cleaned == "":
        raise ValueError(f"Row {row_num}: amount is empty.")
    try:
        return float(cleaned)
    except ValueError:
        raise ValueError(f"Row {row_num}: couldn't parse amount {s!r} as a number.")


def final_tier_of(results) -> str | None:
    """The tier as of the CHRONOLOGICALLY LAST transaction in this batch of
    results (a single scenario's worth) -- ties on the same date broken by
    original list order, matching the engine's own internal sort. This is
    the single value level_passed checks every row against, since a
    scenario's expected_level is meant to be its final, settled answer --
    not what was true at each date along the way (see actual_level for
    that point-in-time story)."""
    if not results:
        return None
    chrono_order = sorted(range(len(results)), key=lambda i: (results[i].date, i))
    return results[chrono_order[-1]].active_tier_after


def compare_earnings(raw_row: dict, actual: float) -> str:
    """Returns 'TRUE'/'FALSE'/'' (blank = not checked, expected left empty)."""
    raw = str(raw_row.get("expected_earnings") or "").strip()
    if raw == "":
        return ""
    try:
        expected = float(raw.replace("$", "").replace(",", ""))
    except ValueError:
        return ""  # not a usable number -- treat as not-yet-filled-in
    return "TRUE" if abs(actual - expected) <= EARNINGS_TOLERANCE else "FALSE"


def compare_level(raw_row: dict, level_to_check) -> str:
    """Returns 'TRUE'/'FALSE'/'' (blank = not checked, expected left empty).
    Case-insensitive; a level_to_check of None is compared as the literal
    string 'none', so typing 'none' as your expected_level is valid too."""
    raw = str(raw_row.get("expected_level") or "").strip()
    if raw == "":
        return ""
    actual_str = (level_to_check or "none").strip().lower()
    return "TRUE" if raw.lower() == actual_str else "FALSE"


def build_output_rows(raw_rows: list[dict], fieldnames: list[str], results) -> tuple[list[str], list[dict]]:
    """Returns (out_fieldnames, out_rows) -- original columns preserved and
    in order, with the computed RESULT_COLUMNS appended. Storage-agnostic:
    works the same whether raw_rows came from a CSV or a Sheet.

    IMPORTANT: raw_rows/results must be ONE SCENARIO'S WORTH of rows -- this
    function computes ONE final_tier for the whole batch it's given (via
    final_tier_of) and checks every row's expected_level against that same
    value. Call it once per scenario group, not once across multiple
    scenarios combined, or final_level will be meaningless."""
    final_tier = final_tier_of(results)
    out_fieldnames = fieldnames + [c for c in RESULT_COLUMNS if c not in fieldnames]
    out_rows = []
    for raw_row, r in zip(raw_rows, results):
        out_row = dict(raw_row)
        out_row["cum_after"] = round(r.cum_after, 2)
        out_row["earning_quantity"] = round(r.earning_quantity, 2)
        out_row["actual_earnings"] = round(r.points_earned, 4)
        out_row["effective_rate"] = (
            round(r.effective_rate, 6) if r.effective_rate is not None else ""
        )
        out_row["actual_level"] = r.active_tier_after or ""
        out_row["final_level"] = final_tier or ""
        out_row["category_mode"] = r.category_mode
        out_row["earnings_passed"] = compare_earnings(raw_row, r.points_earned)
        out_row["level_passed"] = compare_level(raw_row, final_tier)
        for cat in CATEGORIES:
            out_row[f"running_{cat}"] = round(r.running_totals[cat], 2)
        out_rows.append(out_row)
    return out_fieldnames, out_rows


def scenario_summary_lines(raw_rows, results) -> list[str]:
    """Same math as print_summary, minus the 'Wrote N rows to <dest>' line --
    for printing one block per scenario rather than once for a whole file/sheet."""
    lines = []
    total_points = sum(r.points_earned for r in results)
    final_tier = final_tier_of(results)
    lines.append(f"  Total points earned: {total_points:,.2f}")
    lines.append(f"  Final active tier: {final_tier}")

    earn_checks = [compare_earnings(rr, r.points_earned) for rr, r in zip(raw_rows, results)]
    level_checks = [compare_level(rr, final_tier) for rr in raw_rows]
    earn_checked = [c for c in earn_checks if c != ""]
    level_checked = [c for c in level_checks if c != ""]
    if earn_checked:
        failed = earn_checked.count("FALSE")
        lines.append(f"  Earnings checks: {len(earn_checked) - failed}/{len(earn_checked)} passed"
                      + (f" -- {failed} FAILED" if failed else ""))
    if level_checked:
        failed = level_checked.count("FALSE")
        lines.append(f"  Level checks: {len(level_checked) - failed}/{len(level_checked)} passed"
                      + (f" -- {failed} FAILED" if failed else ""))
    return lines


def print_summary(raw_rows, results, destination_label: str) -> None:
    total_points = sum(r.points_earned for r in results)
    final_tier = final_tier_of(results)
    print(f"Wrote {len(results)} rows to {destination_label}")
    print(f"Total points earned: {total_points:,.2f}")
    print(f"Final active tier: {final_tier}")

    earn_checks = [compare_earnings(rr, r.points_earned) for rr, r in zip(raw_rows, results)]
    level_checks = [compare_level(rr, final_tier) for rr in raw_rows]
    earn_checked = [c for c in earn_checks if c != ""]
    level_checked = [c for c in level_checks if c != ""]
    if earn_checked:
        failed = earn_checked.count("FALSE")
        print(f"Earnings checks: {len(earn_checked) - failed}/{len(earn_checked)} passed"
              + (f" -- {failed} FAILED" if failed else ""))
    if level_checked:
        failed = level_checked.count("FALSE")
        print(f"Level checks: {len(level_checked) - failed}/{len(level_checked)} passed"
              + (f" -- {failed} FAILED" if failed else ""))