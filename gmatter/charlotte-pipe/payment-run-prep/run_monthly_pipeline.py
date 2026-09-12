#!/usr/bin/env python3
"""
run_monthly_pipeline.py
------------------------
One-command runner for the monthly Charlotte Pipe payment-run prep pipeline.

Order of operations:
  1. Clear out old generated SQL files (sql_output, sql_output_credit)
  2. base_mapping_qc.py    -> generates + loads + validates the base pipeline
  3. credit_mapping_qc.py  -> generates + loads + validates the credit pipeline

Both scripts auto-detect this month's mapping document from Drive, so
there's nothing to edit here month-to-month — just run this script.

If a pipeline reports a load or data-quality failure, the other pipeline
still runs (so a base-pipeline issue doesn't block credit from running),
but the final summary will flag it clearly.

Usage:
    python run_monthly_pipeline.py
"""

import re
import subprocess
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent

SQL_OUTPUT_DIRS = [
    SCRIPT_DIR / "sql_output",
    SCRIPT_DIR / "sql_output_credit",
]

STEPS = [
    ("base_mapping_qc.py",   "Base pipeline (generate + load + validate)"),
    ("credit_mapping_qc.py", "Credit pipeline (generate + load + validate)"),
]

_SUMMARY_FIELD_PATTERN = re.compile(
    r"^\s*(Files run|Load failures|Quality failures|Warnings|Rows added)\s*:\s*(.+?)\s*$",
    re.MULTILINE,
)


def clear_sql_output_dirs() -> None:
    print("Clearing previous SQL output...")
    for d in SQL_OUTPUT_DIRS:
        if not d.exists():
            print(f"  (skip) {d} does not exist yet")
            continue
        sql_files = list(d.glob("*.sql"))
        for f in sql_files:
            f.unlink()
        print(f"  Cleared {len(sql_files)} .sql file(s) from {d}")


def _extract_summary_block(output_text: str) -> list[tuple[str, str]] | None:
    """Pull the Files run / Load failures / Quality failures / Warnings /
    Rows added lines out of a run step's own printed SUMMARY block, in the
    order they appear. Returns None if none of these were found (e.g. the
    step crashed before printing a summary)."""
    matches = _SUMMARY_FIELD_PATTERN.findall(output_text)
    return matches or None


def run_step(script_name: str, label: str) -> tuple[bool, str]:
    script_path = SCRIPT_DIR / script_name
    print(f"\n{'=' * 70}")
    print(f"  {label}  ({script_name})")
    print(f"{'=' * 70}")

    if not script_path.exists():
        print(f"  ERROR: {script_path} not found — skipping.")
        return False, ""

    # Stream output live (so you still see progress as it happens) while
    # also capturing it, so the final summary can pull numbers out of it.
    process = subprocess.Popen(
        [sys.executable, str(script_path)],
        cwd=str(SCRIPT_DIR),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
    )
    captured_lines = []
    for line in process.stdout:
        print(line, end="")
        captured_lines.append(line)
    process.wait()

    success = process.returncode == 0
    status = "OK" if success else f"FAILED (exit code {process.returncode})"
    print(f"  -> {status}")
    return success, "".join(captured_lines)


def main():
    clear_sql_output_dirs()

    results = []
    summaries: dict[str, list[tuple[str, str]]] = {}  # label -> summary fields

    for script_name, label in STEPS:
        ok, output_text = run_step(script_name, label)
        results.append((label, ok))
        fields = _extract_summary_block(output_text)
        if fields:
            summaries[label] = fields

    print(f"\n{'=' * 70}")
    print("  PIPELINE SUMMARY")
    print(f"{'=' * 70}")

    any_failed = False
    for label, ok in results:
        symbol = "✓" if ok else "✗"
        print(f"  {symbol} {label}")
        if not ok:
            any_failed = True

    if summaries:
        label_width = max(len(name) for name, _ in next(iter(summaries.values())))
        for label, fields in summaries.items():
            pipeline_name = "credit" if "credit" in label.lower() else "base"
            print(f"\n{pipeline_name}:")
            print(f"{'=' * 60}")
            print("  SUMMARY")
            print(f"{'=' * 60}")
            for name, value in fields:
                print(f"  {name:<{label_width}} : {value}")

    if any_failed:
        print("\n  One or more steps failed or reported quality issues — see output above.")
        sys.exit(1)
    else:
        print("\n  All steps completed successfully.")


if __name__ == "__main__":
    main()