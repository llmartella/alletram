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
but the final summary will flag it clearly, including the names of the
files that failed.

Everything printed to the terminal is also saved to a timestamped log file
in a "logs" folder next to this script (logs/pipeline_YYYYMMDD_HHMMSS.log),
so nothing is lost to terminal scrollback limits. A new log is created on
every run; delete old ones whenever you like.

Usage:
    python run_monthly_pipeline.py
"""

import re
import subprocess
import sys
import traceback
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
LOG_DIR = SCRIPT_DIR / "logs"

SQL_OUTPUT_DIRS = [
    SCRIPT_DIR / "sql_output",
    SCRIPT_DIR / "sql_output_credit",
]

STEPS = [
    ("base_mapping_qc.py",   "Base pipeline (generate + load + validate)"),
    ("credit_mapping_qc.py", "Credit pipeline (generate + load + validate)"),
]

# One "Name : value" line from a pipeline's own SUMMARY block. The value must
# be on the same line — otherwise a section header like "  Load failures:"
# would swallow the line below it.
_SUMMARY_FIELD_PATTERN = re.compile(
    r"^[ \t]*(Files run|Load failures|Quality failures|Warnings|Rows added|Rows replaced)"
    r"[ \t]*:[ \t]*(\S.*?)[ \t]*$",
    re.MULTILINE,
)

# Headers of the detail sections each pipeline prints after its SUMMARY
# block: "  Load failures:", "  Quality failures (...):", "  Warnings:".
_SECTION_HEADER_PATTERN = re.compile(
    r"^  (Load failures:|Quality failures \(.*\):|Warnings:)[ \t]*$",
    re.MULTILINE,
)


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

class _Tee:
    """Writes everything to several streams at once, so the terminal and the
    log file always receive identical output."""

    def __init__(self, *streams):
        self._streams = streams

    def write(self, data):
        for stream in self._streams:
            stream.write(data)
        return len(data)

    def flush(self):
        for stream in self._streams:
            stream.flush()


def _start_logging():
    """Mirror stdout and stderr into logs/pipeline_<timestamp>.log.
    Returns (log_path, log_file)."""
    LOG_DIR.mkdir(exist_ok=True)
    log_path = LOG_DIR / f"pipeline_{datetime.now():%Y%m%d_%H%M%S}.log"
    log_file = open(log_path, "w", encoding="utf-8", buffering=1)
    sys.stdout = _Tee(sys.__stdout__, log_file)
    sys.stderr = _Tee(sys.__stderr__, log_file)
    return log_path, log_file


# ---------------------------------------------------------------------------
# Output parsing
# ---------------------------------------------------------------------------

def _pipeline_name(label: str) -> str:
    return "credit" if "credit" in label.lower() else "base"


def _extract_summary_block(output_text: str) -> list[tuple[str, str]] | None:
    """Pull the Files run / Load failures / Quality failures / Warnings /
    Rows added lines out of a run step's own printed SUMMARY block, in the
    order they appear. Returns None if none of these were found (e.g. the
    step crashed before printing a summary)."""
    matches = _SUMMARY_FIELD_PATTERN.findall(output_text)
    return matches or None


def _extract_failure_details(output_text: str) -> str:
    """Return the 'Load failures' and 'Quality failures' detail sections
    (file names plus the first line of each error) from a pipeline's own
    end-of-run output. Warnings are left out. Returns '' if there are none."""
    start = output_text.rfind("  SUMMARY")
    if start == -1:
        return ""
    tail = output_text[start:]
    headers = list(_SECTION_HEADER_PATTERN.finditer(tail))
    sections = []
    for i, header in enumerate(headers):
        if header.group(1).startswith("Warnings"):
            continue
        end = headers[i + 1].start() if i + 1 < len(headers) else len(tail)
        sections.append(tail[header.start():end].rstrip())
    return "\n\n".join(sections)


def _last_lines(output_text: str, count: int = 8) -> str:
    lines = [line.rstrip() for line in output_text.splitlines() if line.strip()]
    return "\n".join(lines[-count:])


# ---------------------------------------------------------------------------
# Steps
# ---------------------------------------------------------------------------

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


def run_step(script_name: str, label: str) -> tuple[bool, str]:
    script_path = SCRIPT_DIR / script_name
    print(f"\n{'=' * 70}")
    print(f"  {label}  ({script_name})")
    print(f"{'=' * 70}")

    if not script_path.exists():
        message = f"ERROR: {script_path} not found — skipping."
        print(f"  {message}")
        return False, message + "\n"

    # Stream output live (so you still see progress as it happens) while
    # also capturing it, so the final summary can pull details out of it.
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


def _run_pipeline(log_path: Path) -> int:
    print(f"Log file: {log_path}")
    clear_sql_output_dirs()

    results = []
    summaries: dict[str, list[tuple[str, str]]] = {}  # label -> summary fields
    failure_details: dict[str, str] = {}              # label -> failed-file sections
    crash_tails: dict[str, str] = {}                  # label -> last output lines

    for script_name, label in STEPS:
        ok, output_text = run_step(script_name, label)
        results.append((label, ok))
        fields = _extract_summary_block(output_text)
        if fields:
            summaries[label] = fields
            details = _extract_failure_details(output_text)
            if details:
                failure_details[label] = details
        elif not ok:
            # The step stopped before printing its own summary (a crash, a
            # Drive lookup error, a missing script) — keep the end of its
            # output so the cause isn't left behind in the scrollback.
            crash_tails[label] = _last_lines(output_text)

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
            print(f"\n{_pipeline_name(label)}:")
            print(f"{'=' * 60}")
            print("  SUMMARY")
            print(f"{'=' * 60}")
            for name, value in fields:
                print(f"  {name:<{label_width}} : {value}")

    if failure_details or crash_tails:
        print(f"\n{'=' * 70}")
        print("  FAILED FILES / ERRORS")
        print(f"{'=' * 70}")
        for label, details in failure_details.items():
            print(f"\n{_pipeline_name(label)}:")
            print(details)
        for label, tail in crash_tails.items():
            print(f"\n{_pipeline_name(label)} (stopped before finishing — last output):")
            for line in tail.splitlines():
                print(f"  {line}")
        print("\n  Each file's full error text is in the log.")

    if any_failed:
        print("\n  One or more steps failed or reported quality issues — "
              "see the failed files above.")
    else:
        print("\n  All steps completed successfully.")
    print(f"  Full log: {log_path}")

    return 1 if any_failed else 0


def main():
    log_path, log_file = _start_logging()
    try:
        exit_code = _run_pipeline(log_path)
    except Exception:
        traceback.print_exc()  # goes through the tee, so it lands in the log too
        exit_code = 1
    finally:
        sys.stdout.flush()
        sys.stderr.flush()
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__
        log_file.close()
    sys.exit(exit_code)


if __name__ == "__main__":
    main()