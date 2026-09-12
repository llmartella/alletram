"""
pipeline_helpers.py
---------------------
Shared helpers used by both base_mapping_qc.py and credit_mapping_qc.py:

1. Drive mapping-doc lookup — find this month's mapping-document Google
   Sheet inside a year/month folder structure on Google Drive.

   Expected folder layout under a root folder ID:

       <ROOT_FOLDER_ID>/
           2026-01/
               <base mapping doc>      (Google Sheet, name contains e.g. "unspecified")
               <credit mapping doc>    (Google Sheet, name contains e.g. "credit")
           2026-02/
               ...

   Within a given month's folder (searched recursively, since the docs
   aren't always directly inside it), the right doc is told apart from the
   other one purely by filename keyword matching — no manual ID lookup
   needed.

2. Ampersand-sheet sanitizer — DuckDB's read_xlsx() `sheet=` parameter
   appears to XML/HTML-escape its input internally before matching it
   against the workbook's actual sheet names, so a sheet literally named
   e.g. "L&S Pipe Invoices - Stock" can never be found by name. The
   workaround is to save a renamed copy (same filename, separate cache
   folder) with those sheet names changed to be ampersand-free, and point
   read_xlsx() at that copy + renamed sheet instead.
"""

import re
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl

# ---------------------------------------------------------------------------
# openpyxl font-family patch
# ---------------------------------------------------------------------------
# Silences invalid font family values (e.g. 34) that some Template files
# contain, so opening those workbooks to check/rename sheet names doesn't
# raise.
from openpyxl.descriptors.base import Min

_original_min_set = Min.__set__


def _patched_min_set(self, instance, value):
    try:
        _original_min_set(self, instance, value)
    except ValueError:
        pass


Min.__set__ = _patched_min_set


# ---------------------------------------------------------------------------
# Drive mapping-doc lookup
# ---------------------------------------------------------------------------

def _list_children(drive_service, parent_id: str, extra_query: str = "") -> List[Dict]:
    """Return all {id, name} children of parent_id matching extra_query, paginated."""
    results: List[Dict] = []
    page_token = None
    query = f"'{parent_id}' in parents and trashed = false"
    if extra_query:
        query += f" and {extra_query}"

    while True:
        resp = drive_service.files().list(
            q=query,
            fields="nextPageToken, files(id, name)",
            pageToken=page_token,
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
        ).execute()

        results.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break

    return results


def find_month_folder(drive_service, root_folder_id: str,
                       year_month: Optional[str] = None) -> Tuple[str, str]:
    """
    Find the YYYY-MM subfolder under root_folder_id.

    Defaults to the current calendar month. If that exact folder doesn't
    exist yet (e.g. running early before it's been created), falls back to
    the most recent YYYY-MM folder that does exist, printing a warning.

    Returns (folder_id, folder_name).
    """
    target = year_month or datetime.now().strftime("%Y-%m")

    subfolders = _list_children(
        drive_service, root_folder_id,
        extra_query="mimeType = 'application/vnd.google-apps.folder'",
    )
    by_name = {f["name"]: f["id"] for f in subfolders}

    if target in by_name:
        print(f"  Using month folder: '{target}'")
        return by_name[target], target

    month_pattern = re.compile(r"^\d{4}-\d{2}$")
    valid_months = sorted(name for name in by_name if month_pattern.match(name))

    if not valid_months:
        raise FileNotFoundError(
            f"No YYYY-MM folders found under root folder {root_folder_id!r}. "
            f"Folders present: {sorted(by_name)}"
        )

    fallback = valid_months[-1]
    print(f"  WARNING: no '{target}' folder found under root folder — "
          f"falling back to most recent available month folder: '{fallback}'")
    return by_name[fallback], fallback


def _get_all_descendant_folder_ids(drive_service, root_folder_id: str) -> List[str]:
    """Recursively collect root_folder_id and every subfolder ID beneath it,
    at any depth (BFS). Used so the mapping-doc search isn't limited to the
    direct contents of a single folder — real-world Drive layouts sometimes
    nest an extra level (e.g. a per-payment-run subfolder) under the month
    folder."""
    all_ids = [root_folder_id]
    queue = [root_folder_id]

    while queue:
        parent_id = queue.pop()
        subfolders = _list_children(
            drive_service, parent_id,
            extra_query="mimeType = 'application/vnd.google-apps.folder'",
        )
        for folder in subfolders:
            all_ids.append(folder["id"])
            queue.append(folder["id"])

    return all_ids


def find_mapping_sheet_id(drive_service, folder_id: str,
                           include_keyword: str,
                           exclude_keyword: Optional[str] = None) -> Tuple[str, str]:
    """
    Find the Google Sheet whose name contains include_keyword (case-insensitive)
    and, if exclude_keyword is given, does NOT also contain it (used to tell
    two similarly-named mapping docs apart). Searches folder_id AND all of its
    subfolders at any depth, since the mapping docs aren't always directly
    inside the month folder itself.

    Returns (file_id, file_name). Raises if zero or multiple matches are found,
    so a naming-convention drift surfaces immediately instead of silently
    picking the wrong file.
    """
    all_folder_ids = _get_all_descendant_folder_ids(drive_service, folder_id)

    files: List[Dict] = []
    for fid in all_folder_ids:
        files.extend(_list_children(
            drive_service, fid,
            extra_query="mimeType = 'application/vnd.google-apps.spreadsheet'",
        ))

    include_lower = include_keyword.lower()
    exclude_lower = exclude_keyword.lower() if exclude_keyword else None

    matches = []
    for f in files:
        name_lower = f["name"].lower()
        if include_lower not in name_lower:
            continue
        if exclude_lower and exclude_lower in name_lower:
            continue
        matches.append(f)

    if not matches:
        # Diagnostic aid: list EVERYTHING found across the whole subtree (not
        # just sheets), so a naming mismatch is obvious from the error itself
        # instead of requiring another round-trip.
        all_children: List[Dict] = []
        for fid in all_folder_ids:
            all_children.extend(_list_children(drive_service, fid))
        raise FileNotFoundError(
            f"No mapping sheet found in folder {folder_id!r} or its subfolders, "
            f"matching include='{include_keyword}' exclude='{exclude_keyword}'. "
            f"Google Sheets present anywhere in that subtree: {[f['name'] for f in files]}. "
            f"ALL items anywhere in that subtree (any type): "
            f"{[c['name'] for c in all_children] or '(subtree is completely empty)'}"
        )

    if len(matches) > 1:
        raise ValueError(
            f"Multiple mapping sheets matched under folder {folder_id!r} "
            f"(including subfolders): {[f['name'] for f in matches]}. "
            f"Tighten include/exclude keywords."
        )

    return matches[0]["id"], matches[0]["name"]


# ---------------------------------------------------------------------------
# Ampersand-sheet sanitizer
# ---------------------------------------------------------------------------

def sanitize_ampersand_sheets(source_path: Path, cache_dir: Path) -> Tuple[Path, Dict[str, str]]:
    """
    If source_path's workbook has any sheet name containing "&", save a
    renamed copy (same filename) under cache_dir with those sheet names
    changed to use "and" instead of "&", and return (new_path, rename_map)
    where rename_map maps {original_sheet_name: new_sheet_name}.

    If no sheet name contains "&" (or the file can't be opened), returns
    (source_path, {}) unchanged — no copy is made, nothing else changes.
    """
    try:
        wb = openpyxl.load_workbook(source_path, keep_links=False)
    except Exception as e:
        print(f"    WARNING: could not open '{source_path.name}' with openpyxl "
              f"to check for '&' in sheet names ({e}); leaving file as-is.")
        return source_path, {}

    rename_map: Dict[str, str] = {}
    used_names = {ws.title for ws in wb.worksheets}

    for ws in wb.worksheets:
        if "&" in ws.title:
            base_new_name = ws.title.replace("&", "and")
            new_name = base_new_name
            suffix = 1
            while new_name in used_names and new_name != ws.title:
                suffix += 1
                new_name = f"{base_new_name}_{suffix}"
            rename_map[ws.title] = new_name
            used_names.discard(ws.title)
            used_names.add(new_name)
            ws.title = new_name

    if not rename_map:
        return source_path, {}

    cache_dir.mkdir(parents=True, exist_ok=True)
    new_path = cache_dir / source_path.name
    wb.save(new_path)
    print(f"    Sanitized '&' in sheet name(s) for '{source_path.name}': "
          f"{rename_map}")
    return new_path, rename_map