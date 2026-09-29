#!/usr/bin/env python3
"""
CALC earnings + near-miss report.

For each request JSON in the input folder:
  1. POST it to /calculations and GET /programs for the same participant/supplier
  2. Join potential_earnings back to every request transaction by transaction_id
  3. Check every transaction against every program it did NOT earn on, and flag:
       - missed incentive date      exact SKU + eligible type, date within N days outside the period
       - different transaction type exact SKU + within dates, but the wrong transaction type
       - product family             first sku_key segment matches + eligible type + within dates
       - eligible, did not earn     exact SKU + type + dates all match, but nothing earned
                                    (usually a qualification, gate or threshold; toggle in config)
     A transaction earns at most once per program, so if it earned on any offer
     in a program, no other offer in that program is checked for it. Otherwise it
     gets at most one line per program on each sheet: missed date shows only the
     closest offer; the other sheets list the program's remaining offers.
     For distributor suppliers (GET /enrollments), transactions whose types need a
     seller (GET /calculations/requirements) must have that distributor as seller;
     a missing or different seller goes straight to No match with the reason.
     Transaction types that never earn (booking, retailer_orders_from_distribution
     by default) are not checked. Gate definitions and zero-rate stacked sku_sets
     never pay, so they never produce flags either.
  4. Write one Excel workbook per retailer (participant_key): a Summary sheet by
     supplier and program, then one sheet each for Earned, Missed date, Product
     family, Transaction type, Eligible did not earn, and No match, each grouped
     by program supplier

Input filename convention: {participant_key}__{program_supplier_key}.json
  e.g. ag_vend_austin_tx_000__bayer.json

Both API calls send program_supplier_key and time_frame as params and the
participant key as the `participant-key` header.

Usage:
  python earnings_report.py                    # call the API for every input file
  python earnings_report.py --reuse-responses  # rebuild CSVs from saved responses, no API calls
  python earnings_report.py --config other.json
"""
import argparse
import json
import os
import sys
from datetime import date, timedelta
from pathlib import Path

import requests

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl is required for the Excel output. Install it with: pip install openpyxl")

TXN_FIELDS = ["transaction_id", "transaction_type", "sku_key", "uom_key",
              "quantity", "amount", "invoice_date"]
REQUIRED_CONFIG = ("base_url", "time_frame", "input_dir", "output_dir")

FLAG_DATE = "missed incentive date"
FLAG_TYPE = "different transaction type"
FLAG_FAMILY = "product family"
FLAG_ELIGIBLE = "eligible, did not earn"
DEFAULT_NON_EARNING_TYPES = ["booking", "retailer_orders_from_distribution"]


# ---------------------------------------------------------------- config / API

def load_config(path):
    with open(path) as f:
        cfg = json.load(f)
    missing = [k for k in REQUIRED_CONFIG if k not in cfg]
    if missing:
        sys.exit(f"Config is missing: {', '.join(missing)}")
    return cfg


def parse_filename(path):
    """ag_vend_austin_tx_000__bayer.json -> ('ag_vend_austin_tx_000', 'bayer')"""
    parts = path.stem.split("__", 1)
    if len(parts) != 2 or not all(parts):
        return None
    return parts[0], parts[1]


def build_headers(cfg, participant):
    """participant key goes in a header; supplier and time_frame go in params."""
    headers = {"Content-Type": "application/json",
               cfg.get("participant_header", "participant-key"): participant}
    token = cfg.get("api_key") or os.environ.get(cfg.get("api_key_env", "CALC_API_KEY"))
    if token:
        headers[cfg.get("auth_header", "Authorization")] = cfg.get("auth_prefix", "Bearer ") + token
    return headers


def call_api(cfg, method, route, participant, supplier, payload=None):
    """supplier=None leaves program_supplier_key out of the params (e.g. /enrollments)."""
    url = cfg["base_url"].rstrip("/") + route
    params = {"time_frame": str(cfg["time_frame"])}
    if supplier:
        params["program_supplier_key"] = supplier
    r = requests.request(method, url, params=params, json=payload,
                         headers=build_headers(cfg, participant),
                         timeout=cfg.get("timeout_seconds", 120))
    if not r.ok:
        raise RuntimeError(f"{method} {route} HTTP {r.status_code}: {r.text[:500]}")
    body = r.json()
    if body.get("errors"):
        print(f"WARN {method} {route}: response errors: {json.dumps(body['errors'])[:500]}")
    return body


_fetched_this_run = set()


def get_or_load(cfg, reuse, cache_path, method, route, participant, supplier, payload=None):
    """Call the API and save the response, or load the saved copy with --reuse-responses.

    A response already fetched earlier in this run (e.g. /enrollments, which is per
    participant) is read from its saved file instead of being fetched again.
    """
    if cache_path in _fetched_this_run:
        return json.loads(cache_path.read_text())
    if reuse:
        if not cache_path.exists():
            raise RuntimeError(f"no saved response at {cache_path}")
        return json.loads(cache_path.read_text())
    body = call_api(cfg, method, route, participant, supplier, payload)
    cache_path.write_text(json.dumps(body, indent=2))
    _fetched_this_run.add(cache_path)
    return body


# ---------------------------------------------------------------- earnings

def flatten_transactions(payload):
    rows = []
    for entity in payload.get("entities", []):
        for t in entity.get("transactions", []):
            row = {"entity_id": entity.get("id", "")}
            row.update({f: t.get(f, "") for f in TXN_FIELDS})
            row["seller_key"] = (t.get("seller") or {}).get("key", "")
            rows.append(row)
    return rows


def earnings_by_transaction(response):
    """{transaction_id: {offer_key: {program_key, offer_key, incentive_amount}}}

    If the same transaction/offer pair appears more than once (e.g. several
    incentive rates), the incentive amounts are summed into one offer entry.
    """
    out = {}
    for block in response.get("data", []):
        for e in block.get("potential_earnings", []):
            tid = str(e.get("transaction_id", ""))
            offer_key = e.get("offer_key", "")
            rec = out.setdefault(tid, {}).setdefault(offer_key, {
                "program_key": e.get("program_key", ""),
                "offer_key": offer_key,
                "incentive_amount": 0.0,
            })
            rec["incentive_amount"] += float(e.get("incentive_amount") or 0)
    return out


def failed_qualifications(calc_response, programs_response):
    """{(program_key, offer_key): "Missed: <label> (actual x vs target y); ..."} from qualification_results.

    Waived qualifications count as met. Labels come from the program and offer
    qualifications in GET /programs, falling back to the qualification key.
    """
    labels = {}
    for program in programs_response.get("data", []):
        pk = program.get("key", "")
        for q in program.get("qualifications") or []:
            labels[(pk, q.get("key"))] = q.get("label") or q.get("key")
        for offer in program.get("offers") or []:
            for q in offer.get("qualifications") or []:
                labels[(pk, q.get("key"))] = q.get("label") or q.get("key")

    results = []
    for block in calc_response.get("data", []):
        results.extend(block.get("qualification_results") or [])
    results.extend(calc_response.get("qualification_results") or [])

    notes = {}
    for r in results:
        if r.get("met") or r.get("waived"):
            continue
        pk, ok, qk = r.get("program_key", ""), r.get("offer_key", ""), r.get("qualification_key", "")
        label = labels.get((pk, qk), qk)
        try:
            nums = f" (actual {float(r.get('actual')):,.2f} vs target {float(r.get('target')):,.2f})"
        except (TypeError, ValueError):
            nums = ""
        text = f"Missed: {label}{nums}"
        existing = notes.setdefault((pk, ok), [])
        if text not in existing:
            existing.append(text)
    return {k: "; ".join(v) for k, v in notes.items()}


# ---------------------------------------------------------------- program rules

def parse_date(value):
    try:
        return date.fromisoformat(str(value)[:10])
    except (TypeError, ValueError):
        return None


def first_segment(sku_key):
    return str(sku_key).split("_", 1)[0]


def offer_rules(offer):
    """Normalize one offer into rules: which SKUs, which transaction type, which period.

    role is "earning" for components that pay (standard/gated incentive rates,
    paying stacked sku_sets, stepped filters), or "gate" / "stack requirement" for
    components that must be present but never pay (gate_definition, zero-rate
    stacked sku_sets). Non-earning rules never produce near-miss flags; a
    transaction that fully matches one is simply expected and left alone.
    """
    strategy = offer.get("strategy_key")
    raw = []  # (period_key, transaction_type, sku_keys, role)

    if strategy in ("standard", "gated"):
        for inc in offer.get("incentives") or []:
            by_period = {}
            for r in inc.get("rates") or []:
                pk = r.get("period_key") or inc.get("period_key")   # gated puts it on the rate
                by_period.setdefault(pk, []).append(r.get("sku_key"))
            for pk, skus in by_period.items():
                raw.append((pk, inc.get("transaction_type"), skus, "earning"))
            gate = inc.get("gate_definition")
            if gate:
                gate_skus = [r.get("sku_key") for r in inc.get("rates") or []]
                raw.append((gate.get("period_key"), gate.get("transaction_type"), gate_skus, "gate"))

    elif strategy == "stacked":
        for s in (offer.get("incentive") or {}).get("sku_sets") or []:
            rates = s.get("rates") or []
            pays = any((r.get("rate") or 0) > 0 for r in rates)
            raw.append((s.get("period_key"), s.get("transaction_type"),
                        [r.get("sku_key") for r in rates],
                        "earning" if pays else "stack requirement"))

    elif strategy == "stepped":
        for f in (offer.get("incentive") or {}).get("filters") or []:
            raw.append((f.get("period_key"), f.get("transaction_type"), f.get("sku_keys") or [], "earning"))

    periods = {p.get("key"): p for p in offer.get("time_periods") or []}
    rules = []
    for pk, ttype, skus, role in raw:
        period = periods.get(pk)
        skus = {s for s in skus if s}
        start = parse_date(period.get("start_date")) if period else None
        end = parse_date(period.get("end_date")) if period else None
        if not (ttype and skus and start and end):
            continue
        rules.append({"period_key": pk, "transaction_type": ttype, "sku_keys": skus,
                      "families": {first_segment(s) for s in skus},
                      "start": start, "end": end, "role": role})
    return rules


def load_offers(programs_response):
    """[{program_key, offer_key, strategy, qualifier, rules}], plus warnings for unreadable offers."""
    offers, warnings = [], []
    for program in programs_response.get("data", []):
        for offer in program.get("offers") or []:
            rules = offer_rules(offer)
            if not rules:
                warnings.append(f"{offer.get('key')} (strategy '{offer.get('strategy_key')}'): "
                                f"couldn't read SKUs/types/dates, skipped for near-miss checks")
                continue
            offers.append({
                "program_key": program.get("key", ""),
                "offer_key": offer.get("key", ""),
                "strategy": offer.get("strategy_key", ""),
                "qualifier": offer.get("qualifier") or program.get("qualifier") or "",
                "rules": rules,
            })
    return offers, warnings


# ---------------------------------------------------------------- near-miss checks

def check_rule(txn, rule, window_days):
    """Return None, {"flag": "match"}, or a review dict with the flag and its details."""
    d = parse_date(txn["invoice_date"])
    sku, ttype = txn["sku_key"], txn["transaction_type"]
    exact = sku in rule["sku_keys"]
    family = not exact and first_segment(sku) in rule["families"]
    type_ok = ttype == rule["transaction_type"]
    start, end = rule["start"], rule["end"]
    in_dates = d is not None and start <= d <= end
    base = {"period_key": rule["period_key"], "period_start": start, "period_end": end}

    if exact and type_ok and in_dates:
        return {"flag": "match"}
    if exact and type_ok and d is not None:
        window = timedelta(days=window_days)
        if start - window <= d < start:
            return {**base, "flag": FLAG_DATE, "days_off": (start - d).days, "direction": "before start"}
        if end < d <= end + window:
            return {**base, "flag": FLAG_DATE, "days_off": (d - end).days, "direction": "after end"}
        return None
    if exact and not type_ok and in_dates:
        return {**base, "flag": FLAG_TYPE, "required_type": rule["transaction_type"]}
    if family and type_ok and in_dates:
        similar = [s for s in rule["sku_keys"] if first_segment(s) == first_segment(sku)]
        # closest = longest shared start of the key, then alphabetical
        similar.sort(key=lambda s: (-len(os.path.commonprefix([s, sku])), s))
        return {**base, "flag": FLAG_FAMILY, "closest_sku": similar[0], "other_skus": similar[1:]}
    return None


def review_transaction(txn, earned, offers, cfg):
    """All review dicts for one transaction, one per offer + flag type.

    A transaction earns at most once per program, so once it has earned on any
    offer in a program, the other offers in that program are not checked.
    """
    earned_programs = {o["program_key"] for o in earned}
    # Bookings and orders never earn on their own (they only satisfy a gate or a
    # zero-rate stack requirement), so near-miss flags don't apply to them.
    non_earning = set(cfg.get("non_earning_transaction_types", DEFAULT_NON_EARNING_TYPES))
    if txn["transaction_type"] in non_earning:
        return []
    window = cfg.get("near_miss_days", 30)
    flag_eligible = cfg.get("flag_eligible_not_earned", True)
    reviews = []
    for offer in offers:
        if offer["program_key"] in earned_programs:
            continue
        ident = {"program_key": offer["program_key"], "offer_key": offer["offer_key"]}
        results = [check_rule(txn, rule, window) for rule in offer["rules"]]
        matched = [rule for rule, res in zip(offer["rules"], results) if res and res["flag"] == "match"]
        if matched:
            # Fully matches this offer: expected for gate/stack pieces; otherwise worth a look.
            earning = [r for r in matched if r["role"] == "earning"]
            if flag_eligible and earning:
                reviews.append({**ident, "flag": FLAG_ELIGIBLE, "qualifier": offer["qualifier"],
                                "period_key": earning[0]["period_key"],
                                "period_start": earning[0]["start"], "period_end": earning[0]["end"]})
            continue
        # Only earning components produce near-miss flags. Gate definitions and
        # zero-rate stack requirements never pay, so they only suppress flags above.
        seen = set()
        for rule, res in zip(offer["rules"], results):
            if rule["role"] != "earning":
                continue
            if res and res["flag"] not in seen:
                seen.add(res["flag"])
                reviews.append({**ident, **res})
    return collapse_per_program(reviews)


def collapse_per_program(reviews):
    """At most one line per program and flag type.

    Missed date keeps only the offer the transaction came closest to (staggered
    offers: Jan-Feb, Mar-Apr, May-Jun). Other flags keep the first offer by key and
    list the program's remaining offers with the same miss in other_offers.
    """
    grouped = {}
    for rv in reviews:
        grouped.setdefault((rv["program_key"], rv["flag"]), []).append(rv)
    out = []
    for (_, flag), items in grouped.items():
        if flag == FLAG_DATE:
            best = min(items, key=lambda rv: (rv["days_off"], rv["offer_key"]))
            out.append({**best, "other_offers": []})
        else:
            items.sort(key=lambda rv: rv["offer_key"])
            out.append({**items[0], "other_offers": [rv["offer_key"] for rv in items[1:]]})
    return out


# ---------------------------------------------------------------- seller rule

def supplier_type(enrollments, supplier):
    """program_supplier_type for this supplier from GET /enrollments, or None if not listed."""
    for e in (enrollments or {}).get("data", []):
        if e.get("program_supplier_key") == supplier:
            return e.get("program_supplier_type")
    return None


def seller_transaction_types(requirements):
    """Transaction types that require seller.key, from GET /calculations/requirements."""
    types = set()
    for block in (requirements or {}).get("data", []):
        for f in ((block.get("transaction_requirements") or {}).get("fields")) or []:
            if f.get("field") == "seller":
                types.update(f.get("transaction_types") or [])
    return types


def seller_problem(txn, supplier, seller_types):
    """For distributor programs: the seller must be the distributor itself."""
    if txn["transaction_type"] not in seller_types:
        return None
    seller = txn.get("seller_key") or ""
    if not seller:
        return ("Missing seller", f"{txn['transaction_type']} needs seller {supplier}")
    if seller != supplier:
        return ("Different seller", f"seller is {seller}; {supplier} programs need seller {supplier}")
    return None


# ---------------------------------------------------------------- no-match reasons

def no_match_reason(txn, offers, cfg):
    """Why a transaction didn't earn and wasn't a near miss: (reason, detail)."""
    ttype, sku = txn["transaction_type"], txn["sku_key"]
    if ttype in set(cfg.get("non_earning_transaction_types", DEFAULT_NON_EARNING_TYPES)):
        return ("Transaction type never earns", f"{ttype} only counts toward gates or stacks")
    rules = [r for o in offers for r in o["rules"] if r["role"] == "earning"]
    if not any(sku in r["sku_keys"] for r in rules):
        if any(first_segment(sku) in r["families"] for r in rules):
            return ("Product family only, type or date also off",
                    f"'{first_segment(sku)}' products are in a program, but not this SKU")
        return ("SKU not in any program", "")
    if any(check_rule(txn, r, 0) == {"flag": "match"} for r in rules):
        return ("Matched an offer but did not earn", "eligible flag is turned off in config")
    if any(sku in r["sku_keys"] and ttype == r["transaction_type"] for r in rules):
        days = cfg.get("near_miss_days", 30)
        return (f"Date more than {days} days outside incentive period", "")
    return ("Transaction type and date both off", "")


# ---------------------------------------------------------------- records

def build_records(txns, earnings, offers, cfg, supplier, seller_types, qual_notes=None):
    """[(transaction, earned offers, review dicts)] for every request transaction.

    Transactions with no earnings and no reviews get txn["no_match_reason"] and
    txn["no_match_detail"] for the No match sheet.
    """
    records = []
    for t in txns:
        earned = list(earnings.get(str(t["transaction_id"]), {}).values())
        problem = None if earned else seller_problem(t, supplier, seller_types)
        reviews = [] if problem else review_transaction(t, earned, offers, cfg)
        for rv in reviews:
            rv["qualification_note"] = (qual_notes or {}).get((rv["program_key"], rv["offer_key"]), "")
        if not earned and not reviews:
            t["no_match_reason"], t["no_match_detail"] = problem or no_match_reason(t, offers, cfg)
        records.append((t, earned, reviews))
    return records


# ---------------------------------------------------------------- Excel workbook

FONT_NAME = "Arial"
F_BASE = Font(name=FONT_NAME, size=10)
F_BOLD = Font(name=FONT_NAME, size=10, bold=True)
F_SHEET_TITLE = Font(name=FONT_NAME, size=14, bold=True)
F_SECTION = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
F_GROUP = Font(name=FONT_NAME, size=10, bold=True, color="305496")
FILL_SECTION = PatternFill("solid", fgColor="305496")
FILL_HEADER = PatternFill("solid", fgColor="D9E1F2")
FILL_TOTAL = PatternFill("solid", fgColor="F2F2F2")
FILL_PROGRAM = PatternFill("solid", fgColor="E2EFDA")
FMT_MONEY = '$#,##0.00;($#,##0.00);"-"'
FMT_COUNT = '#,##0;(#,##0);"-"'
FMT_QTY = '#,##0.##'
FMT_DATE = "yyyy-mm-dd"

# (header, transaction field, number format)
TXN_COLS = [("Transaction ID", "transaction_id", None),
            ("Transaction type", "transaction_type", None),
            ("SKU", "sku_key", None),
            ("UOM", "uom_key", None),
            ("Quantity", "quantity", FMT_QTY),
            ("Amount", "amount", FMT_MONEY),
            ("Invoice date", "invoice_date", FMT_DATE),
            ("Seller", "seller_key", None)]


def _num(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return value


def _txn_values(txn):
    vals = []
    for _, field, fmt in TXN_COLS:
        v = txn.get(field, "")
        if fmt == FMT_DATE:
            v = parse_date(v) or v
        elif fmt in (FMT_QTY, FMT_MONEY):
            v = _num(v)
        vals.append(v)
    return vals


def _put(ws, row, col, value, font=F_BASE, fmt=None, fill=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = font
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    return cell


def _section_title(ws, row, text, width):
    for c in range(1, width + 1):
        _put(ws, row, c, None, fill=FILL_SECTION)
    _put(ws, row, 1, text, font=F_SECTION, fill=FILL_SECTION)
    return row + 1


def _header(ws, row, headers):
    for c, h in enumerate(headers, start=1):
        _put(ws, row, c, h, font=F_BOLD, fill=FILL_HEADER)
    return row + 1


def _rows(ws, row, values_list, formats):
    for values in values_list:
        for c, (v, fmt) in enumerate(zip(values, formats), start=1):
            _put(ws, row, c, v, fmt=fmt)
        row += 1
    return row


def _autosize(ws, max_width=60):
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None or cell.font == F_SECTION or cell.font == F_SHEET_TITLE:
                continue
            v = cell.value
            n = 10 if hasattr(v, "year") else len(str(v))
            widths[cell.column] = max(widths.get(cell.column, 0), n)
    for col, n in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = min(max(n + 2, 10), max_width)


def _sheet_name(name, used):
    clean = "".join(ch for ch in name if ch not in '[]:*?/\\')[:31] or "Sheet"
    base, i = clean, 2
    while clean.lower() in used:
        suffix = f" ({i})"
        clean = base[:31 - len(suffix)] + suffix
        i += 1
    used.add(clean.lower())
    return clean


def _earned_lines(records):
    lines = [(o["program_key"], o["offer_key"], txn, o["incentive_amount"])
             for txn, earned, _ in records for o in earned]
    lines.sort(key=lambda x: (x[0], x[1], str(x[2]["transaction_id"])))
    return lines


def _write_earned_sheet(ws, participant, suppliers):
    """Earned transactions, grouped by supplier, then program and offer, with totals."""
    txn_headers = [h for h, _, _ in TXN_COLS]
    txn_formats = [f for _, _, f in TXN_COLS]
    headers = ["Program", "Offer", *txn_headers, "Earned amount"]
    earn_col = len(headers)
    L = get_column_letter(earn_col)
    _put(ws, 1, 1, f"Earned | {participant}", font=F_SHEET_TITLE)
    row = _header(ws, 3, headers)
    ws.freeze_panes = ws.cell(row=row, column=1)
    sheet_first = row

    for supplier in sorted(suppliers):
        lines = _earned_lines(suppliers[supplier])
        row = _section_title(ws, row, f"{supplier} ({len(lines)} transaction-offer lines)", len(headers))
        supplier_first = group_first = row
        if not lines:
            _put(ws, row, 1, "None")
            row += 2
            continue
        for i, (program, offer, txn, amount) in enumerate(lines):
            row = _rows(ws, row, [[program, offer, *_txn_values(txn), round(amount, 2)]],
                        [None, None, *txn_formats, FMT_MONEY])
            nxt = lines[i + 1] if i + 1 < len(lines) else None
            if nxt is None or (nxt[0], nxt[1]) != (program, offer):
                for c in range(1, earn_col + 1):
                    _put(ws, row, c, None, fill=FILL_TOTAL)
                _put(ws, row, 2, "Offer total", font=F_BOLD, fill=FILL_TOTAL)
                _put(ws, row, earn_col, f"=SUBTOTAL(9,{L}{group_first}:{L}{row - 1})",
                     font=F_BOLD, fmt=FMT_MONEY, fill=FILL_TOTAL)
                row += 1
                group_first = row
        _put(ws, row, 1, f"{supplier} total", font=F_BOLD)
        _put(ws, row, earn_col, f"=SUBTOTAL(9,{L}{supplier_first}:{L}{row - 1})",
             font=F_BOLD, fmt=FMT_MONEY)
        row += 2

    _put(ws, row, 1, "Grand total", font=F_BOLD)
    _put(ws, row, earn_col, f"=SUBTOTAL(9,{L}{sheet_first}:{L}{row - 1})", font=F_BOLD, fmt=FMT_MONEY)
    _autosize(ws)


def _write_review_sheet(ws, title, participant, suppliers, flag, extra_headers, extra_formats, extra_values):
    """One near-miss reason, grouped by supplier, one line per transaction + offer."""
    txn_headers = [h for h, _, _ in TXN_COLS]
    txn_formats = [f for _, _, f in TXN_COLS]
    headers = ["Program", "Offer", *txn_headers, *extra_headers]
    _put(ws, 1, 1, f"{title} | {participant}", font=F_SHEET_TITLE)
    row = _header(ws, 3, headers)
    ws.freeze_panes = ws.cell(row=row, column=1)
    for supplier in sorted(suppliers):
        items = [(rv, txn) for txn, _, reviews in suppliers[supplier] for rv in reviews if rv["flag"] == flag]
        items.sort(key=lambda x: (x[0]["program_key"], x[0]["offer_key"], str(x[1]["transaction_id"])))
        row = _section_title(ws, row, f"{supplier} ({len(items)})", len(headers))
        if not items:
            _put(ws, row, 1, "None")
            row += 1
        else:
            row = _rows(ws, row,
                        [[rv["program_key"], rv["offer_key"], *_txn_values(txn), *extra_values(rv)]
                         for rv, txn in items],
                        [None, None, *txn_formats, *extra_formats])
        row += 1
    _autosize(ws)


def _write_no_match_sheet(ws, participant, suppliers):
    """Transactions that didn't earn and weren't a near miss on anything."""
    txn_headers = [h for h, _, _ in TXN_COLS]
    txn_formats = [f for _, _, f in TXN_COLS]
    headers = [*txn_headers, "Reason", "Detail"]
    _put(ws, 1, 1, f"No match | {participant}", font=F_SHEET_TITLE)
    row = _header(ws, 3, headers)
    ws.freeze_panes = ws.cell(row=row, column=1)
    for supplier in sorted(suppliers):
        txns = [t for t, earned, reviews in suppliers[supplier] if not earned and not reviews]
        txns.sort(key=lambda t: (t.get("no_match_reason", ""), str(t["transaction_id"])))
        row = _section_title(ws, row, f"{supplier} ({len(txns)})", len(headers))
        if not txns:
            _put(ws, row, 1, "None")
            row += 1
        else:
            row = _rows(ws, row,
                        [[*_txn_values(t), t.get("no_match_reason", ""), t.get("no_match_detail", "")]
                         for t in txns],
                        [*txn_formats, None, None])
        row += 1
    _autosize(ws)


REVIEW_SHEETS = [
    # (sheet name, flag, extra headers, extra formats, extra values)
    ("Missed date", FLAG_DATE,
     ["Incentive start", "Incentive end", "Missed by (days)", "Before / after", "Qualification"],
     [FMT_DATE, FMT_DATE, FMT_COUNT, None, None],
     lambda rv: [rv["period_start"], rv["period_end"], rv["days_off"], rv["direction"],
                 rv.get("qualification_note", "")]),
    ("Product family", FLAG_FAMILY,
     ["Closest offer SKU", "Other matching offer SKUs", "Qualification", "Other offers"],
     [None, None, None, None],
     lambda rv: [rv["closest_sku"], ", ".join(rv["other_skus"]), rv.get("qualification_note", ""),
                 ", ".join(rv.get("other_offers", []))]),
    ("Transaction type", FLAG_TYPE,
     ["Earning transaction type", "Incentive start", "Incentive end", "Qualification", "Other offers"],
     [None, FMT_DATE, FMT_DATE, None, None],
     lambda rv: [rv["required_type"], rv["period_start"], rv["period_end"],
                 rv.get("qualification_note", ""), ", ".join(rv.get("other_offers", []))]),
    ("Eligible, did not earn", FLAG_ELIGIBLE,
     ["Qualification", "Incentive start", "Incentive end"],
     [None, FMT_DATE, FMT_DATE],
     lambda rv: [rv.get("qualification_note") or "No failed qualification found; check gates/thresholds",
                 rv["period_start"], rv["period_end"]]),
]


def _new_bucket():
    return {"earning": set(), "earnings": 0.0, "date": set(), "family": set(), "type": set(), "eligible": set()}


def _summary_rows(records):
    """{program_key: {"all": bucket, "offers": {offer_key: bucket}}}, counting distinct transactions."""
    progs = {}
    flag_field = {FLAG_DATE: "date", FLAG_FAMILY: "family", FLAG_TYPE: "type", FLAG_ELIGIBLE: "eligible"}

    def buckets(program, offer):
        p = progs.setdefault(program, {"all": _new_bucket(), "offers": {}})
        return p["all"], p["offers"].setdefault(offer, _new_bucket())

    for txn, earned, reviews in records:
        tid = str(txn["transaction_id"])
        for o in earned:
            for b in buckets(o["program_key"], o["offer_key"]):
                b["earning"].add(tid)
                b["earnings"] += o["incentive_amount"]
        for rv in reviews:
            for b in buckets(rv["program_key"], rv["offer_key"]):
                b[flag_field[rv["flag"]]].add(tid)
    return progs


def _bucket_values(b):
    return [len(b["earning"]), round(b["earnings"], 2), len(b["date"]),
            len(b["family"]), len(b["type"]), len(b["eligible"])]


def write_workbook(path, participant, suppliers, cfg):
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"

    ALL = "All offers"
    headers = ["Program supplier", "Program", "Offer", "Earning transactions", "Total earnings",
               "Missed date only", "Product family only", "Transaction type only",
               "Eligible, did not earn"]
    formats = [None, None, None, FMT_COUNT, FMT_MONEY, FMT_COUNT, FMT_COUNT, FMT_COUNT, FMT_COUNT]
    first_num = 4  # first numeric column
    _put(ws, 1, 1, f"Summary | {participant}", font=F_SHEET_TITLE)
    _put(ws, 2, 1, f"time_frame {cfg['time_frame']}. Counts are distinct transactions. "
                   f"\"{ALL}\" rows count each transaction once per program, so they can be lower "
                   f"than the sum of the offer rows below them. Totals add up the \"{ALL}\" rows only.")
    row = _header(ws, 4, headers)
    table_first = row
    offer_col = get_column_letter(3)

    def total_row(r, label, first, last, font=F_BOLD, fill=None):
        for c in range(1, len(headers) + 1):
            _put(ws, r, c, None, fill=fill)
        _put(ws, r, 2, label, font=font, fill=fill)
        for c in range(first_num, len(headers) + 1):
            col = get_column_letter(c)
            _put(ws, r, c, f'=SUMIFS({col}{first}:{col}{last},{offer_col}{first}:{offer_col}{last},"{ALL}")',
                 font=font, fmt=formats[c - 1], fill=fill)

    for supplier in sorted(suppliers):
        progs = _summary_rows(suppliers[supplier])
        first = row
        _put(ws, row, 1, supplier, font=F_GROUP)
        row += 1
        if not progs:
            _put(ws, row, 2, "No earnings or near misses")
            row += 1
        for program in sorted(progs):
            p = progs[program]
            # program row (bold), then its offers indented beneath, collapsible
            for c, (v, fmt) in enumerate(zip([None, program, ALL, *_bucket_values(p["all"])], formats), start=1):
                _put(ws, row, c, v, font=F_BOLD, fmt=fmt, fill=FILL_PROGRAM if c > 1 else None)
            row += 1
            offers_first = row
            for offer in sorted(p["offers"]):
                row = _rows(ws, row, [[None, None, offer, *_bucket_values(p["offers"][offer])]], formats)
            if row > offers_first:
                ws.row_dimensions.group(offers_first, row - 1, outline_level=1, hidden=False)
        total_row(row, f"{supplier} total", first, row - 1, fill=FILL_TOTAL)
        row += 2

    total_row(row, "Grand total", table_first, row - 1)
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.freeze_panes = ws.cell(row=table_first, column=1)
    _autosize(ws)

    _write_earned_sheet(wb.create_sheet("Earned"), participant, suppliers)
    for name, flag, headers_x, formats_x, values_x in REVIEW_SHEETS:
        _write_review_sheet(wb.create_sheet(name), name, participant, suppliers,
                            flag, headers_x, formats_x, values_x)
    _write_no_match_sheet(wb.create_sheet("No match"), participant, suppliers)

    wb.save(path)


# ---------------------------------------------------------------- main

def process_file(cfg, path, out_dir, reuse):
    """Call the APIs for one input file and return (participant, supplier, records)."""
    keys = parse_filename(path)
    if not keys:
        print(f"SKIP {path.name}: expected {{participant_key}}__{{program_supplier_key}}.json")
        return None
    participant, supplier = keys

    payload = json.loads(path.read_text())
    if payload.get("participant_key") and payload["participant_key"] != participant:
        print(f"WARN {path.name}: participant_key in file is '{payload['participant_key']}'")

    try:
        calc = get_or_load(cfg, reuse, out_dir / "responses" / path.name,
                           "POST", "/calculations", participant, supplier, payload)
        programs = get_or_load(cfg, reuse, out_dir / "programs" / path.name,
                               "GET", "/programs", participant, supplier)
    except Exception as e:
        print(f"FAIL {path.name}: {e}")
        return None

    # Seller rule: only for distributor suppliers, only on the transaction types
    # /calculations/requirements says need a seller. Failures here don't stop the run.
    seller_types = set()
    try:
        enrollments = get_or_load(cfg, reuse, out_dir / "enrollments" / f"{participant}.json",
                                  "GET", "/enrollments", participant, None)
        stype = supplier_type(enrollments, supplier)
        if stype is None:
            print(f"WARN {path.name}: '{supplier}' isn't in /enrollments for {participant}; seller check skipped")
        elif stype == "distributor":
            requirements = get_or_load(cfg, reuse, out_dir / "requirements" / path.name,
                                       "GET", "/calculations/requirements", participant, supplier)
            seller_types = seller_transaction_types(requirements)
    except Exception as e:
        print(f"WARN {path.name}: seller check skipped ({e})")

    txns = flatten_transactions(payload)
    earnings = earnings_by_transaction(calc)
    offers, warnings = load_offers(programs)
    for w in warnings:
        print(f"WARN {path.name}: {w}")

    orphans = set(earnings) - {str(t["transaction_id"]) for t in txns}
    if orphans:
        print(f"WARN {path.name}: {len(orphans)} earning transaction_id(s) not in request, "
              f"e.g. {sorted(orphans)[:5]}")

    records = build_records(txns, earnings, offers, cfg, supplier, seller_types,
                            failed_qualifications(calc, programs))
    earned = sum(1 for _, e, _ in records if e)
    flagged = sum(1 for _, _, r in records if r)
    print(f"OK   {path.name}: {len(records)} transactions, {earned} earned, {flagged} flagged for review, "
          f"{len(offers)} offers checked")
    return participant, supplier, records


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--config", default="config.json")
    ap.add_argument("--reuse-responses", action="store_true",
                    help="rebuild the workbook from saved API responses instead of calling the API")
    args = ap.parse_args()

    # Resolve paths relative to the script (for config.json) and to the config
    # file (for input/output), so it works no matter which folder you run it from.
    config_path = Path(args.config)
    if not config_path.is_absolute():
        config_path = Path(__file__).resolve().parent / config_path
    cfg = load_config(config_path)
    base = config_path.parent
    in_dir = base / cfg["input_dir"]
    out_dir = base / cfg["output_dir"]
    for sub in ("responses", "programs", "enrollments", "requirements"):
        (out_dir / sub).mkdir(parents=True, exist_ok=True)

    files = sorted(in_dir.glob("*.json"))
    if not files:
        sys.exit(f"No .json files in {in_dir}")
    by_participant = {}
    for path in files:
        result = process_file(cfg, path, out_dir, args.reuse_responses)
        if result:
            participant, supplier, records = result
            by_participant.setdefault(participant, {})[supplier] = records

    for participant, suppliers in sorted(by_participant.items()):
        wb_path = out_dir / f"{participant}.xlsx"
        write_workbook(wb_path, participant, suppliers, cfg)
        print(f"XLSX {participant}: {len(suppliers)} supplier(s), 7 sheets -> {wb_path}")


if __name__ == "__main__":
    main()