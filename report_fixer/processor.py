"""Clean up the Printing Daily Report workbook.

The original workbook is parsed into a simple model (sheets -> blocks of
title rows and tables), every correction rule is applied to that model, and
``writer.py`` then renders a brand-new, consistently formatted workbook.

Rules applied (numbers match the requirement document):

1. Rows whose item name has no value in any other column are removed.
2. Day Consumption is converted from tons to kilograms and
   "No of days Consumption" is recalculated from the kg value.
3. "All RAW Aging Stock": unneeded columns removed.
4. "All TagEnv Aging Stock": unneeded columns removed and repeated item
   names grouped (one summary row with average Qty / Holding Days, the
   individual lots grouped underneath and collapsed).
5/7. Zero values are replaced by "-" on every sheet.
6. Decimal values are rounded to the nearest whole number (half-up).
8. Formatting is handled by ``writer.py``.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal
from typing import Any

import openpyxl

DASH = "-"

# Columns to remove (rules 3 and 4), compared after normalising the header.
RAW_AGING_DROP = {"lot code", "lot no", "supplier lot no", "category", "gsm", "qty", "machine line"}
TAGENV_AGING_DROP = {"lot code", "lot no", "supplier lot no", "width", "gsm", "source type"}

# Obvious spelling mistakes in the source headers.
HEADER_FIXES = {
    "consumbtion": "Consumption",
    "watage": "Wastage",
    "ouput": "Output",
}


# --------------------------------------------------------------------------
# Model
# --------------------------------------------------------------------------
@dataclass
class Row:
    values: list[Any]
    kind: str = "data"  # data | total | group | detail
    group_size: int = 0  # for kind == "group": number of detail rows below


@dataclass
class Table:
    header: list[str]
    rows: list[Row] = field(default_factory=list)
    title: str | None = None  # section title shown above the header
    numeric_cols: set[int] = field(default_factory=set)


@dataclass
class Sheet:
    name: str
    kind: str  # daily | raw_aging | tagenv_aging | ink | generic
    title: str | None = None  # report title on top of the sheet
    tables: list[Table] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)  # human readable change log


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def norm(text: Any) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip().lower()


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() in ("", "NULL", "null", "#N/A"))


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def is_formula(value: Any) -> bool:
    return isinstance(value, str) and value.startswith("=")


def round_half_up(value: float) -> int:
    return int(Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def fix_header(text: Any) -> str:
    out = re.sub(r"\s+", " ", str(text or "")).strip()
    for wrong, right in HEADER_FIXES.items():
        out = re.sub(wrong, right, out, flags=re.IGNORECASE)
    if norm(out) == "row labels":
        out = "Item Name"
    return out


def find_col(header: list[str], *needles: str, exact: bool = False) -> int | None:
    for i, h in enumerate(header):
        n = norm(h)
        for needle in needles:
            if (n == needle) if exact else (needle in n):
                return i
    return None


def is_total_label(value: Any) -> bool:
    n = norm(value)
    return n.startswith("grand total") or n == "total"


def sheet_kind(name: str) -> str:
    n = norm(name)
    if "tagenv" in n or ("tag" in n and "env" in n and "aging" in n):
        return "tagenv_aging"
    if "raw" in n and "aging" in n:
        return "raw_aging"
    if "ink" in n:
        return "ink"
    if "daily" in n:
        return "daily"
    return "generic"


# --------------------------------------------------------------------------
# Parsing
# --------------------------------------------------------------------------
def _looks_like_header(cells: list[Any]) -> bool:
    filled = [c for c in cells if not is_blank(c)]
    if len(filled) < 2:
        return False
    texts = [norm(c) for c in filled if isinstance(c, str)]
    if len(texts) != len(filled):
        return False
    return any("item name" in t or t == "row labels" for t in texts)


def parse_sheet(ws) -> Sheet:
    sheet = Sheet(name=ws.title, kind=sheet_kind(ws.title))
    rows = [list(r) for r in ws.iter_rows(values_only=True)]

    pending_title: str | None = None
    current: Table | None = None
    width = 0

    for cells in rows:
        filled = [c for c in cells if not is_blank(c)]
        if current is None:
            if not filled:
                continue
            if _looks_like_header(cells):
                width = max(i for i, c in enumerate(cells) if not is_blank(c)) + 1
                current = Table(header=[fix_header(c) for c in cells[:width]], title=pending_title)
                pending_title = None
                sheet.tables.append(current)
            elif len(filled) == 1 and isinstance(filled[0], str):
                # A stand-alone text line: report title or a section title.
                if sheet.title is None and not sheet.tables and pending_title is None and sheet.kind == "daily":
                    sheet.title = filled[0].strip()
                else:
                    pending_title = filled[0].strip()
            continue

        # Inside a table: a blank row ends it.
        body = (cells + [None] * width)[:width]  # cells beyond the header are stray (e.g. scratch formulas)
        if all(is_blank(c) for c in body):
            current = None
            continue
        if _looks_like_header(cells):
            width = max(i for i, c in enumerate(cells) if not is_blank(c)) + 1
            current = Table(header=[fix_header(c) for c in cells[:width]], title=pending_title)
            pending_title = None
            sheet.tables.append(current)
            continue
        body = [None if (is_blank(c) or is_formula(c)) else c for c in body]
        kind = "total" if is_total_label(body[0]) else "data"
        if kind == "total":
            body[0] = "Grand Total"
        current.rows.append(Row(values=body, kind=kind))

    for table in sheet.tables:
        _detect_numeric(table)
    return sheet


def _detect_numeric(table: Table) -> None:
    table.numeric_cols = set()
    for c in range(len(table.header)):
        vals = [r.values[c] for r in table.rows if not is_blank(r.values[c])]
        if vals and sum(is_number(v) for v in vals) >= max(1, len(vals) * 0.8):
            table.numeric_cols.add(c)


# --------------------------------------------------------------------------
# Rules
# --------------------------------------------------------------------------
def drop_columns(table: Table, drop: set[str]) -> list[str]:
    keep = [i for i, h in enumerate(table.header) if norm(h) not in drop]
    removed = [h for i, h in enumerate(table.header) if i not in keep]
    table.header = [table.header[i] for i in keep]
    for r in table.rows:
        r.values = [r.values[i] for i in keep]
    _detect_numeric(table)
    return removed


def remove_empty_rows(table: Table, zero_is_empty: bool) -> list[str]:
    """Rule 1: drop rows that only carry an item name."""
    key = find_col(table.header, "item name")
    if key is None:
        key = 0

    def empty(v: Any) -> bool:
        return is_blank(v) or (zero_is_empty and is_number(v) and v == 0)

    removed, kept = [], []
    for r in table.rows:
        others = [v for i, v in enumerate(r.values) if i != key]
        if r.kind == "data" and not is_blank(r.values[key]) and all(empty(v) for v in others):
            removed.append(str(r.values[key]))
        elif r.kind == "data" and all(is_blank(v) for v in r.values):
            continue
        else:
            kept.append(r)
    table.rows = kept
    return removed


def fix_day_consumption(table: Table) -> bool:
    """Rule 2: Day Consumption tons -> kg and recompute No. of days."""
    dc = find_col(table.header, "day consumption", exact=True)
    days = find_col(table.header, "no of days consumption", "no. of days consumption", "days consumption")
    if dc is None:
        return False
    if days == dc:
        days = None
    closing = find_col(table.header, "closing")

    for r in table.rows:
        if r.kind != "data":
            continue
        v = r.values[dc]
        r.values[dc] = v * 1000 if is_number(v) else None
        if days is not None and closing is not None:
            stock, per_day = r.values[closing], r.values[dc]
            r.values[days] = stock / per_day if is_number(stock) and is_number(per_day) and per_day else None

    data = [r for r in table.rows if r.kind == "data"]
    for r in table.rows:
        if r.kind != "total":
            continue
        total_dc = sum(x.values[dc] for x in data if is_number(x.values[dc]))
        r.values[dc] = total_dc
        if days is not None and closing is not None:
            stock = r.values[closing]
            r.values[days] = stock / total_dc if is_number(stock) and total_dc else None

    table.header[dc] = "Day Consumption in Kgs"
    if days is not None:
        table.header[days] = "No of days Consumption"
    _detect_numeric(table)
    return True


def fix_ratio_percentages(table: Table) -> list[str]:
    """Percent columns stored as fractions (0.15) would round to 0; show them as 15."""
    changed = []
    for c, h in enumerate(table.header):
        n = norm(h)
        if "%" not in n and "percent" not in n:
            continue
        vals = [r.values[c] for r in table.rows if is_number(r.values[c])]
        if vals and any(v != 0 for v in vals) and all(abs(v) <= 1 for v in vals):
            for r in table.rows:
                if is_number(r.values[c]):
                    r.values[c] = r.values[c] * 100
            changed.append(h)
    return changed


def fill_missing_totals(table: Table) -> list[str]:
    """Grand Total cells left blank in the source are filled with the column sum."""
    filled = []
    data = [r for r in table.rows if r.kind == "data"]
    for r in table.rows:
        if r.kind != "total":
            continue
        for c in table.numeric_cols:
            n = norm(table.header[c])
            if c == 0 or not is_blank(r.values[c]) or "%" in n or "percent" in n or "days" in n:
                continue
            nums = [x.values[c] for x in data if is_number(x.values[c])]
            if nums:
                r.values[c] = sum(nums)
                filled.append(table.header[c])
    return filled


def round_and_dash(table: Table) -> tuple[int, int]:
    """Rules 5, 6 and 7: round decimals, then show zeros (and empty numbers) as '-'."""
    rounded = dashed = 0
    for r in table.rows:
        for c, v in enumerate(r.values):
            if c == 0:
                continue
            if is_number(v):
                if isinstance(v, float) and v != int(v):
                    rounded += 1
                v = round_half_up(v)
                if v == 0:
                    v = DASH
                    dashed += 1
                r.values[c] = v
            elif v is None and c in table.numeric_cols and r.kind != "detail":
                r.values[c] = DASH
    return rounded, dashed


def group_repeated_items(table: Table) -> int:
    """Rule 4: one summary row per item name, its lots grouped below."""
    key = find_col(table.header, "item name") or 0
    qty = find_col(table.header, "qty", "qtn", "quantity")
    hold = find_col(table.header, "holding days")

    order: list[Any] = []
    buckets: dict[Any, list[Row]] = {}
    for r in table.rows:
        if r.kind != "data":
            continue
        name = r.values[key]
        if name not in buckets:
            buckets[name] = []
            order.append(name)
        buckets[name].append(r)

    header = list(table.header)
    header.insert(key + 1, "No. of Lots")
    if qty is not None:
        header[qty + (1 if qty > key else 0)] = "Qty (Avg)"
    if hold is not None:
        header[hold + (1 if hold > key else 0)] = "Holding Days (Avg)"

    def avg(rows: list[Row], col: int | None) -> Any:
        if col is None:
            return None
        nums = [x.values[col] for x in rows if is_number(x.values[col])]
        return sum(nums) / len(nums) if nums else None

    def single(rows: list[Row], col: int) -> Any:
        distinct = {x.values[col] for x in rows if not is_blank(x.values[col])}
        if len(distinct) == 1:
            return distinct.pop()
        return "Multiple" if distinct else None

    new_rows: list[Row] = []
    grouped = 0
    for name in order:
        lots = buckets[name]
        if len(lots) == 1:
            vals = list(lots[0].values)
            vals.insert(key + 1, 1)
            new_rows.append(Row(values=vals, kind="data"))
            continue
        grouped += 1
        summary = []
        for c in range(len(table.header)):
            if c == key:
                summary.append(name)
            elif c in (qty, hold):
                summary.append(avg(lots, c))
            elif c in table.numeric_cols:
                summary.append(avg(lots, c))
            else:
                summary.append(single(lots, c))
        summary.insert(key + 1, len(lots))
        new_rows.append(Row(values=summary, kind="group", group_size=len(lots)))
        for lot in lots:
            vals = list(lot.values)
            vals.insert(key + 1, None)
            new_rows.append(Row(values=vals, kind="detail"))

    table.header = header
    table.rows = new_rows + [r for r in table.rows if r.kind == "total"]
    _detect_numeric(table)
    lots_col = key + 1
    table.numeric_cols.add(lots_col)
    return grouped


# --------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------
def process_sheet(sheet: Sheet, zero_is_empty: bool = True) -> Sheet:
    for table in sheet.tables:
        label = f"{table.title} – " if table.title else ""

        if sheet.kind == "raw_aging":
            removed = drop_columns(table, RAW_AGING_DROP)
            if removed:
                sheet.notes.append(f"Removed columns: {', '.join(removed)}")
        elif sheet.kind == "tagenv_aging":
            removed = drop_columns(table, TAGENV_AGING_DROP)
            if removed:
                sheet.notes.append(f"Removed columns: {', '.join(removed)}")

        if fix_day_consumption(table):
            sheet.notes.append(f"{label}Day Consumption converted from tons to Kgs; No of days Consumption recalculated")

        for h in fix_ratio_percentages(table):
            sheet.notes.append(f"{label}'{h}' converted from fraction to percentage")

        gone = remove_empty_rows(table, zero_is_empty)
        if gone:
            sheet.notes.append(f"{label}Removed {len(gone)} empty item row(s): {', '.join(gone)}")

        totals = fill_missing_totals(table)
        if totals:
            sheet.notes.append(f"{label}Filled missing Grand Total for: {', '.join(totals)}")

        if sheet.kind == "tagenv_aging":
            n = group_repeated_items(table)
            sheet.notes.append(f"Grouped {n} repeated item name(s) – click the + in the margin to expand the lots")

        rounded, dashed = round_and_dash(table)
        if rounded:
            sheet.notes.append(f"{label}Rounded {rounded} decimal value(s)")
        if dashed:
            sheet.notes.append(f"{label}Replaced {dashed} zero value(s) with '-'")
    return sheet


def process_workbook(data: bytes, zero_is_empty: bool = True) -> list[Sheet]:
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=False)
    return [process_sheet(parse_sheet(ws), zero_is_empty) for ws in wb.worksheets]
