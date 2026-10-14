# dashboard_data.py
"""Reads the batch Excel files and turns them into one clean table for the
Analytics Dashboard. No Streamlit in here, so it can be tested on its own.

The sheet layout it understands (written by excel_writer.py):
    A Submitter | B Date | C Elevation | D Drop | E Floor | F Tower | G Damage | H Dimension | I Unit | J Photo ...
    ... | O Repair Dimension | P Repair Unit
    (older files have no Unit column, so their photo is in column I instead)

    Row 1-2 are headers. Every whiteboard is a block of rows: the details
    (A-F) are merged down the whole block, so only the block's first row holds
    them. Inside a block, a title row names the material (Sealant, Concrete...)
    and the rows under it are that material's defects. A row with an empty
    Damage cell is just a placeholder and is not a defect.
"""
import re
import datetime

import openpyxl
import pandas as pd

from config import REFERENCE_DATA, MATERIAL_RULES
from defect_rules import BOARD_UNITS, REPAIR_UNITS, calculate_repair

SHEET_NAME = "Whiteboard Data"
FIRST_DATA_ROW = 3
UNKNOWN = "Unknown"

COLUMNS = [
    "Folder", "File", "Submitter", "Date", "Elevation", "Drop", "Floor", "Tower",
    "Material", "Defect Code", "Defect Type", "Dimension", "Unit", "Repair Dimension", "Repair Unit",
]

# Title rows hold a material name, matched ignoring capitals
_MATERIAL_BY_LOWER = {name.lower(): name for name in MATERIAL_RULES}

# "TOWER" followed by a separator (or straight by a digit), then the tower name:
# a number (optionally with one letter) or a single letter. 'TOWERS 1' and
# 'TOWER FRONT' deliberately don't match.
_TOWER_PATTERN = re.compile(
    r"(?<![A-Za-z])TOWER(?:[\s_\-]+|(?=\d))(\d+[A-Za-z]?|[A-Za-z])(?![A-Za-z0-9])",
    re.IGNORECASE,
)

# First token of the Damage cell is the base defect code, e.g. "DS CC" -> "DS"
_TYPE_BY_CODE = {row[0]: row[1] for row in REFERENCE_DATA[1:]}

# Sealant codes that carry location modifiers (CC, CF/FC -> CF, GG, FG/GF -> FG, FF)
_SEALANT_CODES = ("DS", "MS")
_CANONICAL_MODIFIERS = {
    "CC": "CC",
    "CF": "CF",
    "FC": "CF",   # FC values merged under CF
    "GG": "GG",
    "FG": "FG",
    "GF": "FG",   # GF values merged under FG
    "FF": "FF",
}
_MODIFIER_ORDER = ("CC", "CF", "GG", "FG", "FF")

# Defect codes that are merged into a single row in the dashboard
_MERGED_DEFECT_CODES = {
    "CC-": "C-",
    "FRAME": "GLASS-FRAME",
}

# Defect codes whose dimensions are suppressed in the dashboard (kept as None placeholder until enabled)
GLASS_FRAME_DIMENSION_PLACEHOLDER = None  # Placeholder for GLASS-FRAME dimension; None hides it in the dashboard
_NO_DIMENSION_DASHBOARD_CODES = {"NP", "BG", "GLASS-FRAME"}

# Valid base defect codes (excluding standalone location modifiers)
_VALID_BASE_CODES = (set(_TYPE_BY_CODE.keys()) - set(_CANONICAL_MODIFIERS.keys())) | set(_MERGED_DEFECT_CODES.keys())


def _base_code(code):
    """Returns the base defect code (e.g. 'DS CF' -> 'DS')."""
    parts = str(code or "").split()
    return parts[0] if parts else ""


def _clean_token(token):
    return str(token or "").upper().strip()


def _canonical_modifier(token):
    """Returns the canonical sealant modifier ('CC', 'CF', 'GG', 'FG', 'FF') if `token` is one, else None."""
    cleaned = _clean_token(token).replace("-", "")
    return _CANONICAL_MODIFIERS.get(cleaned)


def _normalize_defect_code(parts):
    """Normalizes raw damage tokens into the canonical Defect Code:
    - Merges CC- -> C- and FRAME -> GLASS-FRAME
    - Handles slash-written sealant codes like 'DS/MS CC' or 'MS/DS CF' by picking the first sealant code
    - For DS and MS, requires and attaches a canonical location modifier (CC, CF, GG, FG, FF),
      mapping FC -> CF and GF -> FG.
    - Returns None if the row is blank, unrecognized, or a standalone location modifier."""
    if not parts:
        return None

    first = _clean_token(parts[0])
    if "/" in first:
        slash_parts = [p.strip() for p in first.split("/") if p.strip()]
        if slash_parts and slash_parts[0] in _SEALANT_CODES:
            first = slash_parts[0]

    # Ignore standalone location modifiers (e.g. stray "CC" or "FC")
    if _canonical_modifier(first) is not None and len(parts) == 1:
        return None

    # Handle compact strings like "DSCC", "MSCF", "MSFC"
    for s_code in _SEALANT_CODES:
        if first.startswith(s_code) and len(first) > len(s_code):
            mod = _canonical_modifier(first[len(s_code):])
            if mod:
                return f"{s_code} {mod}"

    base = _MERGED_DEFECT_CODES.get(first, first)
    if base not in _VALID_BASE_CODES:
        return None

    if base in _SEALANT_CODES:
        for token in parts[1:]:
            mod = _canonical_modifier(token)
            if mod:
                return f"{base} {mod}"
        # Bare DS or MS without a valid location modifier is incomplete/invalid
        return None

    return base


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

def _text(value):
    """Cell value as clean text ('' for empty). Whole numbers lose the '.0'."""
    if value is None:
        return ""
    if isinstance(value, datetime.datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, datetime.date):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _is_empty(value):
    return value is None or (isinstance(value, str) and not value.strip())


def _number(value):
    """A cell as a number, or None when it holds no number (empty, text, an Excel error)."""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        try:
            return float(value.strip().replace(",", ""))
        except ValueError:
            return None
    return None


_DIMENSION_TEXT = re.compile(r"^\s*(\d+(?:\.\d+)?)\s*([A-Za-z]+(?:\s*(?:\^\s*2|\u00b2|2))?)?")


def _split_dimension(value):
    """A Dimension cell as (number, unit). New files hold a plain number; older ones hold
    text such as '60 CM' or '120 CM2'. (None, '') when there is no number."""
    number = _number(value)
    if number is not None:
        return number, ""
    match = _DIMENSION_TEXT.match(value) if isinstance(value, str) else None
    if not match:
        return None, ""
    unit = re.sub(r"\s*(?:\^\s*2|2)$", "\u00b2", (match.group(2) or "").upper().strip())
    return float(match.group(1)), unit


def tower_from_text(text):
    """Finds a tower in text such as a batch file name, e.g.
    'CERCA NUVEO TOWER 1- FRONT- DROP 7.xlsx' -> '1'. Returns '' when the text
    names no tower, or names more than one (it can't tell which is meant)."""
    found = {m.upper() for m in _TOWER_PATTERN.findall(str(text or ""))}
    return found.pop() if len(found) == 1 else ""


def natural_key(text):
    """Sort key so '2' comes before '10' and 'G' sits with the letters:
    'Floor 2' < 'Floor 10'."""
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r"(\d+)", str(text))]


# ---------------------------------------------------------------------------
# One workbook -> rows
# ---------------------------------------------------------------------------

def parse_workbook(path, folder_label, filename):
    """Reads one batch file and returns a list of defect rows (dicts keyed by COLUMNS).
    Completely blank photo blocks or cleared rows are ignored."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.worksheets[0]
        ws.reset_dimensions()  # some files report a wrong size, which would hide rows

        all_rows = ws.iter_rows(min_row=1, values_only=True)
        header = next(all_rows, None)
        if header is None or len(header) < 8:
            raise ValueError("the sheet doesn't have the expected columns")
        has_unit_col = len(header) > 9 and header[9] == "Whiteboard Photo"
        repair_at = (9 if has_unit_col else 8) + 5  # index of Repair Dimension; Repair Unit is the next one

        file_tower = tower_from_text(filename)

        rows = []
        details = [""] * 6        # Submitter, Date, Elevation, Drop, Floor, Tower of the current block
        material = ""
        pending_bare_sealant = None  # holds a bare DS/MS row in case its modifier was split onto the very next row

        for index, cells in enumerate(all_rows, start=2):
            if index < FIRST_DATA_ROW:
                continue  # second header row
            cells = list(cells) + [None] * (16 - len(cells))  # pad short rows

            # A new block starts where the merged details cells hold a value
            if any(not _is_empty(c) for c in cells[:6]):
                details = [_text(c) for c in cells[:6]]
                details[5] = tower_from_text(details[5]) or details[5] or file_tower
                material = ""
                pending_bare_sealant = None

            damage = cells[6]
            if _is_empty(damage):
                continue  # placeholder row of a material with no defects, or a blank/cleared row
            damage_text = _text(damage)

            # Title row: the material's name, with nothing in the Dimension cell
            title = _MATERIAL_BY_LOWER.get(damage_text.lower())
            if title and _is_empty(cells[7]):
                material = title
                pending_bare_sealant = None
                continue

            # Ignore rows that appear outside any Material section
            if not material:
                continue

            parts = damage_text.split()
            if not parts:
                continue

            # If an older batch split "DS CC" or "MS FC" across two rows (Row 1 = "MS" with dimension,
            # Row 2 = "FC" with no dimension), re-attach the modifier to the pending sealant row.
            if len(parts) == 1 and _canonical_modifier(parts[0]) is not None:
                mod = _canonical_modifier(parts[0])
                if pending_bare_sealant is not None and pending_bare_sealant["Material"] == material:
                    s_base = pending_bare_sealant["Defect Code"]
                    full_code = f"{s_base} {mod}"
                    pending_bare_sealant["Defect Code"] = full_code
                    recalc_repair, recalc_unit = calculate_repair(full_code.split(), pending_bare_sealant["Dimension"])
                    if recalc_repair is not None:
                        pending_bare_sealant["Repair Dimension"] = recalc_repair
                    if recalc_unit:
                        pending_bare_sealant["Repair Unit"] = recalc_unit
                    rows.append(pending_bare_sealant)
                    pending_bare_sealant = None
                continue

            # If this row is a bare "DS" or "MS" without a location modifier, hold it briefly in case
            # the next row is its split modifier; otherwise it is dropped (since DS/MS requires a modifier).
            first_clean = _clean_token(parts[0])
            if len(parts) == 1 and first_clean in _SEALANT_CODES:
                number, unit_in_text = _split_dimension(cells[7])
                unit = (_text(cells[8]) if has_unit_col else "") or unit_in_text
                pending_bare_sealant = {
                    "Folder": folder_label,
                    "File": filename,
                    "Submitter": details[0],
                    "Date": details[1],
                    "Elevation": details[2],
                    "Drop": details[3],
                    "Floor": details[4],
                    "Tower": details[5],
                    "Material": material,
                    "Defect Code": first_clean,
                    "Defect Type": _TYPE_BY_CODE.get(first_clean, ""),
                    "Dimension": number,
                    "Unit": unit,
                    "Repair Dimension": None,
                    "Repair Unit": _text(cells[repair_at + 1]),
                }
                continue

            pending_bare_sealant = None
            code = _normalize_defect_code(parts)
            if not code:
                continue
            base = _base_code(code)

            # Ensure the defect code actually belongs to the material section it sits under
            allowed_for_material = set(MATERIAL_RULES.get(material, [])) | {"GLASS-FRAME", "C-"}
            if base not in allowed_for_material:
                continue

            # Glass-Frame / Frame has no dimensions (uses placeholder None)
            if code in _NO_DIMENSION_DASHBOARD_CODES:
                number, unit, repair, repair_unit = GLASS_FRAME_DIMENSION_PLACEHOLDER, "", None, ""
            else:
                number, unit_in_text = _split_dimension(cells[7])
                unit = (_text(cells[8]) if has_unit_col else "") or unit_in_text

                repair = _number(cells[repair_at])
                repair_unit = _text(cells[repair_at + 1])
                if repair is None:
                    repair, calculated_unit = calculate_repair(code.split(), number)
                    repair_unit = repair_unit or calculated_unit

            rows.append({
                "Folder": folder_label,
                "File": filename,
                "Submitter": details[0],
                "Date": details[1],
                "Elevation": details[2],
                "Drop": details[3],
                "Floor": details[4],
                "Tower": details[5],
                "Material": material,
                "Defect Code": code,
                "Defect Type": _TYPE_BY_CODE.get(base, ""),
                "Dimension": number,
                "Unit": unit,
                "Repair Dimension": repair,
                "Repair Unit": repair_unit,
            })
        return rows
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# Rows -> table, filters and counts
# ---------------------------------------------------------------------------

def rows_to_dataframe(rows):
    """One table of every defect. Blank Tower/Floor/Material/Type become 'Unknown'
    so those defects still show up in the charts instead of vanishing."""
    df = pd.DataFrame(rows, columns=COLUMNS)
    for column in ("Floor", "Tower", "Material", "Defect Type"):
        df[column] = df[column].replace("", UNKNOWN).fillna(UNKNOWN)
    df["Dimension"] = pd.to_numeric(df["Dimension"], errors="coerce")
    df["Repair Dimension"] = pd.to_numeric(df["Repair Dimension"], errors="coerce")
    return df


def apply_filters(df, towers=None, floors=None, materials=None):
    """Keeps only the chosen values. An empty/None list for a filter means 'all'."""
    if towers:
        df = df[df["Tower"].isin(towers)]
    if floors:
        df = df[df["Floor"].isin(floors)]
    if materials:
        df = df[df["Material"].isin(materials)]
    return df


def count_by(df, column):
    """Defect frequency per value of `column`, as a table with that column and a
    'Defects' count. Floors/towers are in natural order (2 before 10); the rest
    are most-frequent first."""
    if df.empty:
        return pd.DataFrame({column: [], "Defects": []})
    counts = df.groupby(column).size().reset_index(name="Defects")
    if column in ("Floor", "Tower"):
        order = sorted(counts[column], key=natural_key)
        counts = counts.set_index(column).loc[order].reset_index()
    else:
        counts = counts.sort_values(["Defects", column], ascending=[False, True])
    return counts.reset_index(drop=True)


# Build the canonical display order for defect codes:
#   1. All Linear Meter (L.M. -> LM) defects first (C-, CC+, C+, DG, DS CC..FF, MS CC..FF)
#   2. All Square Meter (SQ.M. -> SQ LM) defects second (BH, US, DP, FP, BP)
#   3. Non-dimensional defects last (NP, BG, GLASS-FRAME)
def _build_code_order():
    linear_codes = []
    area_codes = []
    other_codes = []

    for row in REFERENCE_DATA[1:]:
        code = row[0]
        if code in ("CC-", "FRAME") or code in _CANONICAL_MODIFIERS:
            continue

        rep_unit = (REPAIR_UNITS.get(code, "") or "").upper().strip()
        if rep_unit in ("L.M.", "L.M", "LM"):
            bucket = linear_codes
        elif rep_unit in ("SQ.M.", "SQ.M", "SQM", "SQ M", "SQ LM", "SQ.L.M."):
            bucket = area_codes
        else:
            bucket = other_codes

        if code in _SEALANT_CODES:
            for mod in _MODIFIER_ORDER:
                bucket.append(f"{code} {mod}")
        else:
            bucket.append(code)

    order = linear_codes + area_codes + other_codes
    return {code: idx for idx, code in enumerate(order)}


_CODE_ORDER = _build_code_order()

DEFECT_SUMMARY_COLUMNS = [
    "Defect Code", "Defect Type", "Defects", "Board Unit", "Total Dimension",
    "Repair Unit", "Total Repair Dimension", "Missing Dimension",
]


def _most_common_text(values):
    """The most frequent non-empty text in `values` ('' if there is none)."""
    counts = {}
    for value in values:
        if isinstance(value, str) and value:
            counts[value] = counts.get(value, 0) + 1
    return max(counts, key=counts.get) if counts else ""


def _convert_board_dimension(total_value, raw_unit):
    """Converts board dimensions from CM -> LM (/ 100) and CM² / SQ CM -> SQ LM (/ 10,000)."""
    unit_clean = (raw_unit or "").upper().strip()
    if total_value is None:
        if unit_clean == "CM":
            return None, "LM"
        if unit_clean in ("CM\u00b2", "CM2", "CM^2", "SQ CM", "SQ.CM", "SQ. CM"):
            return None, "SQ LM"
        return None, raw_unit

    if unit_clean == "CM":
        return float(total_value) / 100.0, "LM"
    if unit_clean in ("CM\u00b2", "CM2", "CM^2", "SQ CM", "SQ.CM", "SQ. CM"):
        return float(total_value) / 10000.0, "SQ LM"
    return float(total_value), raw_unit


def _normalize_repair_unit(raw_unit):
    """Standardizes repair units (L.M. -> LM, SQ.M. -> SQ LM) so both columns match."""
    unit_clean = (raw_unit or "").upper().strip()
    if unit_clean in ("L.M.", "L.M", "LM"):
        return "LM"
    if unit_clean in ("SQ.M.", "SQ.M", "SQM", "SQ M", "SQ LM", "SQ.L.M."):
        return "SQ LM"
    return raw_unit


def _unit_group_rank(repair_unit):
    """Secondary sort rank so LM rows always precede SQ LM rows, which precede non-dimensional rows."""
    if repair_unit == "LM":
        return 0
    if repair_unit == "SQ LM":
        return 1
    return 2


def dimension_by_defect(df):
    """One row per defect code, ordered with all Linear Meter (LM) defects first,
    then Square Meter (SQ LM) defects, then non-dimensional defects."""
    if df.empty:
        return pd.DataFrame(columns=DEFECT_SUMMARY_COLUMNS)

    work_df = df.copy()
    work_df["Defect Code"] = work_df["Defect Code"].apply(
        lambda c: _normalize_defect_code(str(c or "").split())
    )
    work_df = work_df[work_df["Defect Code"].notna() & (work_df["Defect Code"] != "")]

    grouped = {code: group for code, group in work_df.groupby("Defect Code", sort=False)}

    has_sealant = (work_df["Material"] == "Sealant").any() or any(
        _base_code(c) in _SEALANT_CODES for c in grouped
    )
    all_codes = set(grouped.keys())
    if has_sealant:
        for s_code in _SEALANT_CODES:
            for mod in _MODIFIER_ORDER:
                all_codes.add(f"{s_code} {mod}")

    rows = []
    for code in all_codes:
        base = _base_code(code)
        group = grouped.get(code)

        if code in _NO_DIMENSION_DASHBOARD_CODES:
            rows.append({
                "Defect Code": code,
                "Defect Type": _TYPE_BY_CODE.get(base, "") or UNKNOWN,
                "Defects": len(group) if group is not None else 0,
                "Board Unit": "",
                "Total Dimension": GLASS_FRAME_DIMENSION_PLACEHOLDER,
                "Repair Unit": "",
                "Total Repair Dimension": GLASS_FRAME_DIMENSION_PLACEHOLDER,
                "Missing Dimension": 0,
            })
            continue

        if group is None or group.empty:
            _, display_board_unit = _convert_board_dimension(0.0, BOARD_UNITS.get(base, "CM"))
            display_repair_unit = _normalize_repair_unit(REPAIR_UNITS.get(base, "L.M."))
            rows.append({
                "Defect Code": code,
                "Defect Type": _TYPE_BY_CODE.get(base, "") or UNKNOWN,
                "Defects": 0,
                "Board Unit": display_board_unit,
                "Total Dimension": 0.0,
                "Repair Unit": display_repair_unit,
                "Total Repair Dimension": 0.0,
                "Missing Dimension": 0,
            })
            continue

        dimensions = group["Dimension"].dropna()
        repairs = group["Repair Dimension"].dropna()

        raw_board_unit = _most_common_text(group["Unit"]) or BOARD_UNITS.get(base, "")
        raw_board_total = float(dimensions.sum()) if len(dimensions) else None
        converted_board_total, display_board_unit = _convert_board_dimension(raw_board_total, raw_board_unit)

        raw_repair_unit = _most_common_text(group["Repair Unit"]) or REPAIR_UNITS.get(base, "")
        display_repair_unit = _normalize_repair_unit(raw_repair_unit)

        rows.append({
            "Defect Code": code,
            "Defect Type": _TYPE_BY_CODE.get(base, "") or UNKNOWN,
            "Defects": len(group),
            "Board Unit": display_board_unit,
            "Total Dimension": converted_board_total,
            "Repair Unit": display_repair_unit,
            "Total Repair Dimension": float(repairs.sum()) if len(repairs) else None,
            "Missing Dimension": int(group["Dimension"].isna().sum()) if BOARD_UNITS.get(base) else 0,
        })
    summary = pd.DataFrame(rows, columns=DEFECT_SUMMARY_COLUMNS)
    summary["_unit_rank"] = summary["Repair Unit"].map(_unit_group_rank)
    summary["_order"] = summary["Defect Code"].map(lambda code: _CODE_ORDER.get(code, len(_CODE_ORDER)))
    return (
        summary.sort_values(["_unit_rank", "_order", "Defect Code"])
        .drop(columns=["_unit_rank", "_order"])
        .reset_index(drop=True)
    )