# dashboard_data.py
"""Reads the batch Excel files and turns them into one clean table for the
Analytics Dashboard. No Streamlit in here, so it can be tested on its own.

The sheet layout it understands (written by excel_writer.py):
    A Submitter | B Date | C Elevation | D Drop | E Floor | F Tower | G Damage | H Dimension | I Unit | J Photo ...
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

SHEET_NAME = "Whiteboard Data"
FIRST_DATA_ROW = 3
UNKNOWN = "Unknown"

COLUMNS = [
    "Folder", "File", "Submitter", "Date", "Elevation", "Drop", "Floor", "Tower",
    "Material", "Defect Code", "Defect Type", "Dimension", "Unit",
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

# First token of the Damage cell is the defect code, e.g. "DS CC" -> "DS"
_TYPE_BY_CODE = {row[0]: row[1] for row in REFERENCE_DATA[1:]}


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


def tower_from_text(text):
    """Finds a tower in text such as a batch file name, e.g.
    'CERCA NUVEO TOWER 1- FRONT- DROP 7.xlsx' -> '1'. Returns '' when the text
    names no tower, or names more than one (it can't tell which is meant).
    Works in any capitalization, and with a space, dash or underscore after the word
    ('Tower 2', 'TOWER-2', 'tower_2', 'TOWER2'); a tower can be a number or a
    letter, optionally with a letter after the number ('3', 'A', '2B')."""
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
    Raises an exception if the file can't be opened or has no readable sheet."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.worksheets[0]
        ws.reset_dimensions()  # some files report a wrong size, which would hide rows

        all_rows = ws.iter_rows(min_row=1, values_only=True)
        header = next(all_rows, None)
        if header is None or len(header) < 8:
            raise ValueError("the sheet doesn't have the expected columns")
        has_unit_col = len(header) > 9 and header[9] == "Whiteboard Photo"

        # Workers sometimes forget to write the tower on the board, but the batch
        # name usually has it ("... TOWER 1 - FRONT - DROP 7"), so it fills the gap.
        file_tower = tower_from_text(filename)

        rows = []
        details = [""] * 6        # Submitter, Date, Elevation, Drop, Floor, Tower of the current block
        material = ""
        for index, cells in enumerate(all_rows, start=2):
            if index < FIRST_DATA_ROW:
                continue  # second header row
            cells = list(cells) + [None] * (9 - len(cells))  # pad short rows

            # A new block starts where the merged details cells hold a value
            if any(not _is_empty(c) for c in cells[:6]):
                details = [_text(c) for c in cells[:6]]
                # What is written on the board wins ("TOWER 1" is shortened to "1" so it
                # matches the others); only a blank tower is filled in from the file name.
                details[5] = tower_from_text(details[5]) or details[5] or file_tower
                material = ""

            damage = cells[6]
            if _is_empty(damage):
                continue  # placeholder row of a material with no defects, or a blank row
            damage_text = _text(damage)

            # Title row: the material's name, with nothing in the Dimension cell
            title = _MATERIAL_BY_LOWER.get(damage_text.lower())
            if title and _is_empty(cells[7]):
                material = title
                continue

            code = damage_text.split()[0] if damage_text.split() else ""
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
                "Defect Type": _TYPE_BY_CODE.get(code, ""),
                "Dimension": cells[7] if not _is_empty(cells[7]) else None,
                "Unit": _text(cells[8]) if has_unit_col else "",
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
