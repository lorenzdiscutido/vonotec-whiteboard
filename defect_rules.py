# defect_rules.py
"""Reads the per-defect rules (units and repair formulas) from config.py.
Nothing is defined here: to change a unit or a formula, edit config.py."""
import re
from decimal import Decimal, ROUND_HALF_UP, ROUND_UP

from config import REFERENCE_DATA, MATERIAL_RULES

_COLUMN = {name: i for i, name in enumerate(REFERENCE_DATA[0])}


def _value(row, column_name):
    """A column of a REFERENCE_DATA row by its header name ("" if missing)."""
    i = _COLUMN.get(column_name)
    return row[i] if i is not None and i < len(row) else ""


# Defect code -> unit of the dimension written on the board ("" = none)
BOARD_UNITS = {row[0]: _value(row, "Board Unit") for row in REFERENCE_DATA[1:]}

# Defect code -> (repair unit, formula template, formula template with location code)
_REPAIR_RULES = {
    row[0]: (
        _value(row, "Repair Unit"),
        _value(row, "Repair Formula"),
        _value(row, "Repair Formula (with location)"),
    )
    for row in REFERENCE_DATA[1:]
}

# Standardize sealant location modifiers so FC uses CF and GF uses FG
_MODIFIER_ALIASES = {
    "CC": "CC",
    "CF": "CF",
    "FC": "CF",
    "GG": "GG",
    "FG": "FG",
    "GF": "FG",
    "FF": "FF",
}


def _build_defect_units():
    """(material, [codes], unit) groups, used by the Gemini prompt and the unit fixer.
    Built from the codes each material allows (MATERIAL_RULES) and their Board Unit."""
    groups = []
    for material, allowed_codes in MATERIAL_RULES.items():
        by_unit = {}
        for code in allowed_codes:
            if code in BOARD_UNITS:  # codes without a reference row are skipped
                by_unit.setdefault(BOARD_UNITS[code], []).append(code)
        groups += [(material, codes, unit) for unit, codes in by_unit.items()]
    return groups


DEFECT_UNITS = _build_defect_units()


def repair_rule(parts, h_cell):
    """Returns (excel_formula_without_equals, repair_unit) for one damage entry.
    `parts` is the damage text split into words ('DS CC' -> ['DS', 'CC']) and
    `h_cell` is the cell holding the board dimension. The formula is "" when the
    defect has no calculated repair dimension."""
    code = parts[0].upper() if parts else ""
    unit, formula, formula_with_location = _REPAIR_RULES.get(code, ("", "", ""))

    raw_mod = parts[1].upper().replace("-", "") if len(parts) > 1 else None
    modifier = _MODIFIER_ALIASES.get(raw_mod) if raw_mod else None
    if modifier and formula_with_location:
        return formula_with_location.format(h=h_cell, modifier=modifier), unit
    return (formula.format(h=h_cell) if formula else ""), unit


# Defect code -> unit of the calculated repair dimension ("" = none)
REPAIR_UNITS = {code: rule[0] for code, rule in _REPAIR_RULES.items()}

# Code -> Repair Limit (column F of REFERENCE_DATA), for the VLOOKUPs inside the formulas
_REPAIR_LIMITS = {
    row[0]: row[5] for row in REFERENCE_DATA[1:]
    if len(row) > 5 and isinstance(row[5], (int, float)) and not isinstance(row[5], bool)
}


# ---------------------------------------------------------------------------
# The repair formulas, worked out in Python
# ---------------------------------------------------------------------------

_LIMIT_LOOKUP = re.compile(r"VLOOKUP\(\"([^\"]+)\",\s*'Reference Data'!\$A:\$F,\s*6,\s*FALSE\)")


def _decimal(value):
    """A float as an exact decimal, ignoring floating-point noise (Excel works to 15 digits)."""
    return Decimal(repr(round(float(value), 10)))


def _roundup(value, digits=0):
    """Excel ROUNDUP: rounds away from zero."""
    return float(_decimal(value).quantize(Decimal(1).scaleb(-int(digits)), rounding=ROUND_UP))


def _mround(value, multiple):
    """Excel MROUND: the nearest multiple (halves go away from zero)."""
    step = _decimal(multiple)
    return float((_decimal(value) / step).quantize(Decimal(1), rounding=ROUND_HALF_UP) * step)


_EXCEL_FUNCTIONS = {
    "ROUNDUP": _roundup,
    "MROUND": _mround,
    "IF": lambda condition, if_true, if_false: if_true if condition else if_false,
}


def calculate_repair(parts, number):
    """Python twin of the Excel repair formula: returns (repair_dimension, repair_unit).
    `parts` is the damage text split into words ('DS CC' -> ['DS', 'CC']) and `number` the
    board dimension as a number."""
    code = parts[0].upper() if parts else ""
    unit = REPAIR_UNITS.get(code, "")
    if not parts or number is None:
        return None, unit

    formula, unit = repair_rule(parts, repr(float(number)))
    if not formula:
        return None, unit
    try:
        expression = _LIMIT_LOOKUP.sub(lambda m: repr(float(_REPAIR_LIMITS[m.group(1)])), formula)
        value = eval(expression, {"__builtins__": {}}, _EXCEL_FUNCTIONS)  # config.py text + numbers only
        return round(float(value), 2), unit
    except (KeyError, ArithmeticError, SyntaxError, NameError, TypeError, ValueError):
        return None, unit