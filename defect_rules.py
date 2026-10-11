# defect_rules.py
"""Reads the per-defect rules (units and repair formulas) from config.py.
Nothing is defined here: to change a unit or a formula, edit config.py."""
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
    code = parts[0]
    unit, formula, formula_with_location = _REPAIR_RULES.get(code, ("", "", ""))

    modifier = parts[1] if len(parts) > 1 else None
    if modifier and formula_with_location:
        return formula_with_location.format(h=h_cell, modifier=modifier), unit
    return (formula.format(h=h_cell) if formula else ""), unit
