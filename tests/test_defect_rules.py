"""Tests for defect_rules module."""
import unittest
from defect_rules import repair_rule, calculate_repair, BOARD_UNITS, REPAIR_UNITS


class TestDefectRules(unittest.TestCase):
    def test_units_loaded(self):
        self.assertIsInstance(BOARD_UNITS, dict)
        self.assertIsInstance(REPAIR_UNITS, dict)
        self.assertIn("DS", BOARD_UNITS)

    def test_repair_rule_empty(self):
        formula, unit = repair_rule([], "H2")
        self.assertEqual(formula, "")
        self.assertEqual(unit, "")

    def test_repair_rule_standard_code(self):
        # Test for DS (Delamination / Sealant) or other standard codes
        if "DS" in BOARD_UNITS:
            formula, unit = repair_rule(["DS"], "H2")
            self.assertIsInstance(formula, str)
            self.assertIsInstance(unit, str)

    def test_repair_rule_with_location_alias(self):
        # 'FC' should be mapped to 'CF' alias
        formula_fc, _ = repair_rule(["DS", "FC"], "H2")
        formula_cf, _ = repair_rule(["DS", "CF"], "H2")
        self.assertEqual(formula_fc, formula_cf)

    def test_calculate_repair_none(self):
        val, unit = calculate_repair([], None)
        self.assertIsNone(val)
        val, unit = calculate_repair(["DS"], None)
        self.assertIsNone(val)

    def test_calculate_repair_with_values(self):
        # If repair formulas exist for a code, verify it produces a valid float/number or None cleanly
        val, unit = calculate_repair(["DS"], 10.0)
        if val is not None:
            self.assertIsInstance(val, (int, float))


if __name__ == "__main__":
    unittest.main()
