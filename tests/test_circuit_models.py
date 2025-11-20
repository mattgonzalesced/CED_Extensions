"""Unit tests for circuit sizing model objects."""
import os
import sys
import unittest

CURRENT_DIR = os.path.dirname(__file__)
LIB_PATH = os.path.abspath(os.path.join(CURRENT_DIR, "..", "AE pyTools.extension", "lib"))
if LIB_PATH not in sys.path:
    sys.path.insert(0, LIB_PATH)

from CEDElectrical.circuit_sizing.models.circuit_branch import CircuitBranchModel, CircuitSettings


class CircuitModelTests(unittest.TestCase):
    """Validates lightweight model behaviors for circuit sizing."""

    def test_settings_serialization_round_trip(self):
        settings = CircuitSettings(
            min_wire_size="10",
            max_wire_size="500",
            min_breaker_size=30,
            auto_calculate_breaker=True,
        )
        data = settings.to_dict()
        restored = CircuitSettings.from_dict(data)

        self.assertEqual(restored.min_wire_size, "10")
        self.assertEqual(restored.max_wire_size, "500")
        self.assertEqual(restored.min_breaker_size, 30)
        self.assertTrue(restored.auto_calculate_breaker)

    def test_settings_defaults_respected(self):
        partial = {"min_wire_size": "8"}
        restored = CircuitSettings.from_dict(partial)
        self.assertEqual(restored.min_wire_size, "8")
        self.assertEqual(restored.max_wire_size, CircuitSettings().max_wire_size)
        self.assertEqual(restored.min_breaker_size, CircuitSettings().min_breaker_size)

    def test_voltage_drop_rule_changes_with_feeder_flag(self):
        model = CircuitBranchModel(
            circuit_id=1,
            panel="P1",
            circuit_number="1",
            name="Test",
            branch_type="BRANCH",
        )
        self.assertEqual(model.max_voltage_drop, model.settings.max_branch_voltage_drop)

        model.is_feeder = True
        self.assertEqual(model.max_voltage_drop, model.settings.max_feeder_voltage_drop)

    def test_classification_prioritizes_transformer_flags(self):
        model = CircuitBranchModel(
            circuit_id=2,
            panel="XFMR",
            circuit_number="PRI",
            name="Transformer Primary",
            branch_type="FEEDER",
        )
        model.is_transformer_primary = True
        self.assertEqual(model.classification, "TRANSFORMER_PRIMARY")

        model.is_transformer_primary = False
        model.is_transformer_secondary = True
        self.assertEqual(model.classification, "TRANSFORMER_SECONDARY")

        model.is_transformer_secondary = False
        model.is_feeder = True
        self.assertEqual(model.classification, "FEEDER")

        model.is_feeder = False
        model.is_spare = True
        self.assertEqual(model.classification, "SPARE")

        model.is_spare = False
        self.assertEqual(model.classification, "BRANCH")


if __name__ == "__main__":
    unittest.main()
