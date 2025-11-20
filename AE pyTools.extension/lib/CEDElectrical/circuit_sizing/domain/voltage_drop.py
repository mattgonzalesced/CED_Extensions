from __future__ import annotations

from typing import Optional

from CEDElectrical.refdata.impedance_table import WIRE_IMPEDANCE_TABLE
from CEDElectrical.circuit_sizing.domain.helpers import normalize_wire_size
from CEDElectrical.circuit_sizing.models.circuit_branch import CircuitBranchModel


class VoltageDropCalculator:
    """Computes voltage drop percentage using impedance lookup tables."""

    @staticmethod
    def calculate(model: CircuitBranchModel, wire_size: Optional[str], sets: int = 1) -> Optional[float]:
        """Return voltage drop fraction for the supplied circuit model."""
        if not wire_size:
            return None

        try:
            length = model.length
            voltage = model.voltage
            pf = model.power_factor or 0.9
            phase = model.phase
            amps = model.circuit_load_current

            if not amps or not length or not voltage:
                return 0

            material = model.wire_info.get("wire_material", "CU")
            conduit_material = model.wire_info.get("conduit_material_type")
            normalized_wire = normalize_wire_size(wire_size, model.settings.wire_size_prefix)
            impedance = WIRE_IMPEDANCE_TABLE.get(normalized_wire)
            if not impedance:
                return None

            R = impedance["R"].get(material, {}).get(conduit_material)
            X = impedance["X"].get(conduit_material)
            if R is None or X is None:
                return None

            R = R / sets
            X = X / sets
            sin_phi = (1 - pf ** 2) ** 0.5

            if phase == 3:
                drop = (1.732 * amps * (R * pf + X * sin_phi) * length) / 1000.0
            else:
                drop = (2 * amps * (R * pf + X * sin_phi) * length) / 1000.0

            return drop / voltage
        except Exception:
            return None
