# -*- coding: utf-8 -*-
"""Top-level orchestration for circuit calculations."""

from CEDElectrical.Model.CircuitBranch import CircuitCalculationResult
from CEDElectrical.Domain.override_validator import OverrideValidator
from CEDElectrical.Domain.wire_sizer import WireSizer
from CEDElectrical.Domain.ground_sizer import GroundSizer
from CEDElectrical.Domain.conduit_sizer import ConduitSizer
from CEDElectrical.Domain.voltage_drop import VoltageDropCalculator


class CircuitEvaluator(object):
    def __init__(self, settings):
        self.settings = settings
        self.override_validator = OverrideValidator(settings)
        self.wire_sizer = WireSizer(settings)
        self.ground_sizer = GroundSizer()
        self.conduit_sizer = ConduitSizer(settings)
        self.voltage_drop_calculator = VoltageDropCalculator()

    def _calculate_breaker(self, model, overrides):
        if overrides.get('breaker_override'):
            try:
                return int(overrides.get('breaker_override'))
            except Exception:
                pass
        if model.rating:
            return model.rating
        load_current = model.circuit_load_current or 0
        if load_current:
            size = load_current * 1.25
            if size < self.settings.min_breaker_size:
                size = self.settings.min_breaker_size
            return size
        return None

    def evaluate(self, model):
        overrides = self.override_validator.clean(model.overrides)
        result = CircuitCalculationResult(model)

        breaker_rating = self._calculate_breaker(model, overrides)
        result.breaker_rating = breaker_rating

        hot_selection = self.wire_sizer.size_hot_conductor(model, overrides, breaker_rating)
        if hot_selection:
            result.hot_wire_size = hot_selection.get('hot_size')
            result.number_of_sets = hot_selection.get('wire_sets') or 1
            result.circuit_base_ampacity = hot_selection.get('ampacity')
            result.wire_material = hot_selection.get('material')
            result.wire_temp_rating = "%s C" % hot_selection.get('temperature')
            result.wire_insulation = hot_selection.get('insulation')

        include_neutral = overrides.get('include_neutral')
        if include_neutral is None:
            include_neutral = model.base_wire_info.get('include_neutral', True)

        neutral_size = self.wire_sizer.size_neutral(hot_selection, overrides)
        if include_neutral:
            result.neutral_wire_size = neutral_size

        include_ground = True
        ground_size = self.ground_sizer.size_ground(breaker_rating)
        if include_ground:
            result.ground_wire_size = ground_size

        include_isolated_ground = overrides.get('include_isolated_ground')
        if include_isolated_ground is None:
            include_isolated_ground = False
        if include_isolated_ground:
            result.isolated_ground_wire_size = ground_size

        poles = model.poles or 0
        result.hot_wire_quantity = poles * (result.number_of_sets or 1)
        if result.neutral_wire_size:
            result.neutral_wire_quantity = (result.number_of_sets or 1)
        if result.ground_wire_size:
            result.ground_wire_quantity = (result.number_of_sets or 1)
        if result.isolated_ground_wire_size:
            result.isolated_ground_wire_quantity = (result.number_of_sets or 1)

        quantities = {
            'hot': result.hot_wire_quantity or 0,
            'neutral': result.neutral_wire_quantity or 0,
            'ground': result.ground_wire_quantity or 0,
            'isolated_ground': result.isolated_ground_wire_quantity or 0
        }
        result.number_of_wires = sum(quantities.values())

        conduit_type = overrides.get('conduit_type_override') or model.base_wire_info.get('conduit_type') or 'EMT'
        conduit_size, fill_pct = self.conduit_sizer.size_conduit(
            conduit_type,
            result.hot_wire_size,
            result.neutral_wire_size,
            result.ground_wire_size,
            result.isolated_ground_wire_size,
            quantities
        )
        result.conduit_type = conduit_type
        result.conduit_size = conduit_size
        result.conduit_fill_percentage = fill_pct

        conduit_material_type = model.base_wire_info.get('conduit_material_type', 'Non-Magnetic')
        result.voltage_drop_percentage = self.voltage_drop_calculator.calculate_percentage(
            model,
            result.hot_wire_size,
            result.number_of_sets,
            result.wire_material or model.base_wire_info.get('wire_material', 'CU'),
            conduit_material_type
        )

        return result
