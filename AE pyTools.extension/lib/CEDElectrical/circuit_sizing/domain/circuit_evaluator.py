from CEDElectrical.circuit_sizing.domain.conduit_sizer import ConduitSizer
from CEDElectrical.circuit_sizing.domain.override_validator import OverrideValidator
from CEDElectrical.circuit_sizing.domain.wire_sizer import WireSizer
from CEDElectrical.circuit_sizing.models.circuit_branch import CircuitBranchModel, CircuitCalculationResult


class CircuitEvaluator:
    """Coordinates sizing logic for a circuit model."""

    @staticmethod
    def evaluate(model):
        overrides = OverrideValidator(model).cleaned_overrides()
        model.overrides = overrides

        wire_sizer = WireSizer(model)
        results = wire_sizer.evaluate(overrides)

        conduit_sizer = ConduitSizer(model, results)
        if overrides.conduit_size_override:
            results.conduit.size = overrides.conduit_size_override
        else:
            conduit_sizer.calculate_conduit()
        conduit_sizer.calculate_fill()

        return results
