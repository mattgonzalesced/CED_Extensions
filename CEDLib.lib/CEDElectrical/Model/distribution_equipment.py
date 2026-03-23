# -*- coding: utf-8 -*-
"""Revit-agnostic distribution equipment models."""


class DistributionEquipment(object):
    """Domain model for electrical distribution equipment metadata."""

    def __init__(self, **kwargs):
        self.id = int(kwargs.get("id", 0) or 0)
        self.name = kwargs.get("name")
        self.equipment_type = kwargs.get("equipment_type")
        self.part_type = kwargs.get("part_type")

        self.voltage = kwargs.get("voltage")
        self.poles = kwargs.get("poles")
        self.distribution_system = kwargs.get("distribution_system")
        self.distribution_system_secondary = kwargs.get("distribution_system_secondary")

        self.supply_circuits = list(kwargs.get("supply_circuits") or [])
        self.branch_circuits = list(kwargs.get("branch_circuits") or [])
        self.branch_circuit_options = list(kwargs.get("branch_circuit_options") or [])

        self.mains_rating = kwargs.get("mains_rating")
        self.mains_type = kwargs.get("mains_type")
        self.has_ocp = kwargs.get("has_ocp")
        self.ocp_type = kwargs.get("ocp_type")
        self.ocp_rating = kwargs.get("ocp_rating")

        self.has_feed_thru_lugs = kwargs.get("has_feed_thru_lugs")
        self.has_neutral_bus = kwargs.get("has_neutral_bus")
        self.has_ground_bus = kwargs.get("has_ground_bus")
        self.has_isolated_ground_bus = kwargs.get("has_isolated_ground_bus")

        self.max_poles = kwargs.get("max_poles")
        self.short_circuit_rating = kwargs.get("short_circuit_rating")

        self.power_connected_total = kwargs.get("power_connected_total")
        self.current_connected_total = kwargs.get("current_connected_total")
        self.power_demand_total = kwargs.get("power_demand_total")
        self.current_demand_total = kwargs.get("current_demand_total")

    @property
    def ID(self):
        """Compatibility alias for id."""
        return self.id

    @property
    def Name(self):
        """Compatibility alias for name."""
        return self.name

    @property
    def EquipmentType(self):
        """Compatibility alias for equipment_type."""
        return self.equipment_type

    def to_dict(self):
        """Serialize model data for UI/report consumers."""
        return dict(self.__dict__)


class Transformer(DistributionEquipment):
    """Distribution equipment specialization for transformers."""

    def __init__(self, **kwargs):
        DistributionEquipment.__init__(self, **kwargs)
        self.xfmr_rating = kwargs.get("xfmr_rating")
        self.xfmr_impedance = kwargs.get("xfmr_impedance")
        self.xfmr_kfactor = kwargs.get("xfmr_kfactor")


class PowerBus(DistributionEquipment):
    """Distribution equipment specialization for panel/switch/data buses."""

    def __init__(self, **kwargs):
        DistributionEquipment.__init__(self, **kwargs)
        self.has_panel_schedule = bool(kwargs.get("has_panel_schedule", False))
        self.panel_configuration = kwargs.get("panel_configuration")

    def has_panelschedule(self):
        """Return True when this bus has a panel schedule instance in the model."""
        return bool(self.has_panel_schedule)
