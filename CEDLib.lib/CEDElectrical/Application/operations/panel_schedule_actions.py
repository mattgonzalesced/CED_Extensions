# -*- coding: utf-8 -*-
"""Panel schedule action operations for staged UI workflows."""

from CEDElectrical.Model.panel_schedule_enums import PanelScheduleOperationKey as OpKey
from CEDElectrical.Model.panel_schedule_manager import PanelScheduleManager


def _panel_option_lookup_from_request(request):
    options = getattr(request, "options", None) or {}
    lookup = options.get("panel_option_lookup")
    if isinstance(lookup, dict):
        return lookup
    return {}


def _placement_from_request(request):
    options = getattr(request, "options", None) or {}
    placement = options.get("placement")
    if isinstance(placement, dict):
        return placement
    return {}


class PanelScheduleAddSpareOperation(object):
    """Operation key for adding SPARE rows."""

    key = OpKey.ADD_SPARE

    def execute(self, request, doc):
        manager = PanelScheduleManager(doc, panel_option_lookup=_panel_option_lookup_from_request(request))
        placement = _placement_from_request(request)
        return manager.apply_add_action(placement)


class PanelScheduleAddSpaceOperation(object):
    """Operation key for adding SPACE rows."""

    key = OpKey.ADD_SPACE

    def execute(self, request, doc):
        manager = PanelScheduleManager(doc, panel_option_lookup=_panel_option_lookup_from_request(request))
        placement = _placement_from_request(request)
        return manager.apply_add_action(placement)


class PanelScheduleRemoveSpareOperation(object):
    """Operation key for removing SPARE rows."""

    key = OpKey.REMOVE_SPARE

    def execute(self, request, doc):
        manager = PanelScheduleManager(doc, panel_option_lookup=_panel_option_lookup_from_request(request))
        placement = _placement_from_request(request)
        return manager.apply_remove_action(placement)


class PanelScheduleRemoveSpaceOperation(object):
    """Operation key for removing SPACE rows."""

    key = OpKey.REMOVE_SPACE

    def execute(self, request, doc):
        manager = PanelScheduleManager(doc, panel_option_lookup=_panel_option_lookup_from_request(request))
        placement = _placement_from_request(request)
        return manager.apply_remove_action(placement)


class PanelScheduleMoveCircuitToPanelOperation(object):
    """Primitive operation key for SelectPanel panel transfer."""

    key = OpKey.MOVE_TO_PANEL

    def execute(self, request, doc):
        manager = PanelScheduleManager(doc, panel_option_lookup=_panel_option_lookup_from_request(request))
        options = getattr(request, "options", None) or {}
        return manager.move_circuit_to_panel(
            circuit_id=int(options.get("circuit_id", 0) or 0),
            target_panel_id=int(options.get("target_panel_id", 0) or 0),
        )


class PanelScheduleMoveCircuitInPanelOperation(object):
    """Primitive operation key for MoveSlotTo within one panel."""

    key = OpKey.MOVE_IN_PANEL

    def execute(self, request, doc):
        manager = PanelScheduleManager(doc, panel_option_lookup=_panel_option_lookup_from_request(request))
        options = getattr(request, "options", None) or {}
        return manager.move_circuit_in_panel(
            panel_id=int(options.get("panel_id", 0) or 0),
            circuit_id=int(options.get("circuit_id", 0) or 0),
            target_slot=int(options.get("target_slot", 0) or 0),
        )


class PanelScheduleMoveCircuitToSpecificSlotOperation(object):
    """Composite operation key for move-to-slot actions."""

    key = OpKey.MOVE_TO_SPECIFIC_SLOT

    def execute(self, request, doc):
        manager = PanelScheduleManager(doc, panel_option_lookup=_panel_option_lookup_from_request(request))
        placement = _placement_from_request(request)
        return manager.apply_move_action(placement)
