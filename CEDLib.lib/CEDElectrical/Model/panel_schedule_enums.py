# -*- coding: utf-8 -*-
"""Shared enum-like constants for panel schedule workflows."""


class PanelSpecialKind(object):
    """Special row kind values used by panel schedule operations."""

    SPARE = "spare"
    SPACE = "space"

    @classmethod
    def all(cls):
        return (cls.SPARE, cls.SPACE)

    @classmethod
    def normalize(cls, value, default=None):
        token = str(value or "").strip().lower()
        if token in cls.all():
            return token
        return default

    @classmethod
    def is_valid(cls, value):
        return cls.normalize(value, None) is not None


class PanelStagedAction(object):
    """Staged placement action values shared by UI + manager layers."""

    ADD_SPARE = "add_spare"
    ADD_SPACE = "add_space"
    REMOVE_SPARE = "remove_spare"
    REMOVE_SPACE = "remove_space"
    MOVE = "move"

    @classmethod
    def normalize(cls, value):
        return str(value or "").strip().lower()

    @classmethod
    def is_add_spare(cls, value):
        return cls.normalize(value).startswith(cls.ADD_SPARE)

    @classmethod
    def is_add_space(cls, value):
        return cls.normalize(value).startswith(cls.ADD_SPACE)

    @classmethod
    def is_remove_spare(cls, value):
        return cls.normalize(value).startswith(cls.REMOVE_SPARE)

    @classmethod
    def is_remove_space(cls, value):
        return cls.normalize(value).startswith(cls.REMOVE_SPACE)


class PanelUiActionType(object):
    """Action types used by panel schedule UI tooling."""

    ADD = "add"
    REMOVE = "remove"

    @classmethod
    def all(cls):
        return (cls.ADD, cls.REMOVE)

    @classmethod
    def normalize(cls, value, default=ADD):
        token = str(value or "").strip().lower()
        if token in cls.all():
            return token
        return default


class PanelUiMode(object):
    """Add/remove mode values used by panel schedule UI tooling."""

    SPARE = "spare"
    SPACE = "space"
    BOTH = "both"
    MIXED = "mixed"

    @classmethod
    def add_modes(cls):
        return (cls.SPARE, cls.SPACE, cls.MIXED)

    @classmethod
    def remove_modes(cls):
        return (cls.SPARE, cls.SPACE, cls.BOTH)

    @classmethod
    def normalize_for_add(cls, value, default=SPACE):
        token = str(value or "").strip().lower()
        if token == cls.BOTH:
            return cls.MIXED
        if token in cls.add_modes():
            return token
        return default

    @classmethod
    def normalize_for_remove(cls, value, default=BOTH):
        token = str(value or "").strip().lower()
        if token in cls.remove_modes():
            return token
        return default


class PanelScheduleOperationKey(object):
    """Operation keys registered in the panel-schedule operation registry."""

    ADD_SPARE = "panel_schedule_add_spare"
    ADD_SPACE = "panel_schedule_add_space"
    REMOVE_SPARE = "panel_schedule_remove_spare"
    REMOVE_SPACE = "panel_schedule_remove_space"
    MOVE_TO_PANEL = "panel_schedule_move_circuit_to_panel"
    MOVE_IN_PANEL = "panel_schedule_move_circuit_in_panel"
    MOVE_TO_SPECIFIC_SLOT = "panel_schedule_move_circuit_to_specific_slot"
