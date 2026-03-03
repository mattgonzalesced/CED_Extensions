# -*- coding: utf-8 -*-
"""Planet Fitness space rules for the CircuitBySpace tool."""

from __future__ import absolute_import


DEFAULT_RULES = {
    "split_by_fixture": False,
    "min_circuits": {"standard": 1, "emergency": 1},
}

SPECIAL_SPACES = {
    "CARDIO/CIRCUIT/STRENGTH/FREEWEIGHTS": {
        "split_by_fixture": True,
        "min_circuits": {"standard": 3, "emergency": 1},
    },
    "FUNCTIONAL TRAINING/MOBILITY": {
        "split_by_fixture": True,
        "min_circuits": {"standard": 2, "emergency": 1},
    },
}


def _normalize(label):
    return (label or "").strip().upper()


def _build_rules(template):
    min_circuits = template.get("min_circuits") or {}
    return {
        "split_by_fixture": bool(template.get("split_by_fixture")),
        "min_circuits": {
            "standard": max(int(min_circuits.get("standard", 1)), 1),
            "emergency": max(int(min_circuits.get("emergency", 1)), 1),
        },
    }


def get_space_rules(space_label):
    """Return grouping rules for the provided space label."""
    normalized = _normalize(space_label)
    for marker, template in SPECIAL_SPACES.items():
        if marker in normalized:
            return _build_rules(template)
    return _build_rules(DEFAULT_RULES)
