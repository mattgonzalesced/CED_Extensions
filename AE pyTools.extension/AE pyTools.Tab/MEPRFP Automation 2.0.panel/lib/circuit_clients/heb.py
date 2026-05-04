# -*- coding: utf-8 -*-
"""
HEB client adapter.

Inherits the default base behaviour and adds:

  * **CASECONTROLLER_<prefix> grouping**: when the captured
    ``CKT_Circuit Number_CEDT`` carries a ``CASECONTROLLER_<prefix>``
    token, classify the item into a per-prefix bucket so the
    grouper bundles all controllers of the same prefix into one
    circuit (instead of treating the literal token as a circuit
    number).

  * **Distance-ranked multi-panel choices**: when ``CKT_Panel_CEDT``
    parses to multiple candidates (e.g. ``"L1, L3"``), reorder them so
    the panel physically nearest the fixture is first. Pre-existing
    legacy behaviour, minus the BA/DA spatial-mapping hack which the
    user explicitly dropped for the 2.0 rewrite.

  * **Load-name space tag**: append the enclosing space's name to
    ``CKT_Load Name_CEDT`` when the fixture sits inside a recognised
    space. Cleaner than the legacy keyword-map approach — uses
    whichever Revit space the fixture is actually in.

The HEB-specific BA/DA panel-from-space lookup and the case-rotation
phase trick are *not* ported. The base classifier covers everything
else.
"""

import math
import re

from circuit_clients.base import (
    CircuitClient,
    BUCKET_NORMAL,
)


# ``CASECONTROLLER_<prefix>`` matches the entire token; prefix is
# returned for bucketing. Case-insensitive, leading/trailing whitespace
# tolerated by the caller.
_CASECONTROLLER_RE = re.compile(
    r"^\s*CASECONTROLLER[_\-:\s]+(?P<prefix>[A-Z0-9_\-]+)\s*$",
    re.IGNORECASE,
)


def _casecontroller_prefix(text):
    if not text:
        return None
    m = _CASECONTROLLER_RE.match(str(text))
    if not m:
        return None
    prefix = (m.group("prefix") or "").strip().upper()
    return prefix or None


_HEB_POSITION_RULES = (
    {"keyword": "CHECKSTAND RECEPT", "group_size": 2},
    {"keyword": "CHECKSTAND JBOX", "group_size": 2},
    {"keyword": "SELF CHECKOUT", "group_size": 2},
    {"keyword": "TABLE", "group_size": 3, "label": "Grouped Tables"},
    {"keyword": "DESK QUAD", "group_size": 3, "label": "Grouped Desks - Quad"},
    {"keyword": "DESK DUPLEX", "group_size": 3, "label": "Grouped Desks - Duplex"},
    {"keyword": "ARTISAN BREAD", "group_size": 3},
    {
        "keyword": "ELECTRIC CARTS",
        "cluster_radius": 5.0,
        "max_group_load": 1800.0,
        "include_singles": True,
    },
)


_HEB_EXTRA_RUN_KEYWORDS = ("CASECONTROLLER",)


class HebClient(CircuitClient):
    key = "heb"
    display_name = "HEB"

    position_rules = _HEB_POSITION_RULES
    extra_run_keywords = _HEB_EXTRA_RUN_KEYWORDS

    # ----- circuit-number classification --------------------------------

    def post_enrich_classify(self, item):
        """If the original circuit-number string was a
        ``CASECONTROLLER_<prefix>`` token, bucket the item under a
        synthetic ``casecontroller_<prefix>`` bucket so the grouper
        bundles all controllers sharing a prefix into one circuit.
        """
        prefix = _casecontroller_prefix(getattr(item, "circuit_number_raw", ""))
        if prefix is None:
            return None
        bucket = "casecontroller_{}".format(prefix.lower())
        token = "CC_{}".format(prefix)
        return (bucket, token)

    # ----- panel ranking ------------------------------------------------

    def rank_panel_choices(self, item, context):
        """Sort ``item.panel_choices`` by Euclidean distance from the
        fixture to the panel's location. Falls back to the existing
        order when distances aren't computable.
        """
        choices = list(item.panel_choices or [])
        if len(choices) <= 1:
            return choices
        fixture_pt = getattr(item, "world_pt", None)
        panels_by_name = (context or {}).get("panels_by_name") or {}
        if not fixture_pt or not panels_by_name:
            return choices

        def _key(name):
            target = panels_by_name.get((name or "").strip().lower())
            if target is None:
                return float("inf")
            pt = getattr(target, "world_pt", None)
            if not pt:
                return float("inf")
            dx = float(fixture_pt[0]) - float(pt[0])
            dy = float(fixture_pt[1]) - float(pt[1])
            dz = float(fixture_pt[2]) - float(pt[2])
            return math.sqrt(dx * dx + dy * dy + dz * dz)

        sorted_choices = sorted(choices, key=_key)
        return sorted_choices

    # ----- load-name decoration ----------------------------------------

    def decorate_load_name(self, item, context):
        """If the fixture is inside a Revit space, append the space's
        name as a suffix on the load name (e.g. ``"Outlet 1 - BAKERY"``).
        Avoids duplication if the suffix is already present.
        """
        space_name = (context or {}).get("space_name_for_item", lambda _: None)(item)
        if not space_name:
            return item.load_name
        space_name = str(space_name).strip()
        if not space_name:
            return item.load_name
        load = (item.load_name or "").strip()
        if not load:
            return space_name.upper()
        suffix = " - {}".format(space_name.upper())
        if load.upper().endswith(suffix.upper()):
            return load
        return load + suffix
