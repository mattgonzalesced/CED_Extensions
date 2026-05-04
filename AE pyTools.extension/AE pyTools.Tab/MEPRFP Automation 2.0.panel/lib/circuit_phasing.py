# -*- coding: utf-8 -*-
"""
Pure-Python phase-balance helpers for SuperCircuit.

The legacy V5 didn't actually load-balance — it just picked a 3-phase
distribution system when the group's max-pole count required it,
otherwise let Revit assign whatever was available. The 2.0 rewrite
does the same matching here, plus a deterministic A/B/C round-robin
seed per panel so sequential single-pole circuits don't all stack on
one phase when the user creates a batch in one go.

The Revit-API edge in ``circuit_apply.py`` consumes
``select_distribution_system_id`` to pick the right distribution and
``next_phase_for_panel`` for the seed (Revit then assigns the actual
slot — we just nudge the order).
"""


PHASE_A = "A"
PHASE_B = "B"
PHASE_C = "C"
PHASE_AB = "AB"
PHASE_BC = "BC"
PHASE_CA = "CA"
PHASE_ABC = "ABC"

_SINGLE_PHASE_ROTATION = (PHASE_A, PHASE_B, PHASE_C)
_TWO_POLE_ROTATION = (PHASE_AB, PHASE_BC, PHASE_CA)


# ---------------------------------------------------------------------
# Per-panel phase tracker
# ---------------------------------------------------------------------

class PanelPhaseTracker(object):
    """Keeps running phase indexes per (panel, pole-count) so sequential
    single/two-pole circuits each rotate independently.

    The counter is keyed by ``(panel_lowercase, pole_count)`` so a 2-pole
    circuit doesn't bump the 1-pole rotation and vice versa. Three-pole
    circuits always return ``ABC`` and don't advance any counter.

    Use one tracker per SuperCircuit run; consult
    ``next_phase_for_panel(panel, poles)`` once per group.
    """

    def __init__(self):
        self._counts = {}  # {(panel_lower, pole_count): int}

    def reset(self):
        self._counts = {}

    def next_phase_for_panel(self, panel_name, poles):
        """Return the next phase label for a circuit on ``panel_name``
        with the given pole count.

        Three-phase circuits always get ``ABC``. Two-pole rotates
        AB → BC → CA. Single-pole rotates A → B → C. Each pole count
        keeps its own rotation index per panel so they don't interfere.
        """
        n = int(poles or 1)
        if n >= 3:
            return PHASE_ABC
        panel_key = (panel_name or "").strip().lower() or "__no_panel__"
        key = (panel_key, n)
        idx = self._counts.get(key, 0)
        self._counts[key] = idx + 1
        if n == 2:
            return _TWO_POLE_ROTATION[idx % 3]
        return _SINGLE_PHASE_ROTATION[idx % 3]


# ---------------------------------------------------------------------
# Distribution-system selection (Revit edge)
# ---------------------------------------------------------------------

def select_distribution_system_id(panel_distribution_system_ids, poles):
    """From a panel's supported distribution-system IDs, return the one
    that matches the requested ``poles`` best, or ``None`` if no
    candidates were supplied.

    Selection rule:

      * If ``poles >= 3`` and the list has more than one entry, prefer
        the LAST entry — Revit tends to enumerate single-phase systems
        first and three-phase last.
      * Otherwise, return the first entry.

    The Revit-API edge fetches ``panel.MEPModel.GetAssignedElectricalSystems()``
    or queries the panel's distribution-system property to populate
    the list before calling this.
    """
    ids = list(panel_distribution_system_ids or [])
    if not ids:
        return None
    if int(poles or 1) >= 3 and len(ids) > 1:
        return ids[-1]
    return ids[0]
