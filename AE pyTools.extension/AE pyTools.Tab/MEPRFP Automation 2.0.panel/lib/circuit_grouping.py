# -*- coding: utf-8 -*-
"""
Pure-Python grouping logic for SuperCircuit.

The Revit-API edge collects raw fixture data into ``CircuitItem``
records; this module decides which items share a circuit and produces
``CircuitGroup`` plans that the apply step turns into Revit
``ElectricalSystem`` instances.

Grouping rules (applied in order):

  1. **Dedicated** (``BUCKET_DEDICATED``) — one circuit per item.
  2. **By-parent** (``BUCKET_BYPARENT``) — group by panel +
     parent_element_id. All children of a single parent share one
     circuit.
  3. **Second-by-parent** (``BUCKET_SECONDBYPARENT``) — separate lane
     of the same shape, so a parent can drive two circuits in
     parallel without merging into the primary BYPARENT circuit.
  4. **Custom buckets** (e.g. HEB's ``casecontroller_<prefix>``) —
     grouped by the bucket key, ignoring panel/circuit-number.
  5. **Normal** — group by ``(panel, circuit_number_token)``. Items
     with empty panel or circuit go to a sentinel "needs review"
     bucket that the UI surfaces.

No Revit-API imports. ``CircuitItem`` carries every field the grouper
needs (resolved panel/circuit/load/poles/world_pt/parent id), so unit
tests can run the pipeline without Revit.
"""


# ---------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------

NEEDS_REVIEW_PANEL = "(unassigned)"
NEEDS_REVIEW_CIRCUIT = "(?)"

# Re-exported for caller convenience.
from circuit_clients.base import (  # noqa: E402
    BUCKET_NORMAL,
    BUCKET_DEDICATED,
    BUCKET_BYPARENT,
    BUCKET_SECONDBYPARENT,
)


# ---------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------

class CircuitItem(object):
    """One placed fixture in the circuiting pipeline.

    Filled in stages:

      1. ``circuit_workflow.collect_items`` populates the raw fields
         from the Revit element + Element_Linker + YAML LED.
      2. The client's ``enrich_item`` hook runs.
      3. ``classify_circuit_token`` sets ``bucket`` and
         ``circuit_token`` (with optional ``post_enrich_classify``
         override).
      4. The grouper consumes ``bucket`` + ``circuit_token`` + the
         resolved ``panel_name`` to produce ``CircuitGroup``s.
    """

    __slots__ = (
        # Revit-side identity
        "element", "element_id", "family_name", "type_name",
        # YAML / Element_Linker linkage
        "linker", "led", "profile",
        # Raw CKT data (as captured in the Revit param or the YAML)
        "panel_raw", "circuit_number_raw", "load_name_raw",
        # Parsed / resolved CKT data
        "panel_choices",       # [str, ...] candidates from parsing
        "panel_name",          # resolved single panel name
        "panel_element",       # Revit panel element (set by Revit edge)
        "circuit_token",       # normalized circuit-number token
        "load_name",           # resolved load name
        "rating", "voltage", "poles",
        "schedule_notes",      # CKT_Schedule Notes_CEDT
        # Geometry / hosting
        "world_pt",
        "parent_element_id",
        # Classification
        "bucket",
        # Optional client / context tags
        "client_tags",
        # User overrides applied via the preview UI (mirror the resolved
        # CKT fields; the workflow uses these on apply if non-None).
        "user_panel", "user_circuit_token", "user_load_name",
    )

    def __init__(self, element=None, element_id=None,
                 family_name="", type_name="",
                 linker=None, led=None, profile=None,
                 panel_raw="", circuit_number_raw="", load_name_raw="",
                 panel_choices=None, panel_name="", panel_element=None,
                 circuit_token="", load_name="",
                 rating=None, voltage=None, poles=None,
                 schedule_notes="",
                 world_pt=None, parent_element_id=None,
                 bucket=BUCKET_NORMAL, client_tags=None,
                 user_panel=None, user_circuit_token=None, user_load_name=None):
        self.element = element
        self.element_id = element_id
        self.family_name = family_name or ""
        self.type_name = type_name or ""
        self.linker = linker
        self.led = led
        self.profile = profile
        self.panel_raw = panel_raw or ""
        self.circuit_number_raw = circuit_number_raw or ""
        self.load_name_raw = load_name_raw or ""
        self.panel_choices = list(panel_choices or [])
        self.panel_name = panel_name or ""
        self.panel_element = panel_element
        self.circuit_token = circuit_token or ""
        self.load_name = load_name or load_name_raw or ""
        self.rating = rating
        self.voltage = voltage
        self.poles = poles
        self.schedule_notes = schedule_notes or ""
        self.world_pt = tuple(world_pt) if world_pt else None
        self.parent_element_id = parent_element_id
        self.bucket = bucket
        self.client_tags = dict(client_tags or {})
        self.user_panel = user_panel
        self.user_circuit_token = user_circuit_token
        self.user_load_name = user_load_name

    @property
    def effective_panel(self):
        return self.user_panel if self.user_panel is not None else self.panel_name

    @property
    def effective_circuit_token(self):
        return (
            self.user_circuit_token
            if self.user_circuit_token is not None
            else self.circuit_token
        )

    @property
    def effective_load_name(self):
        return (
            self.user_load_name
            if self.user_load_name is not None
            else self.load_name
        )

    def __repr__(self):
        return "<CircuitItem id={} fam={!r} bucket={!r} panel={!r} ckt={!r}>".format(
            self.element_id, self.family_name,
            self.bucket, self.effective_panel,
            self.effective_circuit_token,
        )


class CircuitGroup(object):
    """A planned circuit — one Revit ``ElectricalSystem`` after apply."""

    __slots__ = (
        "key",            # opaque dedup key — used to merge during edits
        "bucket",
        "members",        # [CircuitItem, ...]
        "panel_name",     # resolved panel name (UI may edit)
        "panel_element",  # set at apply time
        "circuit_token",  # display token; resolved to int by Revit
        "load_name",
        "poles",          # max(member.poles)
        "rating",         # max(member.rating) — applied to RBS_ELEC_CIRCUIT_RATING_PARAM
        "schedule_notes", # first non-empty member.schedule_notes
        "needs_review",   # True if panel or circuit is unresolved
        "warnings",       # list of strings
    )

    def __init__(self, key, bucket, members, panel_name="", panel_element=None,
                 circuit_token="", load_name="", poles=1,
                 rating=None, schedule_notes="",
                 needs_review=False, warnings=None):
        self.key = key
        self.bucket = bucket
        self.members = list(members or [])
        self.panel_name = panel_name or ""
        self.panel_element = panel_element
        self.circuit_token = circuit_token or ""
        self.load_name = load_name or ""
        self.poles = max(1, int(poles or 1))
        self.rating = rating
        self.schedule_notes = schedule_notes or ""
        self.needs_review = bool(needs_review)
        self.warnings = list(warnings or [])

    @property
    def member_count(self):
        return len(self.members)

    def __repr__(self):
        return "<CircuitGroup {!r} panel={!r} ckt={!r} N={}>".format(
            self.bucket, self.panel_name, self.circuit_token, self.member_count
        )


# ---------------------------------------------------------------------
# Grouping
# ---------------------------------------------------------------------

def assemble_groups(items):
    """Return a list of ``CircuitGroup`` plans for the given items.

    Items are grouped per the rules described at the top of this module.
    Order within a group is the input order (so the caller controls
    determinism).
    """
    out = []
    seen_keys = {}

    def _emit(group):
        existing = seen_keys.get(group.key)
        if existing is not None:
            existing.members.extend(group.members)
            existing.poles = max(existing.poles, group.poles)
            # Rating is the *largest* member rating — sizing should
            # cover the heaviest member.
            if group.rating is not None:
                existing.rating = (
                    group.rating if existing.rating is None
                    else max(existing.rating, group.rating)
                )
            # Schedule notes: first non-empty wins. (Most circuits
            # will have homogeneous notes; preserving the first match
            # avoids overwriting a meaningful value with an empty
            # one from a later member.)
            if group.schedule_notes and not existing.schedule_notes:
                existing.schedule_notes = group.schedule_notes
            existing.needs_review = existing.needs_review or group.needs_review
            existing.warnings.extend(group.warnings)
        else:
            seen_keys[group.key] = group
            out.append(group)

    # ----- pre-pass: order items into bucket lanes ------------------

    by_bucket = {
        BUCKET_DEDICATED: [],
        BUCKET_BYPARENT: [],
        BUCKET_SECONDBYPARENT: [],
        BUCKET_NORMAL: [],
    }
    custom_buckets = {}

    for item in items or []:
        if not isinstance(item, CircuitItem):
            continue
        b = item.bucket or BUCKET_NORMAL
        if b in by_bucket:
            by_bucket[b].append(item)
        else:
            custom_buckets.setdefault(b, []).append(item)

    # ----- 1. dedicated --------------------------------------------

    for item in by_bucket[BUCKET_DEDICATED]:
        key = ("ded", _id_or_anon(item))
        _emit(_group_for_single(key, BUCKET_DEDICATED, item))

    # ----- 2. by-parent / by-sibling ------------------------------

    for item in by_bucket[BUCKET_BYPARENT]:
        anchor = item.parent_element_id
        if anchor is None:
            anchor = "no_parent_{}".format(_id_or_anon(item))
        key = ("byp", _resolved_panel(item).lower(), str(anchor))
        _emit(_group_for_member(key, BUCKET_BYPARENT, item))

    # SECONDBYPARENT lives in its own keyspace so it doesn't merge with
    # the primary BYPARENT lane even when panel + parent match.
    for item in by_bucket[BUCKET_SECONDBYPARENT]:
        anchor = item.parent_element_id
        if anchor is None:
            anchor = "no_parent_{}".format(_id_or_anon(item))
        key = ("byp2", _resolved_panel(item).lower(), str(anchor))
        _emit(_group_for_member(key, BUCKET_SECONDBYPARENT, item))

    # ----- 3. custom buckets (e.g. casecontroller_<prefix>) -------

    for bucket_name in sorted(custom_buckets.keys()):
        for item in custom_buckets[bucket_name]:
            key = ("custom", bucket_name, _resolved_panel(item).lower())
            _emit(_group_for_member(key, bucket_name, item))

    # ----- 4. normal ----------------------------------------------

    for item in by_bucket[BUCKET_NORMAL]:
        panel = _resolved_panel(item).strip()
        ckt = (item.effective_circuit_token or "").strip()
        if not panel or not ckt:
            key = ("normal", "review", _id_or_anon(item))
            grp = _group_for_member(key, BUCKET_NORMAL, item)
            grp.needs_review = True
            if not panel:
                grp.warnings.append("Panel unresolved")
                grp.panel_name = NEEDS_REVIEW_PANEL
            if not ckt:
                grp.warnings.append("Circuit number missing")
                grp.circuit_token = NEEDS_REVIEW_CIRCUIT
            _emit(grp)
            continue
        key = ("normal", panel.lower(), ckt.lower())
        _emit(_group_for_member(key, BUCKET_NORMAL, item))

    return out


# ---------------------------------------------------------------------
# Group construction helpers
# ---------------------------------------------------------------------

def _group_for_single(key, bucket, item):
    return CircuitGroup(
        key=key,
        bucket=bucket,
        members=[item],
        panel_name=_resolved_panel(item),
        circuit_token=item.effective_circuit_token,
        load_name=item.effective_load_name,
        poles=item.poles or 1,
        rating=item.rating,
        schedule_notes=item.schedule_notes or "",
    )


def _group_for_member(key, bucket, item):
    return CircuitGroup(
        key=key,
        bucket=bucket,
        members=[item],
        panel_name=_resolved_panel(item),
        circuit_token=item.effective_circuit_token,
        load_name=item.effective_load_name,
        poles=item.poles or 1,
        rating=item.rating,
        schedule_notes=item.schedule_notes or "",
    )


def _resolved_panel(item):
    return item.effective_panel or ""


def _id_or_anon(item):
    return item.element_id if item.element_id is not None else id(item)
