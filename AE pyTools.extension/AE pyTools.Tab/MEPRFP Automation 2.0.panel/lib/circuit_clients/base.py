# -*- coding: utf-8 -*-
"""
``CircuitClient`` — base class every per-client adapter inherits from.

The base implements sensible default behaviour for everything; clients
override only the bits that genuinely differ. Adding a new client is
typically a 30-50 line module that overrides a method or two.

The class is split into four contract surfaces:

  1. **Identity** — ``key``, ``display_name`` (mandatory).
  2. **Element scope** — ``should_circuit(elem, item)``: filter step.
  3. **CKT-string parsing** — panel + circuit-number normalisation.
  4. **Item enrichment** — pre-grouping mutations: panel-choice
     enrichment from spaces, load-name decoration, classification of
     special tokens (``DEDICATED`` / ``BYPARENT`` / etc.).

Pure Python — no Revit-API imports here. The default panel-string
parser and classify pipeline are good enough for any client whose data
follows the convention; clients only override when their CAD authors
encoded something extra (HEB's ``CASECONTROLLER_<prefix>`` token, etc.).
"""

import re


# ---------------------------------------------------------------------
# Constants — circuit-bucket keys used by the grouping engine.
# ---------------------------------------------------------------------

BUCKET_NORMAL = "normal"
BUCKET_DEDICATED = "dedicated"
BUCKET_BYPARENT = "byparent"
BUCKET_SECONDBYPARENT = "secondbyparent"

ALL_BUCKETS = (
    BUCKET_NORMAL,
    BUCKET_DEDICATED,
    BUCKET_BYPARENT,
    BUCKET_SECONDBYPARENT,
)


# Tokens recognised as "circuit number says inherit / dedicate" rather
# than "circuit number is N". Compared case-insensitively after
# whitespace strip. The universal set is intentionally minimal —
# DEDICATED, BYPARENT, SECONDBYPARENT only. Client-specific tokens
# (CASECONTROLLER_<prefix>, position-rule keywords) live on the client
# class.
_DEDICATED_TOKENS = frozenset({"dedicated"})
_BYPARENT_TOKENS = frozenset({
    "byparent", "circuitbyparent",
})
_SECONDBYPARENT_TOKENS = frozenset({
    "secondbyparent", "secondcircuitbyparent",
})

# Default panel-string token splitter — comma, semicolon, pipe, slash,
# and whitespace runs.
_PANEL_SPLIT_RE = re.compile(r"[,;|/\s]+")


# Combined-circuit token: ``"6&9"`` style splitter. Honoured by both
# HEB and PF in the legacy panel — the ``&`` separator means "split
# this group's members across the listed circuit numbers". Pattern
# returns the parts (already trimmed) when at least one ``&`` is
# present and at least two non-empty parts result.
_COMBINED_CIRCUIT_SPLIT_RE = re.compile(r"\s*&\s*")


# ---------------------------------------------------------------------
# Base class
# ---------------------------------------------------------------------

class CircuitClient(object):
    """Default implementation. Subclasses override what they need."""

    # --- identity ----------------------------------------------------

    key = ""           # stable lowercase id, e.g. "heb"
    display_name = ""  # human label for picker dialogs

    # --- declarative data ------------------------------------------

    # Position-rule definitions. Each rule is a dict with keys:
    #
    #   keyword          (required) substring matched against the
    #                    circuit_number string (case-insensitive).
    #   group_size       (optional, int) max members per spatial cluster.
    #   label            (optional, str) load_name override applied to
    #                    every member of the resulting group.
    #   cluster_radius   (optional, float ft) used by load-aware
    #                    rules instead of group_size.
    #   max_group_load   (optional, float VA) cap when clustering by
    #                    radius — additional members start a new cluster.
    #   include_singles  (optional, bool) emit singleton clusters
    #                    rather than dropping unmatched members.
    #
    # The grouping engine consults this list to pre-bucket items
    # before the standard normal/dedicated/byparent passes.
    position_rules = ()

    # Load-name priority lookup. Keys are upper-case load names; values
    # are sort-priority ints (lower = appears earlier in preview).
    # Defaults to empty so every load tied with priority ``DEFAULT_LOAD_PRIORITY``.
    load_priority = {}

    # Run-keyword filter dropdown labels in addition to the standard
    # bucket tokens. The preview UI surfaces these so the user can
    # filter the row list by literal circuit_number tokens (e.g.
    # ``EMERGENCY``).
    extra_run_keywords = ()

    def __repr__(self):
        return "<CircuitClient {!r}>".format(self.key or "default")

    # --- element scope ----------------------------------------------

    def should_circuit(self, elem, item=None):
        """Return True if ``elem`` should participate in circuiting.

        Default: always True. Subclasses can return False for elements
        the client wants to skip (e.g. lighting fixtures handled by a
        separate workflow).
        """
        return True

    # --- panel-string parsing ---------------------------------------

    def parse_panel_string(self, panel_str):
        """Tokenise a raw ``CKT_Panel_CEDT`` string into a list of
        candidate panel names.

        Default: split on common separators, strip each token, drop
        empties, preserve order. Matches the legacy V5 splitter.
        """
        if not panel_str:
            return []
        tokens = _PANEL_SPLIT_RE.split(str(panel_str))
        out = []
        seen = set()
        for tok in tokens:
            t = (tok or "").strip()
            if not t:
                continue
            key = t.lower()
            if key in seen:
                continue
            seen.add(key)
            out.append(t)
        return out

    # --- circuit-number classification ------------------------------

    def classify_circuit_token(self, circuit_str):
        """Return a ``(bucket, normalized_token)`` pair.

        ``bucket`` is one of ``BUCKET_*`` constants; ``normalized_token``
        is the trimmed string for grouping purposes. For a numeric
        circuit, bucket is ``BUCKET_NORMAL`` and the token is the digits.
        For unparseable / blank, bucket is ``BUCKET_NORMAL`` with an
        empty token (caller decides what to do).
        """
        text = (circuit_str or "").strip()
        if not text:
            return (BUCKET_NORMAL, "")
        upper = text.upper()
        lower = text.lower()
        if lower in _DEDICATED_TOKENS:
            return (BUCKET_DEDICATED, upper)
        if lower in _BYPARENT_TOKENS:
            return (BUCKET_BYPARENT, upper)
        if lower in _SECONDBYPARENT_TOKENS:
            return (BUCKET_SECONDBYPARENT, upper)
        # Treat anything else (numeric, "DEDICATED-A", "TVTRUSS", etc.)
        # as a normal circuit bucket; group key uses the token verbatim.
        return (BUCKET_NORMAL, text)

    # --- item enrichment hooks --------------------------------------
    #
    # The workflow calls these in sequence after raw CKT data has been
    # extracted from each element. Default implementations are no-ops;
    # subclasses override when the CAD authors encoded client-specific
    # conventions (HEB's case-controller groups, etc.).

    def enrich_item(self, item, context):
        """Mutate the item in place after CKT data is parsed.

        ``context`` carries shared lookups (panel_index, space_index,
        ...) so the client doesn't have to re-walk the doc.

        Default: no-op.
        """
        return None

    def post_enrich_classify(self, item):
        """Optional second-pass classification after ``enrich_item``.

        A client that wants ``CASECONTROLLER_<prefix>`` tokens to land
        in a custom bucket can return a ``(bucket, token)`` pair here;
        returning ``None`` keeps the bucket assigned by
        ``classify_circuit_token``.

        Default: no override.
        """
        return None

    # --- panel-choice ranking ---------------------------------------

    def rank_panel_choices(self, item, context):
        """Re-order ``item.panel_choices`` so the first entry is the
        client's preferred panel for this fixture.

        Default: leave the list in tokenization order. HEB overrides to
        sort by spatial distance to the candidate panels.
        """
        return item.panel_choices

    # --- load-name decoration ---------------------------------------

    def decorate_load_name(self, item, context):
        """Optionally append a suffix to ``item.load_name`` based on
        spatial / contextual data.

        Default: no change.
        """
        return item.load_name

    # --- combined-circuit (`&`) handling ----------------------------

    def split_combined_circuit_token(self, circuit_str):
        """Parse a ``"6&9"`` style combined-circuit token.

        Returns ``[part, ...]`` (each part trimmed) when the input
        contains at least one ``&`` and yields two or more non-empty
        parts. Returns ``[]`` otherwise — caller should fall back to
        single-circuit handling.

        The grouper calls this so a member list with circuit_number
        ``"6&9"`` gets fanned out across two circuits (one per part).
        Subclasses that want to treat ``&`` differently can override.
        """
        if not circuit_str or "&" not in str(circuit_str):
            return []
        parts = [p.strip() for p in _COMBINED_CIRCUIT_SPLIT_RE.split(str(circuit_str))]
        parts = [p for p in parts if p]
        return parts if len(parts) >= 2 else []

    def combined_circuit_load_overrides(self, parts):
        """Optional load-name override map for a combined-circuit split.

        Receives the trimmed parts (e.g. ``["6", "9"]``) and returns a
        ``{normalized_part: load_label}`` dict. Default base: no
        override. The ``"6"&"9"`` -> SINK 1 / SINK 2 special case lives
        on the HEB/PF clients that ship it.
        """
        return {}

    # --- position-rule matching -------------------------------------

    def match_position_rule(self, circuit_str):
        """Return the first position-rule whose ``keyword`` is a
        substring of ``circuit_str`` (case-insensitive), or ``None``.

        Rules are checked longest-keyword first so a more-specific
        match wins (``"CHECKSTAND RECEPT"`` beats ``"CHECKSTAND"``).
        """
        text = (circuit_str or "").strip().upper()
        if not text:
            return None
        # Sort by descending keyword length once; cache on the class.
        sorted_rules = self._sorted_position_rules()
        for rule in sorted_rules:
            kw = (rule.get("keyword") or "").strip().upper()
            if kw and kw in text:
                return rule
        return None

    def _sorted_position_rules(self):
        cached = getattr(self, "_position_rules_sorted_cache", None)
        if cached is None:
            cached = sorted(
                list(self.position_rules or ()),
                key=lambda r: -len((r.get("keyword") or "")),
            )
            try:
                self._position_rules_sorted_cache = cached
            except AttributeError:
                pass
        return cached

    # --- load priority ----------------------------------------------

    def get_load_priority(self, load_name, default=99):
        """Return the sort-priority for ``load_name`` per
        ``self.load_priority``. Lower number = appears earlier in the
        preview list. Unknown load names get ``default``.
        """
        if not load_name:
            return default
        key = str(load_name).strip().upper()
        return int(self.load_priority.get(key, default))

    # --- run-keyword dropdown ---------------------------------------

    def run_keyword_options(self, present_in_groups=()):
        """Return the run-keyword filter labels for the preview UI.

        ``present_in_groups`` is the de-duplicated set of tokens
        actually appearing on the current group list (so the dropdown
        also surfaces ad-hoc tokens like a literal ``"NIGHTLIGHT"``).
        Adds the standard tokens + ``self.extra_run_keywords``.
        """
        seen = set()
        out = []

        def _add(token):
            t = (token or "").strip().upper()
            if not t or t in seen:
                return
            seen.add(t)
            out.append(t)

        for token in present_in_groups or ():
            _add(token)
        for token in ("DEDICATED", "BYPARENT", "SECONDBYPARENT"):
            _add(token)
        for token in self.extra_run_keywords or ():
            _add(token)
        return out
