# -*- coding: utf-8 -*-
# lib/organized/MEPKit/core/rules.py
from __future__ import absolute_import
import os, re
from organized.MEPKit.core.io import read_json
from organized.MEPKit.core.paths import rules_dir

# ---------- load

def load_identify_rules():
    return read_json(os.path.join(rules_dir(), 'identify_spaces.json'), default={}) or {}

def load_branch_rules():
    return read_json(os.path.join(rules_dir('electrical'), 'branch_circuits.json'), default={}) or {}

# ---------- normalize (fix small typos / aliases)

def normalize_constraints(d):
    """
    Normalize keys to consistent names and convert in→ft where supplied.
    Accepted aliases:
      - avoid_doors_radius_ft | avoid_doors_ft | avoid_doors : float (ft)
      - door_edge_margin_ft | door_edge_margin_in : float
      - avoid_corners_ft
      - door_snap_tolerance_ft
    Also strips accidental trailing ':' in keys.
    """
    if not d: return {}
    nd = {}
    for k, v in d.items():
        key = k.strip().rstrip(':').lower()
        nd[key] = v

    # unify doors radius
    if 'avoid_doors_radius_ft' not in nd:
        for alias in ('avoid_doors_ft', 'avoid_doors'):
            if alias in nd:
                nd['avoid_doors_radius_ft'] = float(nd.get(alias) or 0.0)
                break
    nd.setdefault('avoid_doors_radius_ft', 0.0)

    # in → ft
    if 'door_edge_margin_in' in nd:
        try: nd['door_edge_margin_ft'] = float(nd['door_edge_margin_in'])/12.0
        except: nd['door_edge_margin_ft'] = 0.0
    nd.setdefault('door_edge_margin_ft', float(nd.get('door_edge_margin_ft') or 0.0))

    # defaults
    nd.setdefault('avoid_corners_ft', 2.0)
    nd.setdefault('door_snap_tolerance_ft', 0.05)
    return nd

# ---------- categorize a space by name per identify_spaces.json

def _match_text(text, pattern, use_regex=True, case_sensitive=False):
    if not text: return False
    if use_regex:
        flags = 0 if case_sensitive else re.IGNORECASE
        try: return re.search(pattern, text, flags) is not None
        except: pass  # bad regex → fall through to contains
    return (pattern if case_sensitive else pattern.lower()) in (text if case_sensitive else text.lower())

def _passes_match(text, match_block, defaults):
    inc = (match_block or {}).get('include', []) or []
    exc = (match_block or {}).get('exclude', []) or []
    use_regex = bool(defaults.get('regex', True))
    case_sensitive = bool(defaults.get('case_sensitive', False))

    ok = True if not inc else any(_match_text(text, p, use_regex, case_sensitive) for p in inc)
    if not ok: return False
    if exc and any(_match_text(text, p, use_regex, case_sensitive) for p in exc):
        return False
    return True

def categorize_space_by_name(name, id_rules):
    cats = (id_rules.get('space_categories') or {})
    defaults = id_rules.get('defaults', {})
    for label, spec in cats.items():
        if _passes_match(name or "", (spec or {}).get('match'), defaults):
            return label
    return defaults.get('unknown_category', 'Support')

# ---------- helpers to pull per-category branch circuit config

def get_category_rule(bc_rules, category, fallback='Support'):
    bc = (bc_rules.get('branch_circuits') or {})
    table = (bc.get('receptacle_rules_by_category') or {})
    rule = (table.get(category) or {})
    if not rule and fallback in table:
        rule = table[fallback]
    general = (bc.get('general') or {})
    return rule, general