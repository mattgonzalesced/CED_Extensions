# -*- coding: utf-8 -*-
# lib/rules_loader.py
# IronPython/CPython compatible. Place in your pyRevit extension's lib folder.
# Usage example at bottom.

import os, re, json

# ----------------------------
# Minimal YAML support (optional)
# ----------------------------
# If PyYAML is available, we'll use it. If not, users can save .yaml files as .json.
_YAML_AVAILABLE = False
try:
    import yaml  # PyYAML (often not present in IronPython)
    _YAML_AVAILABLE = True
except:
    pass


def _read_text(path):
    with open(path, 'rb') as f:
        data = f.read()
    # Attempt UTF-8 first, then fallback
    try:
        return data.decode('utf-8')
    except:
        try:
            return data.decode('utf-16')
        except:
            return data.decode('latin-1')


def load_yaml_or_json(path):
    """
    Loads a .yaml/.yml using PyYAML if available, otherwise tries JSON.
    If the file extension is .json, uses json loader directly.
    """
    ext = os.path.splitext(path)[1].lower()
    text = _read_text(path)

    if ext in ('.json',):
        return json.loads(text)

    # Try YAML if we can
    if ext in ('.yaml', '.yml') and _YAML_AVAILABLE:
        return yaml.safe_load(text)

    # Fallback: try JSON parse even for .yaml if PyYAML isn't available.
    # (This allows you to simply save your YAML as JSON without changing code paths.)
    try:
        return json.loads(text)
    except Exception as ex:
        raise RuntimeError(
            "Could not parse file: %s. Either install PyYAML for IronPython or provide a .json equivalent.\nOriginal error: %s"
            % (path, ex)
        )


# ----------------------------
# Deep merge helpers
# ----------------------------
def deep_merge(base, override):
    """
    Recursively merge two dict-like structures.
    Lists in 'override' replace lists in 'base'.
    Scalars in 'override' replace scalars in 'base'.
    """
    if base is None:
        return override
    if override is None:
        return base

    if isinstance(base, dict) and isinstance(override, dict):
        result = {}
        for k in set(list(base.keys()) + list(override.keys())):
            if k in base and k in override:
                result[k] = deep_merge(base[k], override[k])
            elif k in base:
                result[k] = base[k]
            else:
                result[k] = override[k]
        return result
    else:
        # For lists/scalars, override wins entirely
        return override


# ----------------------------
# Classification
# ----------------------------
def _compile_matcher(patterns, regex=True, case_sensitive=False):
    """
    Turn a list of strings into compiled regexes or simple lowercase substrings.
    """
    if not patterns:
        return []

    flags = 0 if case_sensitive else re.IGNORECASE
    compiled = []
    for p in patterns:
        if p is None:
            continue
        p = p.strip()
        if not p:
            continue
        if regex:
            try:
                compiled.append(re.compile(p, flags))
            except Exception:
                # If a regex fails, treat it as a literal substring
                compiled.append(p if case_sensitive else p.lower())
        else:
            compiled.append(p if case_sensitive else p.lower())
    return compiled


def _matches_any(name, patterns, regex=True, case_sensitive=False):
    if not patterns:
        return False

    hay = name if case_sensitive else name.lower()
    for pat in patterns:
        if hasattr(pat, 'search'):
            # regex
            if pat.search(name):
                return True
        else:
            # substring
            sub = pat if case_sensitive else pat.lower()
            if sub in hay:
                return True
    return False


def classify_room(room_name, room_cfg):
    """
    room_cfg is the parsed content of identify_rooms.yaml (or .json)
    Returns category string (e.g., 'Offices') or the default fallback.
    """
    if not room_name:
        room_name = ""

    defaults = room_cfg.get('defaults', {})
    unknown_category = defaults.get('unknown_category', 'Support')
    case_sensitive = bool(defaults.get('case_sensitive', False))
    regex_enabled = bool(defaults.get('regex', True))

    cats = room_cfg.get('space_categories', {})
    # Iterate categories in defined order for deterministic behavior
    for category in cats:
        m = cats[category].get('match', {})
        include = m.get('include', []) or []
        exclude = m.get('exclude', []) or []

        inc = _compile_matcher(include, regex_enabled, case_sensitive)
        exc = _compile_matcher(exclude, regex_enabled, case_sensitive)

        # Evaluate
        if include and not _matches_any(room_name, inc, regex_enabled, case_sensitive):
            continue
        if exclude and _matches_any(room_name, exc, regex_enabled, case_sensitive):
            continue

        # If includes are empty, treat as no constraint (rare)
        if not include and exclude and _matches_any(room_name, exc, regex_enabled, case_sensitive):
            continue

        return category

    return unknown_category


# ----------------------------
# Lighting rules retrieval
# ----------------------------
def get_category_rules(category, lighting_cfg):
    """
    Extracts a ruleset for the given category and merges it with defaults.
    Returns a dict ready for consumption by your placement code.
    """
    defaults = lighting_cfg.get('defaults', {}) or {}
    categories = lighting_cfg.get('categories', {}) or {}
    cat_rules = categories.get(category, {}) or {}
    merged = deep_merge(defaults, cat_rules)

    # Normalize a few expected fields to avoid KeyErrors downstream
    if 'fixture_candidates' not in merged:
        merged['fixture_candidates'] = []
    if 'avoid' not in merged:
        merged['avoid'] = {}
    if 'switching' not in merged:
        merged['switching'] = {}
    if 'circuiting' not in merged:
        merged['circuiting'] = {}

    return merged


def build_rule_for_room(room_name, identify_path, lighting_rules_path):
    """
    1) Loads identify_rooms (YAML/JSON) and lighting_rules (YAML/JSON)
    2) Classifies the room_name -> category
    3) Returns merged rule dict for that category

    Returns: (category, rule_dict)
    """
    room_cfg = load_yaml_or_json(identify_path)
    lighting_cfg = load_yaml_or_json(lighting_rules_path)

    category = classify_room(room_name, room_cfg)
    rule = get_category_rules(category, lighting_cfg)
    return category, rule


# ----------------------------
# Example (safe to remove)
# ----------------------------
if __name__ == "__main__":
    # Example paths (adjust to your actual extension)
    identify_path = r"C:\Path\To\rules\identify_spaces.json"
    lighting_path = r"C:\Path\To\rules\lighting_rules.json"

    # Test a few names
    samples = ["Open Office 102", "Conf Room 201", "Corridor L1", "Janitor", "Toilet 115", "WeirdRoom"]

    for name in samples:
        try:
            cat, rule = build_rule_for_room(name, identify_path, lighting_path)
            print("[TEST] %-18s -> %-12s | fixture_candidates=%d | spacing_ft=%s"
                  % (name, cat, len(rule.get('fixture_candidates', [])), str(rule.get('spacing_ft'))))
        except Exception as ex:
            print("[ERROR] %s -> %s" % (name, ex))