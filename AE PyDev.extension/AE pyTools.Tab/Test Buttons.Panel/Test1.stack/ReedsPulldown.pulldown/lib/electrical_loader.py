# -*- coding: utf-8 -*-
# lib/electrical_loader.py
import os, json

def _read_json(path):
    with open(path, 'rb') as f:
        data = f.read()
    for enc in ('utf-8', 'utf-8-sig', 'utf-16', 'latin-1'):
        try:
            return json.loads(data.decode(enc))
        except Exception:
            continue
    raise RuntimeError("Could not parse JSON at {}".format(path))

def _deep_merge(a, b):
    if isinstance(a, dict) and isinstance(b, dict):
        out = dict(a)
        for k, v in b.items():
            out[k] = _deep_merge(out[k], v) if k in out else v
        return out
    return b

def load_electrical_rules(elec_dir):
    """Reads all .json files in rules/electrical, merges them into one dict."""
    if not os.path.isdir(elec_dir):
        return {}
    payload = {}
    files = [f for f in os.listdir(elec_dir) if f.lower().endswith('.json')]
    # Load everything except _local_overrides first
    for name in sorted(files):
        if name.lower() == '_local_overrides.json':
            continue
        path = os.path.join(elec_dir, name)
        payload = _deep_merge(payload, _read_json(path))
    # Apply local overrides last
    ovr = os.path.join(elec_dir, '_local_overrides.json')
    if os.path.exists(ovr):
        payload = _deep_merge(payload, _read_json(ovr))
    return payload

def get_room_profile(category, elec_rules):
    """Convenience view aggregating room-type-related bits for a category."""
    return {
        "loads": (elec_rules.get("loads", {}).get("area_loads_va_per_ft2", {}).get(category)),
        "receptacle_rules": (elec_rules.get("branch_circuits", {}).get("receptacle_rules_by_category", {}).get(category, {})),
        "controls": (elec_rules.get("lighting_controls", {}).get(category, [])),
        "device_protection": (elec_rules.get("device_protection", {}).get("by_category", {}).get(category, {})),
        "wiring_methods": (elec_rules.get("wiring_methods", {}).get("by_category", {}).get(category, [])),
    }
