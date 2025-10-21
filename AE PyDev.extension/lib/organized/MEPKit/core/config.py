# -*- coding: utf-8 -*-
# IronPython 2.7-safe
from __future__ import absolute_import
import os, json
from .paths import configs_dir
try:
    import yaml  # optional
except:
    yaml = None

DEFAULTS = {
    "DRY_RUN": False,
    "log_level": "INFO",
}

def _read_yaml(path):
    if not yaml: return None
    if os.path.exists(path):
        with open(path, 'r') as f:
            return yaml.safe_load(f) or {}
    return None

def _read_json(path):
    if os.path.exists(path):
        with open(path, 'r') as f:
            return json.load(f) or {}
    return None

def load_app_config(name='app.yaml'):
    base = configs_dir()
    y = _read_yaml(os.path.join(base, name))
    if y is not None:
        d = DEFAULTS.copy(); d.update(y); return d
    j = _read_json(os.path.join(base, name.replace('.yaml', '.json')))
    d = DEFAULTS.copy()
    if j: d.update(j)
    return d

def get_flag(cfg, key, default=False):
    v = cfg.get(key, default)
    if isinstance(v, basestring):
        return v.strip().lower() in ('1','true','yes','y','on')
    return bool(v)
