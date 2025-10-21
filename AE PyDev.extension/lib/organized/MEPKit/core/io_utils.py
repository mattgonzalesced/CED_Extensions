# -*- coding: utf-8 -*-
from __future__ import absolute_import
import os, json
try:
    import yaml
except:
    yaml = None

def read_json(path, default=None):
    if os.path.exists(path):
        with open(path, 'r') as f: return json.load(f)
    return default

def write_json(path, data, indent=2):
    d = os.path.dirname(path)
    if d and not os.path.exists(d): os.makedirs(d)
    with open(path, 'w') as f: json.dump(data, f, indent=indent)

def read_yaml(path, default=None):
    if not yaml: return default
    if os.path.exists(path):
        with open(path, 'r') as f: return yaml.safe_load(f)
    return default