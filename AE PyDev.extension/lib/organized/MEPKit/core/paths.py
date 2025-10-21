# -*- coding: utf-8 -*-
from __future__ import absolute_import
import os

def _dir_of(path):
    return path if os.path.isdir(path) else os.path.dirname(path)

def find_up(start, *targets):
    cur = _dir_of(start) if start else os.getcwd()
    while True:
        if all(os.path.exists(os.path.join(cur, t)) for t in targets):
            return cur
        parent = os.path.dirname(cur)
        if parent == cur: return None
        cur = parent

def extension_root():
    here = os.path.dirname(os.path.abspath(__file__))
    # assume /rules exists at the extension level
    root = find_up(here, 'rules')
    return root if root else os.path.abspath(os.path.join(here, '..', '..', '..'))

def _here():  # this file
    return os.path.dirname(os.path.abspath(__file__))

def _mepkit_root():
    return os.path.abspath(os.path.join(_here(), '..'))  # .../MEPKit

def _organized_root():
    return os.path.dirname(_mepkit_root())               # .../organized

def rules_dir(sub=None):
    base = os.path.join(_organized_root(), 'rules')       # .../organized/rules
    if sub: return os.path.join(base, sub)
    return base

def configs_dir():
    return os.path.join(extension_root(), 'configs')

def script_dir(__file__like):
    return os.path.dirname(os.path.abspath(__file__like))