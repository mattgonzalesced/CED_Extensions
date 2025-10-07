# -*- coding: utf-8 -*-
import logging

# lib/organized/MEPKit/core/log.py
def get_logger(name="MEPKit", level="INFO"):
    try:
        # Use pyRevit's output-aware logger (shows in Output panel)
        from pyrevit import script as _script
        return _script.get_logger()
    except Exception:
        import logging as _logging
        log = _logging.getLogger(name)
        if not log.handlers:
            h = _logging.StreamHandler()
            h.setFormatter(_logging.Formatter('%(levelname)s: %(message)s'))
            log.addHandler(h)
        log.setLevel(getattr(_logging, level.upper(), _logging.INFO))
        return log

def alert(msg, title="MEPKit", warn=False):
    try:
        from pyrevit import forms
        forms.alert(msg, title=title, warn_icon=warn)
    except Exception:
        # Fallback: print if not in pyRevit context
        print("[{}] {}".format(title, msg))