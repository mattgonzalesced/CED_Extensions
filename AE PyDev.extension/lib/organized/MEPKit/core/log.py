# -*- coding: utf-8 -*-
import logging

# lib/organized/MEPKit/core/log.py
def get_logger(name="MEPKit", level="INFO"):
    try:
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
        print("[{}] {}".format(title, msg))

# NEW: explicitly open/focus the pyRevit Output panel
def open_output(title="MEPKit Log", header_md=None):
    try:
        from pyrevit import script
        out = script.get_output()
        if title:
            out.set_title(title)
        # these prints FORCE the window to show
        if header_md:
            out.print_md(header_md)
        else:
            out.print_md("### {}".format(title))
        try:
            out.center()
            out.maximize()
        except Exception:
            pass
        return out
    except Exception:
        # Fallback: nothing to open when not in pyRevit context
        return None

__all__ = ["get_logger", "alert", "open_output"]