# -*- coding: utf-8 -*-
# lib/organized/MEPKit/core/log.py
# Minimal, pyRevit-aware logger that also works outside pyRevit.

_levels = {"DEBUG":10, "INFO":20, "WARNING":30, "ERROR":40}

def get_logger(name="MEPKit", level="INFO", title=None):
    """Return a tiny logger with .debug/.info/.warning/.error that prints to the pyRevit Output panel when available."""
    minlvl = _levels.get(str(level).upper(), 20)
    try:
        from pyrevit import script
        out = script.get_output()
        if title:
            out.set_title(title); out.print_md("## " + title)
        class _Logger(object):
            def _w(self, lvl, msg, *a):
                if _levels[lvl] < minlvl: return
                txt = msg.format(*a) if a else str(msg)
                out.write("{}: {}\n".format(lvl, txt))
            def debug(self, msg, *a):   self._w("DEBUG", msg, *a)
            def info(self, msg, *a):    self._w("INFO", msg, *a)
            def warning(self, msg, *a): self._w("WARNING", msg, *a)
            def error(self, msg, *a):   self._w("ERROR", msg, *a)
        return _Logger()
    except Exception:
        # Fallback to plain print
        class _PrintLogger(object):
            def _w(self, lvl, msg, *a):
                if _levels[lvl] < minlvl: return
                txt = msg.format(*a) if a else str(msg)
                print("{}: {}".format(lvl, txt))
            debug = lambda self,m,*a: _PrintLogger._w(self,"DEBUG",m,*a)
            info  = lambda self,m,*a: _PrintLogger._w(self,"INFO",m,*a)
            warning=lambda self,m,*a: _PrintLogger._w(self,"WARNING",m,*a)
            error = lambda self,m,*a: _PrintLogger._w(self,"ERROR",m,*a)
        return _PrintLogger()

def alert(msg, title="MEPKit", warn=False):
    try:
        from pyrevit import forms
        forms.alert(str(msg), title=title, warn_icon=bool(warn))
    except Exception:
        print("[{}] {}".format(title, msg))