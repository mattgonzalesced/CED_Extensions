# -*- coding: utf-8 -*-
# lib/organized/MEPKit/core/log.py

_levels = {"DEBUG":10, "INFO":20, "WARNING":30, "ERROR":40}

def get_logger(name="MEPKit", level="INFO", title=None):
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
                # ✅ use print_md (supported), not write()
                try:
                    out.print_md(u"**{}**: {}".format(lvl, txt))
                except Exception:
                    out.print_text(u"{}: {}".format(lvl, txt))
            def debug(self, msg, *a):   self._w("DEBUG", msg, *a)
            def info(self, msg, *a):    self._w("INFO", msg, *a)
            def warning(self, msg, *a): self._w("WARNING", msg, *a)
            def error(self, msg, *a):   self._w("ERROR", msg, *a)
        return _Logger()
    except Exception:
        class _PrintLogger(object):
            def _w(self, lvl, msg, *a):
                if _levels[lvl] < minlvl: return
                txt = msg.format(*a) if a else str(msg)
                print(u"{}: {}".format(lvl, txt))
            debug = lambda self,m,*a: _PrintLogger._w(self,"DEBUG",m,*a)
            info  = lambda self,m,*a: _PrintLogger._w(self,"INFO",m,*a)
            warning=lambda self,m,*a: _PrintLogger._w(self,"WARNING",m,*a)
            error = lambda self,m,*a: _PrintLogger._w(self,"ERROR",m,*a)
        return _PrintLogger()

def alert(msg, title="MEPKit", warn=False):
    try:
        from pyrevit import forms
        forms.alert(unicode(msg), title=title, warn_icon=bool(warn))
    except Exception:
        print(u"[{}] {}".format(title, msg))