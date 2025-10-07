# -*- coding: utf-8 -*-
# lib/organized/MEPKit/core/log.py
# IronPython 2.7 / pyRevit-safe

def open_output(title="MEPKit Log", header_md=None):
    try:
        from pyrevit import script
        out = script.get_output()
        if title:
            out.set_title(title)
        if header_md:
            out.print_md(header_md)
        else:
            out.print_md("### {}".format(title))
        try:
            out.center(); out.maximize()
        except Exception:
            pass
        return out
    except Exception:
        return None

def get_logger(name="MEPKit", level="INFO"):
    # Always try pyRevit logger first and bind it to current Output
    try:
        from pyrevit import script
        out = script.get_output()              # forces an Output target
        log = script.get_logger()              # bound to this Output
        # normalize level
        import logging as _logging
        log.setLevel(getattr(_logging, level.upper(), _logging.INFO))
        return log
    except Exception:
        # Fallback to stdlib logger, but try to route it into pyRevit Output if available
        import logging as _logging
        log = _logging.getLogger(name)
        if not log.handlers:
            try:
                from pyrevit import script
                out = script.get_output()
                class _PyRevitOutputHandler(_logging.Handler):
                    def emit(self, record):
                        try:
                            msg = self.format(record)
                            out.write(msg + "\n")
                        except Exception:
                            pass
                h = _PyRevitOutputHandler()
                h.setFormatter(_logging.Formatter('%(levelname)s: %(message)s'))
                log.addHandler(h)
            except Exception:
                # Final fallback: plain stderr
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

__all__ = ["get_logger", "open_output", "alert"]