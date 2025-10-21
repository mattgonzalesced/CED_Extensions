# -*- coding: utf-8 -*-
import time
from functools import wraps

def timeit(label=None, logger=None):
    def deco(fn):
        @wraps(fn)
        def wrap(*a, **k):
            t0 = time.time()
            try:
                return fn(*a, **k)
            finally:
                dt = (time.time() - t0)*1000.0
                msg = "[timeit] {0} took {1:.1f} ms".format(label or fn.__name__, dt)
                (logger.info if logger else print)(msg)
        return wrap
    return deco