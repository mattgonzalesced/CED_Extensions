# -*- coding: utf-8 -*-
import logging

def get_logger(name="MEPKit", level="INFO"):
    log = logging.getLogger(name)
    if not log.handlers:
        h = logging.StreamHandler()
        h.setFormatter(logging.Formatter('%(levelname)s: %(message)s'))
        log.addHandler(h)
    log.setLevel(getattr(logging, level.upper(), logging.INFO))
    return log