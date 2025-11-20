# -*- coding: utf-8 -*-
"""Equipment grounding conductor sizing."""

from CEDElectrical.refdata.egc_table import EGC_TABLE


class GroundSizer(object):
    def size_ground(self, breaker_rating):
        if breaker_rating is None:
            return None
        try:
            rating = float(breaker_rating)
        except Exception:
            return None

        sorted_keys = sorted(EGC_TABLE.keys())
        for key in sorted_keys:
            if rating <= key:
                return EGC_TABLE[key]
        if sorted_keys:
            return EGC_TABLE[sorted_keys[-1]]
        return None
