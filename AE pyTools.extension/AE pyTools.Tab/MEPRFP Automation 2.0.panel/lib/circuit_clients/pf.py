# -*- coding: utf-8 -*-
"""
Planet Fitness client adapter.

Carries the fitness-equipment load priority so treadmills, powered
bikes, and stairmasters appear first in the preview's group ordering.
No position rules, no extra run-keyword labels, no combined-circuit
override beyond the generic ``&`` splitter on the base.
"""

from circuit_clients.base import CircuitClient


_PF_LOAD_PRIORITY = {
    "TREADMILL": 0,
    "POWERED BIKE": 1,
    "POWERED BIKE1": 1,
    "POWERED BIKE2": 1,
    "STAIRMASTER": 2,
    "SINK 1": 3,
    "SINK 2": 3,
}


class PfClient(CircuitClient):
    key = "pf"
    display_name = "Planet Fitness"

    load_priority = _PF_LOAD_PRIORITY
