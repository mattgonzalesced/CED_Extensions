# -*- coding: utf-8 -*-
# tools/agent_tools.py
# Registry for deterministic agent. Each tool exposes: run(ctx, params) -> dict

from place_receptacles import run as place_receptacles_run
from route_circuits import run as route_circuits_run

TOOLS = [
    {
        "name": "place_receptacles",
        "run": place_receptacles_run,
        "defaults": {
            "selection_mode": "all_spaces",
            "avoid_doors": True
        }
    },
    {
        "name": "route_circuits",
        "run": route_circuits_run,
        "defaults": {
            "nearest_panel_only": True
        }
    }
]