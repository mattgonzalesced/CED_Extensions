# -*- coding: utf-8 -*-
"""Single entry point that runs every stage 0 test module.

Skips the existing ``_roundtrip_test.py`` (yaml_io / real-file harness)
and modules that need the Revit API (``links.py``).
"""

from __future__ import print_function

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import _test_geometry
import _test_element_linker
import _test_truth_groups
import _test_migrations
import _test_profile_model
import _test_directives
import _test_capture_idgen
import _test_placement_matching
import _test_merge
import _test_circuit_grouping


def main():
    all_fails = {}
    for module in (
        _test_geometry,
        _test_element_linker,
        _test_truth_groups,
        _test_migrations,
        _test_profile_model,
        _test_directives,
        _test_capture_idgen,
        _test_placement_matching,
        _test_merge,
        _test_circuit_grouping,
    ):
        name = module.__name__.replace("_test_", "")
        print("\n========= {} =========".format(name))
        fails = module.run()
        if fails:
            all_fails[name] = fails

    print("\n=========================================")
    if not all_fails:
        print("ALL TESTS PASSED")
        return 0
    print("FAILURES:")
    for name, fails in all_fails.items():
        print("  {}: {}".format(name, fails))
    return 1


if __name__ == "__main__":
    sys.exit(main())
