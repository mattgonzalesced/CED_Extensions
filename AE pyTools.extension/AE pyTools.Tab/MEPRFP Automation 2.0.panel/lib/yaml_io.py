# -*- coding: utf-8 -*-
"""
YAML parsing and serialization, backed by a vendored copy of PyYAML.

pyRevit's bundled CPython 3 engine has no exposed ``site-packages`` and
neither IronPython 2.7 nor CPython 3 ships PyYAML with pyRevit, so the
panel carries its own copy at ``lib/_vendor/yaml``. The pure-Python
PyYAML loads cleanly under both engines (the C extension fall-back is
silent).

PyYAML handles ``#`` in unquoted keys correctly (a comment requires
preceding whitespace), and quoted ``#`` in values is round-trip safe,
so no manual escape pass is needed for the equipment-definition format.
"""

import os
import sys


_VENDOR_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_vendor")
if _VENDOR_DIR not in sys.path:
    sys.path.insert(0, _VENDOR_DIR)


class YamlError(Exception):
    pass


def _yaml_module():
    try:
        import yaml
    except ImportError as exc:
        raise YamlError(
            "Vendored PyYAML failed to load from {} ({}). "
            "Confirm the package is intact.".format(_VENDOR_DIR, exc)
        )
    return yaml


def parse(text):
    """Parse a YAML string into a Python object.

    Returns ``{}`` for blank or whitespace-only input, never None.
    """
    if text is None or not text.strip():
        return {}
    yaml = _yaml_module()
    try:
        data = yaml.safe_load(text)
    except Exception as exc:
        raise YamlError("Failed to parse YAML: {}".format(exc))
    return data if data is not None else {}


def dump(data):
    """Serialize a Python object to YAML text.

    Block style is used so the result is human-readable. Map ordering is
    preserved when the installed PyYAML supports ``sort_keys``; older
    versions fall back to alphabetical ordering.
    """
    yaml = _yaml_module()
    try:
        try:
            text = yaml.safe_dump(
                data,
                default_flow_style=False,
                sort_keys=False,
                allow_unicode=True,
            )
        except TypeError:
            text = yaml.safe_dump(
                data,
                default_flow_style=False,
                allow_unicode=True,
            )
    except Exception as exc:
        raise YamlError("Failed to serialize YAML: {}".format(exc))

    if text is None:
        text = ""
    if not isinstance(text, str) and hasattr(text, "decode"):
        text = text.decode("utf-8")
    return text
