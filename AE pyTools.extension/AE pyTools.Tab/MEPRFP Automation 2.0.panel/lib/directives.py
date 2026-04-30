# -*- coding: utf-8 -*-
"""
Parameter directives: BYPARENT / BYSIBLING / static.

A directive lives in a LED's ``parameters`` dict, replacing the static
value for a given parameter name. The on-disk shape (preserved from the
legacy v3/v4 schema, kept at v100 for now)::

    parameters:
      Voltage_CED: 120                                # static
      CKT_Panel_CEDT:
        parent_parameter: PanelName                   # BYPARENT
      CKT_Circuit Number_CEDT:
        sibling_parameter: SET-001-LED-002:CircuitNumber   # BYSIBLING

Sibling references use ``"<sibling_led_id>:<parameter_name>"`` to
disambiguate which LED inside the same set is being referenced.
"""


PARENT_DIRECTIVE_KEY = "parent_parameter"
SIBLING_DIRECTIVE_KEY = "sibling_parameter"


class DirectiveError(Exception):
    pass


# ---------------------------------------------------------------------
# classification
# ---------------------------------------------------------------------

def is_parent_directive(value):
    return isinstance(value, dict) and PARENT_DIRECTIVE_KEY in value


def is_sibling_directive(value):
    return isinstance(value, dict) and SIBLING_DIRECTIVE_KEY in value


def is_directive(value):
    return is_parent_directive(value) or is_sibling_directive(value)


def directive_kind(value):
    """Return ``'static'``, ``'parent'``, or ``'sibling'``."""
    if is_parent_directive(value):
        return "parent"
    if is_sibling_directive(value):
        return "sibling"
    return "static"


# ---------------------------------------------------------------------
# constructors
# ---------------------------------------------------------------------

def static(value):
    """Return the value unchanged (passthrough)."""
    return value


def parent_directive(parent_param_name):
    if not parent_param_name or not isinstance(parent_param_name, str):
        raise DirectiveError("parent_param_name must be a non-empty string")
    return {PARENT_DIRECTIVE_KEY: parent_param_name}


def sibling_directive(sibling_led_id, sibling_param_name):
    if not sibling_led_id or not isinstance(sibling_led_id, str):
        raise DirectiveError("sibling_led_id must be a non-empty string")
    if not sibling_param_name or not isinstance(sibling_param_name, str):
        raise DirectiveError("sibling_param_name must be a non-empty string")
    return {SIBLING_DIRECTIVE_KEY: "{}:{}".format(sibling_led_id, sibling_param_name)}


# ---------------------------------------------------------------------
# accessors
# ---------------------------------------------------------------------

def parent_param_name(value):
    if not is_parent_directive(value):
        return None
    return value.get(PARENT_DIRECTIVE_KEY) or None


def sibling_target(value):
    """Return ``(led_id, param_name)`` for a sibling directive, else ``None``."""
    if not is_sibling_directive(value):
        return None
    raw = value.get(SIBLING_DIRECTIVE_KEY)
    if not isinstance(raw, str) or ":" not in raw:
        return None
    led_id, _, param_name = raw.partition(":")
    led_id = led_id.strip()
    param_name = param_name.strip()
    if not led_id or not param_name:
        return None
    return led_id, param_name


# ---------------------------------------------------------------------
# resolution at audit time
# ---------------------------------------------------------------------

def resolve_expected_value(directive, parent_lookup, sibling_lookup):
    """Compute the expected value for a directive given lookup callables.

    ``parent_lookup(param_name) -> Any | None``
    ``sibling_lookup(led_id, param_name) -> Any | None``

    Returns ``(found, expected_value)``. ``found`` is False if the
    referenced parameter doesn't exist on the parent / sibling.
    """
    if is_parent_directive(directive):
        name = parent_param_name(directive)
        if name is None:
            return False, None
        value = parent_lookup(name)
        return value is not None, value
    if is_sibling_directive(directive):
        target = sibling_target(directive)
        if target is None:
            return False, None
        led_id, param_name = target
        value = sibling_lookup(led_id, param_name)
        return value is not None, value
    # static: directive IS the expected value
    return True, directive
