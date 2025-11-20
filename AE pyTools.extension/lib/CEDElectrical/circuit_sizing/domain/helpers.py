"""Helper functions for normalizing sizing inputs."""


def normalize_wire_size(value, prefix="#"):
    """Return a normalized wire size without the configured prefix."""
    if not value:
        return None
    return str(value).replace(prefix, "").strip()


def normalize_conduit_size(value, suffix="C"):
    """Return a normalized conduit size without the configured suffix."""
    if not value:
        return None
    return str(value).replace(suffix, "").strip()


def normalize_temperature_rating(value, suffix="C"):
    """Return a normalized conduit size without the configured suffix."""
    if not value:
        return None
    return int(str(value).replace(suffix, "").strip())


