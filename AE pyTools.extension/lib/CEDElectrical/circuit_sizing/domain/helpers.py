from __future__ import annotations

from typing import Optional


def normalize_wire_size(value: Optional[str], prefix: str = "#") -> Optional[str]:
    if not value:
        return None
    return str(value).replace(prefix, "").strip()


def normalize_conduit_size(value: Optional[str], suffix: str = "C") -> Optional[str]:
    if not value:
        return None
    return str(value).replace(suffix, "").strip()
