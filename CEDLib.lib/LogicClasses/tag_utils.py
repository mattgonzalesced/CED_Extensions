# -*- coding: utf-8 -*-
"""Helpers related to tag metadata."""


def tag_key_from_dict(tag_dict):
    if not tag_dict:
        return None
    return (
        (tag_dict.get("category") or tag_dict.get("category_name") or "").lower(),
        tag_dict.get("family") or tag_dict.get("family_name"),
        tag_dict.get("type") or tag_dict.get("type_name"),
    )


__all__ = ["tag_key_from_dict"]
