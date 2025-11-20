# -*- coding: utf-8 -*-
"""Utility helpers for cleaning user overrides."""

from CEDElectrical.refdata.conductor_area_table import CONDUCTOR_AREA_TABLE


class OverrideValidator(object):
    def __init__(self, settings):
        self.settings = settings

    def normalize_wire_size(self, value):
        if value is None:
            return None
        text = str(value).strip()
        text = text.replace('#', '')
        return text

    def clean(self, overrides):
        cleaned = {}
        cleaned.update(overrides or {})

        hot = cleaned.get('wire_hot_size_override')
        neutral = cleaned.get('wire_neutral_size_override')
        ground = cleaned.get('wire_ground_size_override')

        hot = self.normalize_wire_size(hot)
        neutral = self.normalize_wire_size(neutral)
        ground = self.normalize_wire_size(ground)

        if hot not in CONDUCTOR_AREA_TABLE:
            hot = None
        if neutral not in CONDUCTOR_AREA_TABLE:
            neutral = None
        if ground not in CONDUCTOR_AREA_TABLE:
            ground = None

        cleaned['wire_hot_size_override'] = hot
        cleaned['wire_neutral_size_override'] = neutral
        cleaned['wire_ground_size_override'] = ground

        try:
            sets = int(cleaned.get('wire_sets_override') or 1)
            if sets < 1:
                sets = 1
        except Exception:
            sets = 1
        cleaned['wire_sets_override'] = sets

        return cleaned
