# -*- coding: utf-8 -*-
"""Path helpers for shared UIClasses resource dictionaries."""

import os


ROOT = os.path.abspath(os.path.dirname(__file__))


def get_resources_root():
    return ROOT


def resolve_resource_path(*parts):
    return os.path.abspath(os.path.join(ROOT, *parts))


def themes_root():
    return resolve_resource_path("Themes")


def styles_root():
    return resolve_resource_path("Styles")


def templates_root():
    return resolve_resource_path("Templates")


def icons_root():
    return resolve_resource_path("Icons")


def controls_root():
    return resolve_resource_path("Controls")
