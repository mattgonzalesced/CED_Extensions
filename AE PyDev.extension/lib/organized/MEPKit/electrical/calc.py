# -*- coding: utf-8 -*-
from __future__ import absolute_import
from organized.MEPKit.electrical.devices import apparent_load_va
from Autodesk.Revit.DB import XYZ
import math

def distance_ft(a, b):
    return math.sqrt((a.X-b.X)**2 + (a.Y-b.Y)**2 + (a.Z-b.Z)**2)

def centroid_of_points(points):
    if not points: return None
    sx = sy = sz = 0.0
    for p in points: sx += p.X; sy += p.Y; sz += p.Z
    n = float(len(points))
    return XYZ(sx/n, sy/n, sz/n)

def total_va(elems, default_each_va=None):
    s = 0.0; used_default = 0
    for e in elems:
        v = apparent_load_va(e)
        if v is None and default_each_va is not None:
            v = float(default_each_va); used_default += 1
        if v is not None: s += float(v)
    return s, used_default
