# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import XYZ, Line

def xyz(x=0.0,y=0.0,z=0.0): return XYZ(float(x), float(y), float(z))

def midpoint(p, q):
    return XYZ(0.5*(p.X+q.X), 0.5*(p.Y+q.Y), 0.5*(p.Z+q.Z))

def move_point(p, dx=0, dy=0, dz=0):
    return XYZ(p.X+dx, p.Y+dy, p.Z+dz)

def closest_point_on_line(line, point):
    # line: Autodesk.Revit.DB.Line ; point: XYZ
    p0 = line.GetEndPoint(0); p1 = line.GetEndPoint(1)
    v = p1 - p0
    denom = v.DotProduct(v)
    if denom == 0.0: return p0
    t = (point - p0).DotProduct(v) / denom
    t = max(0.0, min(1.0, t))
    return p0 + t * v