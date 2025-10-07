# -*- coding: utf-8 -*-
# lib/rooms.py
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, SpatialElement, Options,
    BoundingBoxXYZ, XYZ
)

def get_room_level(doc, room):
    try:
        lid = room.LevelId
        return doc.GetElement(lid) if lid and lid.IntegerValue > 0 else None
    except:
        return None

def get_current_view_level(doc, active_view):
    try:
        return doc.GetElement(active_view.GenLevel.Id)
    except:
        return None

def get_selected_rooms_first(doc, uidoc):
    if uidoc is None:
        return []
    ids = list(uidoc.Selection.GetElementIds())
    out = []
    for eid in ids:
        el = doc.GetElement(eid)
        if isinstance(el, SpatialElement) and el.Category and el.Category.Id.IntegerValue == int(BuiltInCategory.OST_Rooms):
            out.append(el)
    return out

def collect_all_rooms(doc):
    return list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType())

def filter_rooms_to_view_level(doc, rooms, level):
    if not level:
        return rooms
    out = []
    for r in rooms:
        rlvl = get_room_level(doc, r)
        if rlvl and rlvl.Id == level.Id:
            out.append(r)
    return out

def get_target_rooms(doc, uidoc, active_view, only_current_level=False, prefer_selection=True):
    sel = get_selected_rooms_first(doc, uidoc) if prefer_selection else []
    if sel:
        return sel
    rooms = collect_all_rooms(doc)
    if only_current_level:
        lvl = get_current_view_level(doc, active_view)
        rooms = filter_rooms_to_view_level(doc, rooms, lvl)
    return rooms

def room_display_name(room):
    try:
        p = room.LookupParameter("Name")
        nm = p.AsString() if p else room.Name
    except:
        nm = getattr(room, "Name", "")
    num = ""
    try:
        q = room.LookupParameter("Number")
        num = q.AsString() if q else ""
    except:
        pass
    return (nm or "").strip(), (num or "").strip()

def get_room_bbox_center(doc, room, active_view):
    # try location point
    try:
        loc = room.Location
        if loc and hasattr(loc, "Point") and loc.Point:
            return loc.Point
    except:
        pass
    # fallback: geometry bbox
    try:
        opt = Options()
        geo = room.get_Geometry(opt)
        bbox = None
        for g in geo:
            try:
                bb = g.GetBoundingBox()
                if bb:
                    if bbox is None:
                        bbox = BoundingBoxXYZ(); bbox.Min = bb.Min; bbox.Max = bb.Max
                    else:
                        bbox.Min = XYZ(min(bbox.Min.X, bb.Min.X),
                                       min(bbox.Min.Y, bb.Min.Y),
                                       min(bbox.Min.Z, bb.Min.Z))
                        bbox.Max = XYZ(max(bbox.Max.X, bb.Max.X),
                                       max(bbox.Max.Y, bb.Max.Y),
                                       max(bbox.Max.Z, bb.Max.Z))
            except:
                continue
        if bbox:
            return (bbox.Min + bbox.Max) * 0.5
    except:
        pass
    return XYZ(0, 0, 0)