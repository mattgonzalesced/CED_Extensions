# -*- coding: utf-8 -*-
"""View lookup and dedicated-view creation for Circuit Element Finder."""

from pyrevit import DB

try:
    from System.Collections.Generic import List
except Exception:
    List = None


DEDICATED_VIEW_NAME = "_circuitElementFinder"
DEDICATED_3D_VIEW_NAME = "_circuitElementFinder_3D"
EXCLUDED_DISCIPLINE_NAMES = set(["mechanical", "plumbing"])


def _id_value(item):
    try:
        return int(item.IntegerValue)
    except Exception:
        return -1


def _to_name(value):
    try:
        return str(value or "").strip()
    except Exception:
        return ""


def _name_key(value):
    return _to_name(value).lower()


def _discipline_name(value):
    if value is None:
        return ""
    try:
        return _to_name(value.ToString()).lower()
    except Exception:
        return _to_name(value).lower()


def _is_excluded_discipline(view):
    try:
        discipline = getattr(view, "Discipline", None)
    except Exception:
        discipline = None
    name = _discipline_name(discipline)
    if not name:
        return False
    return name in EXCLUDED_DISCIPLINE_NAMES


def _discipline_priority(view):
    try:
        discipline = getattr(view, "Discipline", None)
    except Exception:
        discipline = None
    name = _discipline_name(discipline)
    if name == "electrical":
        return 0
    if name == "coordination":
        return 1
    return 2


def _has_visible_required_categories(view, required_category_ids=None):
    category_ids = list(required_category_ids or [])
    if not category_ids:
        return True
    tested = False
    for category_id in category_ids:
        if not isinstance(category_id, DB.ElementId):
            continue
        tested = True
        try:
            hidden = bool(view.GetCategoryHidden(category_id))
        except Exception:
            # If the view can not report this category state, do not block it.
            return True
        if not hidden:
            return True
    if not tested:
        return True
    return False


def _to_element_id_list(element_ids):
    cleaned = [x for x in list(element_ids or []) if isinstance(x, DB.ElementId)]
    if List is None:
        return cleaned
    try:
        net_list = List[DB.ElementId]()
        for element_id in cleaned:
            net_list.Add(element_id)
        return net_list
    except Exception:
        return cleaned


def _visible_selected_element_ids(doc, view, selected_element_ids=None):
    ids = [x for x in list(selected_element_ids or []) if isinstance(x, DB.ElementId)]
    if not ids:
        return set()

    visible_ids = set()
    collector = DB.FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType()

    filtered_ids = None
    try:
        id_filter = DB.ElementIdSetFilter(_to_element_id_list(ids))
        filtered_ids = collector.WherePasses(id_filter).ToElementIds()
    except Exception:
        filtered_ids = collector.ToElementIds()

    target_values = set([_id_value(x) for x in ids if _id_value(x) > 0])
    for element_id in list(filtered_ids or []):
        elem_value = _id_value(element_id)
        if elem_value <= 0 or elem_value not in target_values:
            continue
        try:
            if bool(view.IsElementHidden(element_id)):
                continue
        except Exception:
            pass
        visible_ids.add(elem_value)
    return visible_ids


def _has_required_visible_elements(doc, view, selected_element_ids=None, require_all_visible=True):
    ids = [x for x in list(selected_element_ids or []) if isinstance(x, DB.ElementId)]
    if not ids:
        return True, 0, 0
    target_values = set([_id_value(x) for x in ids if _id_value(x) > 0])
    total = len(target_values)
    if total <= 0:
        return True, 0, 0
    visible_values = _visible_selected_element_ids(doc, view, ids)
    visible_count = len(visible_values)
    if require_all_visible:
        return visible_count == total, visible_count, total
    return visible_count > 0, visible_count, total


def _all_view_name_keys(doc):
    keys = set()
    collector = DB.FilteredElementCollector(doc).OfClass(DB.View).ToElements()
    for view in list(collector or []):
        name = _name_key(getattr(view, "Name", ""))
        if name:
            keys.add(name)
    return keys


def _pick_unique_name(doc, preferred_name, fallback_prefix):
    preferred = _to_name(preferred_name)
    fallback = _to_name(fallback_prefix) or "View"
    existing = _all_view_name_keys(doc)
    if preferred and _name_key(preferred) not in existing:
        return preferred
    base = fallback
    if _name_key(base) not in existing:
        return base
    idx = 2
    while True:
        candidate = "{}_{}".format(base, idx)
        if _name_key(candidate) not in existing:
            return candidate
        idx += 1


def _find_first_level_id(doc):
    levels = list(DB.FilteredElementCollector(doc).OfClass(DB.Level).ToElements() or [])
    if not levels:
        return DB.ElementId.InvalidElementId
    levels.sort(key=lambda x: float(getattr(x, "Elevation", 0.0) or 0.0))
    try:
        return levels[0].Id
    except Exception:
        return DB.ElementId.InvalidElementId


def _resolve_level_id(doc, preferred_level_id=None):
    if preferred_level_id is not None:
        try:
            level = doc.GetElement(preferred_level_id)
        except Exception:
            level = None
        if isinstance(level, DB.Level):
            return level.Id
    return _find_first_level_id(doc)


def _get_view_family_type_id(doc, view_family):
    collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType).ToElements()
    for family_type in list(collector or []):
        try:
            if family_type.ViewFamily == view_family:
                return family_type.Id
        except Exception:
            continue
    return DB.ElementId.InvalidElementId


def find_existing_plan_view(
    doc,
    preferred_level_id=None,
    required_category_ids=None,
    selected_element_ids=None,
    require_all_visible=True,
):
    preferred_level = _id_value(preferred_level_id)
    candidates = []
    collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewPlan).ToElements()
    for view in list(collector or []):
        try:
            if bool(getattr(view, "IsTemplate", False)):
                continue
        except Exception:
            continue
        if _is_excluded_discipline(view):
            continue
        if not _has_visible_required_categories(view, required_category_ids):
            continue
        ok_visible, visible_count, total_count = _has_required_visible_elements(
            doc,
            view,
            selected_element_ids=selected_element_ids,
            require_all_visible=require_all_visible,
        )
        if not ok_visible:
            continue
        view_type = getattr(view, "ViewType", None)
        if view_type not in (DB.ViewType.FloorPlan, DB.ViewType.EngineeringPlan):
            continue
        score = 1
        try:
            gen_level = getattr(view, "GenLevel", None)
            if gen_level is not None and _id_value(getattr(gen_level, "Id", None)) == preferred_level:
                score = 0
        except Exception:
            pass
        discipline_score = _discipline_priority(view)
        coverage_score = 0 if total_count > 0 and visible_count == total_count else 1
        candidates.append((discipline_score, coverage_score, score, _to_name(getattr(view, "Name", "")), view))
    if not candidates:
        return None
    candidates.sort(key=lambda x: (x[0], x[1], x[2], x[3]))
    return candidates[0][4]


def find_existing_3d_view(
    doc,
    required_category_ids=None,
    selected_element_ids=None,
    require_all_visible=True,
):
    candidates = []
    collector = DB.FilteredElementCollector(doc).OfClass(DB.View3D).ToElements()
    for view in list(collector or []):
        try:
            if bool(getattr(view, "IsTemplate", False)):
                continue
        except Exception:
            continue
        if _is_excluded_discipline(view):
            continue
        if not _has_visible_required_categories(view, required_category_ids):
            continue
        ok_visible, visible_count, total_count = _has_required_visible_elements(
            doc,
            view,
            selected_element_ids=selected_element_ids,
            require_all_visible=require_all_visible,
        )
        if not ok_visible:
            continue
        try:
            if bool(getattr(view, "IsPerspective", False)):
                continue
        except Exception:
            pass
        discipline_score = _discipline_priority(view)
        coverage_score = 0 if total_count > 0 and visible_count == total_count else 1
        candidates.append((discipline_score, coverage_score, _to_name(getattr(view, "Name", "")), view))
    if not candidates:
        return None
    candidates.sort(key=lambda x: (x[0], x[1], x[2]))
    return candidates[0][3]


def _find_dedicated_plan_view(doc):
    valid_names = set([_name_key(DEDICATED_VIEW_NAME)])
    collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewPlan).ToElements()
    for view in list(collector or []):
        if _name_key(getattr(view, "Name", "")) not in valid_names:
            continue
        try:
            if bool(getattr(view, "IsTemplate", False)):
                continue
        except Exception:
            pass
        return view
    return None


def _find_dedicated_3d_view(doc):
    valid_names = set([_name_key(DEDICATED_3D_VIEW_NAME), _name_key(DEDICATED_VIEW_NAME)])
    collector = DB.FilteredElementCollector(doc).OfClass(DB.View3D).ToElements()
    for view in list(collector or []):
        if _name_key(getattr(view, "Name", "")) not in valid_names:
            continue
        try:
            if bool(getattr(view, "IsTemplate", False)):
                continue
        except Exception:
            pass
        return view
    return None


def get_or_create_dedicated_view(doc, view_kind, preferred_level_id=None, logger=None):
    kind = _to_name(view_kind).lower()
    if kind not in ("plan", "3d"):
        raise ValueError("view_kind must be 'plan' or '3d'")

    if kind == "plan":
        existing = _find_dedicated_plan_view(doc)
        if existing is not None:
            return existing, False
        level_id = _resolve_level_id(doc, preferred_level_id)
        if level_id == DB.ElementId.InvalidElementId:
            raise Exception("No valid level found for dedicated plan view.")
        family_type_id = _get_view_family_type_id(doc, DB.ViewFamily.FloorPlan)
        if family_type_id == DB.ElementId.InvalidElementId:
            raise Exception("No FloorPlan ViewFamilyType found.")
        tx = DB.Transaction(doc, "Circuit Element Finder: Create Dedicated Plan")
        tx.Start()
        try:
            view = DB.ViewPlan.Create(doc, family_type_id, level_id)
            view.Name = _pick_unique_name(doc, DEDICATED_VIEW_NAME, DEDICATED_VIEW_NAME)
            tx.Commit()
            return view, True
        except Exception:
            try:
                tx.RollBack()
            except Exception:
                pass
            raise

    existing = _find_dedicated_3d_view(doc)
    if existing is not None:
        return existing, False
    family_type_id = _get_view_family_type_id(doc, DB.ViewFamily.ThreeDimensional)
    if family_type_id == DB.ElementId.InvalidElementId:
        raise Exception("No 3D ViewFamilyType found.")
    tx = DB.Transaction(doc, "Circuit Element Finder: Create Dedicated 3D")
    tx.Start()
    try:
        view = DB.View3D.CreateIsometric(doc, family_type_id)
        existing_name_keys = _all_view_name_keys(doc)
        preferred_name = DEDICATED_VIEW_NAME if _name_key(DEDICATED_VIEW_NAME) not in existing_name_keys else DEDICATED_3D_VIEW_NAME
        view.Name = _pick_unique_name(doc, preferred_name, DEDICATED_3D_VIEW_NAME)
        tx.Commit()
        return view, True
    except Exception:
        try:
            tx.RollBack()
        except Exception:
            pass
        raise
