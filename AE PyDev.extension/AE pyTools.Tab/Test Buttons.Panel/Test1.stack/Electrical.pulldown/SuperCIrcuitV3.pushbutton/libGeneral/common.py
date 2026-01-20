from pyrevit import DB
from pyrevit.revit.db import query

try:
    basestring
except NameError:
    basestring = str

try:
    long
except NameError:
    long = int


def safe_strip(value):
    return value.strip() if isinstance(value, basestring) else value


def get_param_value(element, param_name):
    param = query.get_param(element, param_name)
    return safe_strip(query.get_param_value(param)) if param else None


def get_element_location(element):
    location = getattr(element, "Location", None)
    if location is not None:
        point = getattr(location, "Point", None)
        if point:
            return point
        curve = getattr(location, "Curve", None)
        if curve:
            return curve.Evaluate(0.5, True)
    bbox = element.get_BoundingBox(None)
    if bbox:
        return DB.XYZ(
            (bbox.Min.X + bbox.Max.X) * 0.5,
            (bbox.Min.Y + bbox.Max.Y) * 0.5,
            (bbox.Min.Z + bbox.Max.Z) * 0.5,
        )
    return DB.XYZ.Zero


def try_parse_int(value):
    if value is None:
        return None
    if isinstance(value, (int, long)):
        return int(value)
    try:
        text = str(value).strip()
        digits = []
        for ch in text:
            if ch.isdigit():
                digits.append(ch)
            elif digits:
                break
        if not digits:
            return None
        return int("".join(digits))
    except Exception:
        return None


def try_parse_float(value):
    if value is None:
        return None
    if isinstance(value, (int, long, float)):
        return float(value)
    try:
        text = str(value).strip()
        cleaned = []
        decimal_found = False
        for ch in text:
            if ch.isdigit():
                cleaned.append(ch)
            elif ch in (".", ","):
                if not decimal_found:
                    cleaned.append(".")
                    decimal_found = True
            elif cleaned:
                break
        if not cleaned:
            return None
        return float("".join(cleaned))
    except Exception:
        return None


def iterate_collection(collection):
    if not collection:
        return
    try:
        iterator = collection.ForwardIterator()
        while iterator.MoveNext():
            yield iterator.Current
        return
    except AttributeError:
        pass
    try:
        for item in collection:
            yield item
    except TypeError:
        if collection:
            yield collection
