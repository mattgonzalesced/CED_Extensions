# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import XYZ, Line, ElementTransformUtils, IndependentTag, FilteredElementCollector
import math
from collections import defaultdict

# ---------------------- Helper Functions ----------------------

def round_xyz(xyz, precision=6):
    """Rounds the components of an XYZ vector to avoid precision issues."""
    return (round(xyz.X, precision), round(xyz.Y, precision), round(xyz.Z, precision))

def calculate_2d_rotation_angle(current_orientation, target_orientation=None, fixed_angle=None):
    """Calculates the signed 2D rotation angle between current and target orientations, ignoring Z.
    If `fixed_angle` is provided, it will return that angle instead of calculating from orientations.
    """
    if fixed_angle is not None:
        return fixed_angle  # Directly use the fixed angle in radians

    # Calculate the angle in the XY plane (ignore the Z component)
    current_angle = math.atan2(current_orientation.Y, current_orientation.X)
    target_angle = math.atan2(target_orientation.Y, target_orientation.X)

    # Calculate the difference between the two angles
    angle = target_angle - current_angle

    # Normalize the angle to be between -pi and pi
    if angle > math.pi:
        angle -= 2 * math.pi
    elif angle < -math.pi:
        angle += 2 * math.pi

    return angle

def rotate_vector_around_z(vector, angle):
    """Rotates a vector around the Z-axis by the specified angle."""
    x = vector.X * math.cos(angle) - vector.Y * math.sin(angle)
    y = vector.X * math.sin(angle) + vector.Y * math.cos(angle)
    return XYZ(x, y, vector.Z)

def collect_data_for_rotation_or_orientation(doc, elements, adjust_tags=True):
    """Collects and organizes all necessary data before starting the transaction."""
    element_data = defaultdict(list)

    # If tags should be adjusted, prepare to collect them
    if adjust_tags:
        tag_collector = FilteredElementCollector(doc, doc.ActiveView.Id).OfClass(IndependentTag)
        element_ids = {element.Id for element in elements}
        tag_iterator = tag_collector.GetElementIdIterator()
        tag_iterator.Reset()
        tag_map = defaultdict(list)

        # Map tags to their elements
        while tag_iterator.MoveNext():
            tag_id = tag_iterator.Current
            tag = doc.GetElement(tag_id)
            tag_referenced_ids = tag.GetTaggedLocalElementIds()

            # Associate tags with elements they reference
            for element_id in tag_referenced_ids:
                if element_id in element_ids:
                    tag_map[element_id].append(tag)

    # Collect element data grouped by orientation or for rotation
    for element in elements:
        if not hasattr(element, 'FacingOrientation'):
            continue

        # Get element orientation and location
        current_orientation = element.FacingOrientation
        orientation_key = round_xyz(current_orientation)
        loc = element.Location
        element_location = loc.Point if hasattr(loc, 'Point') else None

        # Collect data, optionally including tags
        hosted_tags = tag_map.get(element.Id, []) if adjust_tags else []
        tag_positions = [tag.TagHeadPosition for tag in hosted_tags if tag.TagHeadPosition]

        # Store all collected data
        element_data[orientation_key].append({
            "element": element,
            "element_location": element_location,
            "hosted_tags": hosted_tags,
            "tag_positions": tag_positions,
            "current_orientation": current_orientation
        })

    return element_data

def adjust_tags_to_match_rotation(tags, element_location, angle):
    """Adjusts the positions of tags to maintain their relative positioning after host rotation."""
    for tag in tags:
        tag_position = tag.TagHeadPosition
        if not tag_position:
            continue

        # Calculate the vector between the tag and the element location
        tag_offset_vector = tag_position - element_location

        # Rotate the offset vector around the Z-axis by the specified angle
        rotated_offset_vector = rotate_vector_around_z(tag_offset_vector, angle)

        # Calculate the new tag position by adding the rotated offset vector to the element location
        new_tag_position = element_location + rotated_offset_vector

        # Update the tag position
        tag.TagHeadPosition = new_tag_position

def rotate_elements_group(doc, grouped_data, angle, adjust_tags=True):
    """Rotate a group of elements by the specified angle around their local Z-axis."""
    for data in grouped_data:
        element = data["element"]
        loc_point = data["element_location"]
        if not loc_point:
            continue

        # Rotate around the local Z-axis at the element's location
        rotation_axis_line = Line.CreateBound(loc_point, loc_point + XYZ(0, 0, 1))
        ElementTransformUtils.RotateElement(doc, element.Id, rotation_axis_line, angle)

        # Adjust tags after element rotation, if enabled
        if adjust_tags:
            adjust_tags_to_match_rotation(data["hosted_tags"], loc_point, angle)

def orient_elements_group(doc, grouped_data, target_orientation, adjust_tags=True):
    """Orient a group of elements to the target orientation."""
    for data in grouped_data:
        element = data["element"]
        loc_point = data["element_location"]
        current_orientation = data["current_orientation"]

        if not loc_point:
            continue

        # Calculate the 2D rotation angle in the XY plane (ignoring Z)
        angle = calculate_2d_rotation_angle(current_orientation, target_orientation)

        # Rotate around the local Z-axis at the element's location
        rotation_axis_line = Line.CreateBound(loc_point, loc_point + XYZ(0, 0, 1))
        ElementTransformUtils.RotateElement(doc, element.Id, rotation_axis_line, angle)

        # Adjust tags after element orientation, if enabled
        if adjust_tags:
            adjust_tags_to_match_rotation(data["hosted_tags"], loc_point, angle)
