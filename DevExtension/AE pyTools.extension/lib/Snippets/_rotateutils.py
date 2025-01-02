# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import XYZ, Line, ElementTransformUtils, IndependentTag, FilteredElementCollector, TagOrientation
import math
from collections import defaultdict


# ---------------------- Helper Functions ----------------------

def round_xyz(xyz, precision=6):
    """Rounds the components of an XYZ vector to avoid precision issues."""
    return (round(xyz.X, precision), round(xyz.Y, precision), round(xyz.Z, precision))


def normalize_angle(angle):
    """Normalizes an angle to the range -pi to pi."""
    while angle > math.pi:
        angle -= 2 * math.pi
    while angle < -math.pi:
        angle += 2 * math.pi
    return angle


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
    angle = normalize_angle(target_angle - current_angle)

    # Normalize the angle to be between -pi and pi

    return angle


def rotate_vector_around_z(vector, angle):
    """Rotates a vector around the Z-axis by the specified angle."""
    x = vector.X * math.cos(angle) - vector.Y * math.sin(angle)
    y = vector.X * math.sin(angle) + vector.Y * math.cos(angle)
    return XYZ(x, y, vector.Z)


def collect_data_for_rotation_or_orientation(doc, elements, adjust_tag_position=True):
    """Collects and organizes all necessary data before starting the transaction."""
    element_data = defaultdict(list)

    # If tags should be adjusted, prepare to collect them
    if adjust_tag_position:
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
        hosted_tags = tag_map.get(element.Id, []) if adjust_tag_position else []
        tag_positions = [tag.TagHeadPosition for tag in hosted_tags if tag.TagHeadPosition]
        tag_angles = [tag.RotationAngle for tag in hosted_tags if tag.RotationAngle is not None]

        # Store all collected data
        element_data[orientation_key].append({
            "element": element,
            "element_location": element_location,
            "hosted_tags": hosted_tags,
            "tag_positions": tag_positions,
            "tag_angles": tag_angles,
            "current_orientation": current_orientation
        })

    return element_data


def adjust_tag_locations(grouped_data, angle):
    """Adjusts the positions of tags based on the grouped data structure and specified rotation angle."""
    for data in grouped_data:
        element_location = data["element_location"]
        hosted_tags = data["hosted_tags"]
        original_tag_positions = data["tag_positions"]

        for tag, original_tag_position in zip(hosted_tags, original_tag_positions):
            if not original_tag_position:
                continue

            # Calculate the vector between the original tag position and the element location
            tag_offset_vector = original_tag_position - element_location
            rotated_offset_vector = rotate_vector_around_z(tag_offset_vector, angle)
            new_tag_position = element_location + rotated_offset_vector
            tag.TagHeadPosition = new_tag_position

            # Debugging output
            # print("element_location: {}".format(element_location))
            # print("original_tag_position: {}".format(original_tag_position))
            # print("tag_offset_vector: {}".format(tag_offset_vector))
            # print("rotated_offset_vector: {}".format(rotated_offset_vector))
            # print("new_tag_position: {}".format(new_tag_position))


def adjust_tag_rotations(grouped_data, angle):
    """Adjusts the rotations of tags based on the grouped data structure and specified rotation angle."""

    tolerance = math.radians(5)

    for data in grouped_data:
        hosted_tags = data["hosted_tags"]
        tag_angles = data["tag_angles"]  # Original tag angles

        for tag, original_angle in zip(hosted_tags, tag_angles):
            if original_angle is None:
                continue

            # Calculate the new angle and normalize it
            new_angle = normalize_angle(original_angle + angle)

            # Determine the new orientation based on radians
            if abs(new_angle % (2 * math.pi)) < tolerance or abs(new_angle % (2 * math.pi) - math.pi) < tolerance:
                tag.TagOrientation = TagOrientation.Horizontal
            elif abs(new_angle % (2 * math.pi) - math.pi / 2) < tolerance or abs(new_angle % (2 * math.pi) - 3 * math.pi / 2) < tolerance:
                tag.TagOrientation = TagOrientation.Vertical
            else:
                tag.TagOrientation = TagOrientation.AnyModelDirection
                tag.RotationAngle = new_angle

            # # Debugging output
            # print("Tag: {}".format(tag.Id))
            # print("Original Angle (deg): {:.2f}".format(math.degrees(original_angle)))
            # print("New Angle (deg): {:.2f}".format(new_angle))
            # print("New Orientation: {}".format(tag.TagOrientation))


def rotate_elements_group(doc, grouped_data, angle, adjust_tag_position=True, adjust_tag_rotation=True):
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
        if adjust_tag_position:
            adjust_tag_locations(grouped_data, angle)
            if adjust_tag_rotation:
                adjust_tag_rotations(grouped_data, angle)


def orient_elements_group(doc, grouped_data, target_orientation, adjust_tag_position=True, adjust_tag_rotation=True):
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
        if adjust_tag_position:
            adjust_tag_locations(grouped_data, angle)
            if adjust_tag_rotation:
                adjust_tag_rotations(grouped_data, angle)
