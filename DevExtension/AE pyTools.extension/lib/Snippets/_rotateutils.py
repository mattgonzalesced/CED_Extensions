# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import (XYZ, Line, ElementTransformUtils, IndependentTag, FilteredElementCollector,
                               TagOrientation, LeaderEndCondition, Reference)

import math
from collections import defaultdict
from pyrevit import script
from pyrevit import DB
from pyrevit import HOST_APP

# Initialize logger
logger = script.get_logger()

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
    logger.info(
        "Collecting data for {} elements with adjust_tag_position={}".format(len(elements), adjust_tag_position))
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
            logger.debug("Element {} does not have FacingOrientation.".format(element.Id))
            continue

        # Get element orientation and location
        current_orientation = element.FacingOrientation
        orientation_key = round_xyz(current_orientation)
        loc = element.Location
        element_location = loc.Point if hasattr(loc, 'Point') else None

        # Collect data, optionally including tags
        hosted_tags = tag_map.get(element.Id, []) if adjust_tag_position else []

        tag_positions = []
        tag_angles = []
        leader_data = []

        for tag in hosted_tags:
            # Collect TagHeadPosition if available
            if tag.TagHeadPosition:
                tag_positions.append(tag.TagHeadPosition)

            # Collect RotationAngle if not None
            if tag.RotationAngle is not None:
                tag_angles.append(tag.RotationAngle)

            leader_info = get_leader_info(tag,element)

            # Add leader information if any is present
            if leader_info["leader_elbow"] or leader_info["leader_end"]:
                leader_data.append(leader_info)

        logger.debug("Element {}:"
                     "\n Orientation={},"
                     "\n Location={},"
                     "\n Tags={},"
                     "\n Positions={},"
                     "\n Angles={},"
                     "\n Leaders={}\n".format(
            element.Id, orientation_key, element_location, len(hosted_tags), tag_positions, tag_angles,
            leader_data))

        # Store all collected data
        element_data[orientation_key].append({
            "element": element,
            "element_location": element_location,
            "hosted_tags": hosted_tags,
            "tag_positions": tag_positions,
            "tag_angles": tag_angles,
            "leader_data": leader_data,
            "current_orientation": current_orientation
        })

    return element_data

def get_leader_info(tag, host):
    # Collect leader-related data if applicable, handling exceptions
    leader_info = {"tag": tag, "leader_elbow": None, "leader_end": None}
    refs = tag.GetTaggedReferences()
    for ref in refs:
        logger.debug("reference id: {}".format(ref.ElementId))
        if ref.ElementId == host.Id:
            if HOST_APP.is_newer_than(2022):
                logger.debug("Host App Greater than 2022")
                # Check and collect LeaderEnd if condition is Free
                try:
                    if tag.HasLeader and tag.LeaderEndCondition == LeaderEndCondition.Free:

                        leader_info["leader_end"] = tag.GetLeaderEnd(ref)
                        logger.debug("leader end info: {}, {}".format(tag.HasLeader, tag.LeaderEndCondition))
                        logger.debug("Collected LeaderEnd for Tag {}: {}".format(tag.Id, leader_info["leader_end"]))
                except Exception as e:
                    logger.debug("Tag {} threw an exception when accessing LeaderEnd: {}".format(tag.Id, e))

                    # Check and collect LeaderElbow if HasElbow is True
                try:
                    leader_info["leader_elbow"] = tag.GetLeaderElbow(ref)
                    logger.debug("Collected LeaderElbow for Tag {}: {}".format(tag.Id, leader_info["leader_elbow"]))
                except Exception as e:
                    logger.debug("Tag {} threw an exception when accessing LeaderElbow: {}".format(tag.Id, e))

            else:
                # Check and collect LeaderEnd if condition is Free
                try:
                    if tag.HasLeader and tag.LeaderEndCondition == LeaderEndCondition.Free:
                        leader_info["leader_end"] = tag.LeaderEnd
                        logger.debug("leader end info: {}, {}".format(tag.HasLeader, tag.LeaderEndCondition))
                        logger.debug("Collected LeaderEnd for Tag {}: {}".format(tag.Id, leader_info["leader_end"]))
                except Exception as e:
                    logger.debug("Tag {} threw an exception when accessing LeaderEnd: {}".format(tag.Id, e))

                # Check and collect LeaderElbow if HasElbow is True
                try:
                    if tag.HasElbow:
                        leader_info["leader_elbow"] = tag.LeaderElbow
                except Exception as e:
                    logger.debug("Tag {} threw an exception when accessing LeaderElbow: {}".format(tag.Id, e))


            return leader_info

# Adjust tag locations, including leader positions and elbows
def adjust_tag_locations(grouped_data, angle):
    """Adjusts the positions of tags based on the grouped data structure and specified rotation angle."""
    logger.info("Adjusting tag locations for grouped data with rotation angle {:.4f}".format(angle))
    for data in grouped_data:
        element_location = data["element_location"]
        hosted_tags = data["hosted_tags"]
        original_tag_positions = data["tag_positions"]
        leader_data = data.get("leader_data", [])
        tagged_ref = Reference(data.get("element"))

        for tag, original_tag_position in zip(hosted_tags, original_tag_positions):
            if not original_tag_position:
                continue

            # Calculate the vector between the original tag position and the element location
            tag_offset_vector = original_tag_position - element_location
            rotated_offset_vector = rotate_vector_around_z(tag_offset_vector, angle)
            new_tag_position = element_location + rotated_offset_vector
            tag.TagHeadPosition = new_tag_position

            logger.debug("Tag {}: Offset={}, Rotated Offset={}, New Position={}".format(
                tag.Id, tag_offset_vector, rotated_offset_vector, new_tag_position))

        # Adjust leader positions and elbows, if applicable
        for leader_info in leader_data:
            leader_tag = leader_info["tag"]
            leader_elbow = leader_info["leader_elbow"]
            leader_end = leader_info["leader_end"]

            if leader_elbow:
                leader_elbow_offset = leader_elbow - element_location
                rotated_leader_elbow_offset = rotate_vector_around_z(leader_elbow_offset, angle)
                new_leader_elbow = element_location + rotated_leader_elbow_offset

                if HOST_APP.is_newer_than(2022):
                    leader_tag.SetLeaderElbow(tagged_ref, new_leader_elbow)
                else:
                    leader_tag.LeaderElbow = new_leader_elbow
                logger.debug("Leader Elbow for Tag {}: Offset={}, Rotated Offset={}, New Elbow Position={}".format(
                    leader_tag.Id, leader_elbow_offset, rotated_leader_elbow_offset, new_leader_elbow))

            if leader_end:
                leader_end_offset = leader_end - element_location
                rotated_leader_end_offset = rotate_vector_around_z(leader_end_offset, angle)
                new_leader_end = element_location + rotated_leader_end_offset

                if HOST_APP.is_newer_than(2022):
                    leader_tag.SetLeaderEnd(tagged_ref,new_leader_end)
                else:
                    leader_tag.LeaderEnd = new_leader_end
                logger.debug("Leader End for Tag {}: Offset={}, Rotated Offset={}, New End Position={}".format(
                    leader_tag.Id, leader_end_offset, rotated_leader_end_offset, new_leader_end))


# Adjust tag rotations
def adjust_tag_rotations(grouped_data, angle, keep_model_orientation=False):
    """Adjusts the rotations of tags based on the grouped data structure and specified rotation angle."""
    logger.info("Adjusting tag rotations for grouped data with rotation angle {:.4f}".format(angle))
    tolerance = math.radians(5)

    for data in grouped_data:
        hosted_tags = data["hosted_tags"]
        tag_angles = data["tag_angles"]

        for tag, original_angle in zip(hosted_tags, tag_angles):
            if original_angle is None:
                logger.debug("Original Angle is None: {}".format(original_angle))
                continue

            new_angle = normalize_angle(original_angle + angle)

            if not keep_model_orientation:
                apply_orientation_rules(tag, new_angle, tolerance)
            else:
                tag.TagOrientation = TagOrientation.AnyModelDirection
                tag.RotationAngle = new_angle

            logger.debug("Tag {}: Original Angle={}, New Angle={}, Orientation={}".format(
                tag.Id, math.degrees(original_angle), math.degrees(new_angle), tag.TagOrientation))


def apply_orientation_rules(tag, new_angle, tolerance=math.radians(5)):
    """Applies horizontal/vertical/model orientation rules to a tag based on its new angle."""
    logger.debug("Adjusting Tag of Category {}".format(tag.Category))

    if tag.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_KeynoteTags):
        logger.debug("Tag is Keynote. Make Horizontal")
        tag.TagOrientation = TagOrientation.Horizontal
    elif abs(new_angle % (2 * math.pi)) < tolerance or abs(new_angle % (2 * math.pi) - math.pi) < tolerance:
        tag.TagOrientation = TagOrientation.Horizontal
    elif abs(new_angle % (2 * math.pi) - math.pi / 2) < tolerance or abs(new_angle % (2 * math.pi) - 3 * math.pi / 2) < tolerance:
        tag.TagOrientation = TagOrientation.Vertical
    else:
        tag.TagOrientation = TagOrientation.AnyModelDirection
        tag.RotationAngle = new_angle



# Rotate elements group with leader data handling
def rotate_elements_group(doc, grouped_data, angle, adjust_tag_position=True, adjust_tag_rotation=True, keep_model_orientation=False):
    for data in grouped_data:
        element = data["element"]
        loc_point = data["element_location"]
        if not loc_point:
            continue

        rotation_axis_line = Line.CreateBound(loc_point, loc_point + XYZ(0, 0, 1))
        ElementTransformUtils.RotateElement(doc, element.Id, rotation_axis_line, angle)

        if adjust_tag_position:
            adjust_tag_locations(grouped_data, angle)
            if adjust_tag_rotation:
                adjust_tag_rotations(grouped_data, angle, keep_model_orientation)

def orient_elements_group(doc, grouped_data, target_orientation, adjust_tag_position=True, adjust_tag_rotation=True, keep_model_orientation=False):
    for data in grouped_data:
        element = data["element"]
        loc_point = data["element_location"]
        current_orientation = data["current_orientation"]
        if not loc_point:
            continue

        angle = calculate_2d_rotation_angle(current_orientation, target_orientation)
        rotation_axis_line = Line.CreateBound(loc_point, loc_point + XYZ(0, 0, 1))
        ElementTransformUtils.RotateElement(doc, element.Id, rotation_axis_line, angle)

        if adjust_tag_position:
            adjust_tag_locations(grouped_data, angle)
            if adjust_tag_rotation:
                adjust_tag_rotations(grouped_data, angle, keep_model_orientation)