# -*- coding: utf-8 -*-


class PlacementRule(object):
    def __init__(
        self,
        offset_xyz=None,
        rotation_degrees=None,
        placement_basis=None,
        placement_mode=None,
        rotation_basis=None,
        tags=None,
    ):
        self._offset_xyz = tuple(offset_xyz) if offset_xyz is not None else None
        self._rotation_degrees = rotation_degrees
        self._placement_basis = placement_basis
        self._placement_mode = placement_mode
        self._rotation_basis = rotation_basis
        self._tags = list(tags) if tags else []

    def get_offset_xyz(self):
        return self._offset_xyz

    def set_offset_xyz(self, value):
        self._offset_xyz = tuple(value) if value is not None else None

    def get_rotation_degrees(self):
        return self._rotation_degrees

    def set_rotation_degrees(self, value):
        self._rotation_degrees = value

    def get_placement_basis(self):
        return self._placement_basis

    def set_placement_basis(self, value):
        self._placement_basis = value

    def get_placement_mode(self):
        return self._placement_mode

    def set_placement_mode(self, value):
        self._placement_mode = value

    def get_rotation_basis(self):
        return self._rotation_basis

    def set_rotation_basis(self, value):
        self._rotation_basis = value

    def get_tags(self):
        return list(self._tags or [])

    def set_tags(self, value):
        self._tags = list(value) if value else []

    # Helpers
    def update_offset(self, delta_xyz):
        """Add delta to current offset tuple; stores result."""
        if self._offset_xyz is None:
            self._offset_xyz = tuple(delta_xyz) if delta_xyz is not None else None
            return self._offset_xyz
        if delta_xyz is None:
            return self._offset_xyz
        dx, dy, dz = self._offset_xyz
        ddx, ddy, ddz = delta_xyz
        self._offset_xyz = (dx + ddx, dy + ddy, dz + ddz)
        return self._offset_xyz
