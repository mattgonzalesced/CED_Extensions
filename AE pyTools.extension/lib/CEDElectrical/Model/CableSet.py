
class CableSet(object):
    def __init__(self, cableset_type, wire_size, sets=1):
        self.type = cableset_type            # CableSetType instance
        self.wire_hot_qty = wire_hot_qty
        self.wire_hot_size = wire_hot_size
        self.wire_neutral_qty = wire_neutral_qty
        self.wire_neutral_size = wire_neutral_size
        self.wire_ground_qty = wire_ground_qty
        self.wire_ground_size = wire_ground_size
        self.wire_isoground_qty= wire_isoground_qty
        self.wire_isoground_size = wire_isoground_size
        self.sets = sets                     # number of parallel sets
        self._ampacity = None
        self._voltage_drop = None

    def get_ampacity(self):
        table = WIRE_AMPACITY_TABLE.get(self.type.material, {}).get(self.type.temperature, [])
        for size, amps in table:
            if size == self.wire_size:
                self._ampacity = amps * self.sets
                return self._ampacity
        return None

    def calculate_voltage_drop(self, current, length, voltage, pf=0.9):
        zdata = WIRE_IMPEDANCE_TABLE.get(self.wire_size)
        if not zdata:
            return None

        R = zdata['R'].get(self.type.material, {}).get(self.type.conduit_type)
        X = zdata['XL'].get(self.type.conduit_type)
        if R is None or X is None:
            return None

        R /= self.sets
        X /= self.sets
        sin_phi = (1 - pf**2)**0.5

        phase = 3  # assume 3-phase; update as needed
        if phase == 3:
            drop = (1.732 * current * (R * pf + X * sin_phi) * length) / 1000.0
        else:
            drop = (2 * current * (R * pf + X * sin_phi) * length) / 1000.0

        self._voltage_drop = (drop / voltage) * 100
        return self._voltage_drop
