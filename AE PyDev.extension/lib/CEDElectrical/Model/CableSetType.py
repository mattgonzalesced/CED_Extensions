class CableSetType(object):
    def __init__(self, material, insulation, temperature, conduit_type):
        self.material = material              # 'Copper' or 'Aluminum'
        self.insulation = insulation          # 'THWN', 'XHHW-2', etc.
        self.temperature = int(temperature)   # 60, 75, 90
        self.conduit_type = conduit_type      # 'PVC-40', 'EMT', etc.

    def __repr__(self):
        return "{} {} @{}°C in {}".format(
            self.material, self.insulation, self.temperature, self.conduit_type
        )
