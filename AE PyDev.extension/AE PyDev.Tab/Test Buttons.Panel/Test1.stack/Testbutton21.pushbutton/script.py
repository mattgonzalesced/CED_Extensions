# -*- coding: utf-8 -*-

from pyrevit import DB, script, forms, revit, output
from pyrevit.revit import query
import clr
from Snippets import _elecutils as eu
from Autodesk.Revit.DB.Electrical import *

app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc

console = script.get_output()
logger = script.get_logger()

WIRE_AMPACITY_TABLE = {
    'Copper': {
        60: [
            ('14', 15), ('12', 20), ('10', 30), ('8', 40), ('6', 55),
            ('4', 70), ('3', 85), ('2', 95), ('1', 110), ('1/0', 125),
            ('2/0', 145), ('3/0', 165), ('4/0', 195), ('250', 215), ('300', 240),
            ('350', 260), ('400', 280), ('500', 320), ('600', 355), ('700', 385),
            ('750', 400), ('800', 410), ('900', 435), ('1000', 455), ('1250', 495),
            ('1500', 520), ('1750', 545), ('2000', 560)
        ],
        75: [
            ('14', 20), ('12', 25), ('10', 35), ('8', 50), ('6', 65),
            ('4', 85), ('3', 100), ('2', 115), ('1', 130), ('1/0', 150),
            ('2/0', 175), ('3/0', 200), ('4/0', 230), ('250', 255), ('300', 285),
            ('350', 310), ('400', 335), ('500', 380), ('600', 420), ('700', 460),
            ('750', 475), ('800', 490), ('900', 520), ('1000', 545), ('1250', 590),
            ('1500', 625), ('1750', 650), ('2000', 665)
        ],
        90: [
            ('14', 25), ('12', 30), ('10', 40), ('8', 55), ('6', 75),
            ('4', 95), ('3', 110), ('2', 130), ('1', 150), ('1/0', 170),
            ('2/0', 195), ('3/0', 225), ('4/0', 260), ('250', 290), ('300', 320),
            ('350', 350), ('400', 380), ('500', 430), ('600', 475), ('700', 520),
            ('750', 535), ('800', 555), ('900', 585), ('1000', 615), ('1250', 665),
            ('1500', 705), ('1750', 735), ('2000', 750)
        ]
    },
    'Aluminium': {
        60: [
            ('12', 15), ('10', 25), ('8', 35), ('6', 40), ('4', 55),
            ('3', 65), ('2', 75), ('1', 90), ('1/0', 100), ('2/0', 120),
            ('3/0', 135), ('4/0', 155), ('250', 170), ('300', 190),
            ('350', 210), ('400', 225), ('500', 260), ('600', 285), ('700', 310),
            ('750', 320), ('800', 330), ('900', 355), ('1000', 375), ('1250', 405),
            ('1500', 435), ('1750', 455), ('2000', 470)
        ],
        75: [
            ('12', 20), ('10', 30), ('8', 40), ('6', 50), ('4', 65),
            ('3', 75), ('2', 90), ('1', 100), ('1/0', 120), ('2/0', 135),
            ('3/0', 155), ('4/0', 180), ('250', 205), ('300', 230),
            ('350', 250), ('400', 270), ('500', 310), ('600', 340), ('700', 375),
            ('750', 385), ('800', 395), ('900', 425), ('1000', 445), ('1250', 485),
            ('1500', 520), ('1750', 545), ('2000', 560)
        ],
        90: [
            ('12', 25), ('10', 35), ('8', 45), ('6', 60), ('4', 75),
            ('3', 85), ('2', 100), ('1', 115), ('1/0', 135), ('2/0', 150),
            ('3/0', 175), ('4/0', 205), ('250', 230), ('300', 255),
            ('350', 280), ('400', 305), ('500', 350), ('600', 385), ('700', 420),
            ('750', 435), ('800', 450), ('900', 480), ('1000', 500), ('1250', 545),
            ('1500', 585), ('1750', 615), ('2000', 630)
        ]
    }
}

EGC_TABLE = {
    'Copper': [
        (15, '14'), (20, '12'), (30, '10'), (40, '10'), (60, '10'),
        (100, '8'), (200, '6'), (300, '4'), (400, '3'), (500, '2'),
        (600, '1'), (800, '1/0'), (1000, '2/0'), (1200, '3/0'),
        (1600, '4/0'), (2000, '250'), (2500, '350'), (3000, '400'),
        (4000, '500'), (5000, '700'), (6000, '800')
    ],
    'Aluminium': [
        (15, '12'), (20, '10'), (30, '8'), (40, '8'), (60, '8'),
        (100, '6'), (200, '4'), (300, '2'), (400, '1'), (500, '1/0'),
        (600, '2/0'), (800, '3/0'), (1000, '4/0'), (1200, '250'),
        (1600, '350'), (2000, '400'), (2500, '600'), (3000, '600'),
        (4000, '800'), (5000, '1200'), (6000, '1200')
    ]
}

BREAKER_FRAME_SWITCH_TABLE = {
    15: {'frame': 30, 'switch': 30},
    20: {'frame': 30, 'switch': 30},
    25: {'frame': 30, 'switch': 30},
    30: {'frame': 30, 'switch': 30},
    35: {'frame': 60, 'switch': 60},
    40: {'frame': 60, 'switch': 60},
    45: {'frame': 60, 'switch': 60},
    50: {'frame': 60, 'switch': 60},
    60: {'frame': 60, 'switch': 60},
    70: {'frame': 100, 'switch': 100},
    80: {'frame': 100, 'switch': 100},
    90: {'frame': 100, 'switch': 100},
    100: {'frame': 100, 'switch': 100},
    125: {'frame': 200, 'switch': 200},
    150: {'frame': 200, 'switch': 200},
    175: {'frame': 200, 'switch': 200},
    200: {'frame': 200, 'switch': 200},
    225: {'frame': 225, 'switch': 400},
    250: {'frame': 250, 'switch': 400},
    300: {'frame': 400, 'switch': 400},
    350: {'frame': 400, 'switch': 400},
    400: {'frame': 400, 'switch': 400},
    450: {'frame': 600, 'switch': 600},
    500: {'frame': 600, 'switch': 600},
    600: {'frame': 600, 'switch': 600},
    700: {'frame': 800, 'switch': 800},
    800: {'frame': 800, 'switch': 800},
    1000: {'frame': 1000, 'switch': 1000},
    1200: {'frame': 1200, 'switch': 1200},
    1600: {'frame': 1600, 'switch': 1600},
    2000: {'frame': 2000, 'switch': 2000},
    2500: {'frame': 2500, 'switch': 2500},
    3000: {'frame': 3000, 'switch': 3000},
    4000: {'frame': 4000, 'switch': 4000},
    5000: {'frame': 5000, 'switch': 5000},
    6000: {'frame': 6000, 'switch': 6000}
}

WIRE_IMPEDANCE_TABLE = {
    '12': {
        'XL': {'PVC': 0.054, 'Steel': 0.068},
        'R': {
            'Copper': {'PVC': 2.00, 'Steel': 2.00},
            'Aluminum': {'PVC': 3.20, 'Steel': 3.20}
        }
    },
    '10': {
        'XL': {'PVC': 0.050, 'Steel': 0.063},
        'R': {
            'Copper': {'PVC': 1.20, 'Steel': 1.20},
            'Aluminum': {'PVC': 2.00, 'Steel': 2.00}
        }
    },
    '8': {
        'XL': {'PVC': 0.052, 'Steel': 0.065},
        'R': {
            'Copper': {'PVC': 0.78, 'Steel': 0.78},
            'Aluminum': {'PVC': 1.30, 'Steel': 1.30}
        }
    },
    '6': {
        'XL': {'PVC': 0.051, 'Steel': 0.064},
        'R': {
            'Copper': {'PVC': 0.49, 'Steel': 0.49},
            'Aluminum': {'PVC': 0.81, 'Steel': 0.81}
        }
    },
    '4': {
        'XL': {'PVC': 0.048, 'Steel': 0.060},
        'R': {
            'Copper': {'PVC': 0.31, 'Steel': 0.31},
            'Aluminum': {'PVC': 0.51, 'Steel': 0.51}
        }
    },
    '3': {
        'XL': {'PVC': 0.047, 'Steel': 0.059},
        'R': {
            'Copper': {'PVC': 0.25, 'Steel': 0.25},
            'Aluminum': {'PVC': 0.40, 'Steel': 0.40}
        }
    },
    '2': {
        'XL': {'PVC': 0.045, 'Steel': 0.057},
        'R': {
            'Copper': {'PVC': 0.19, 'Steel': 0.20},
            'Aluminum': {'PVC': 0.32, 'Steel': 0.32}
        }
    },
    '1': {
        'XL': {'PVC': 0.046, 'Steel': 0.057},
        'R': {
            'Copper': {'PVC': 0.15, 'Steel': 0.16},
            'Aluminum': {'PVC': 0.25, 'Steel': 0.25}
        }
    },
    '1/0': {
        'XL': {'PVC': 0.044, 'Steel': 0.055},
        'R': {
            'Copper': {'PVC': 0.12, 'Steel': 0.12},
            'Aluminum': {'PVC': 0.20, 'Steel': 0.20}
        }
    },
    '2/0': {
        'XL': {'PVC': 0.043, 'Steel': 0.054},
        'R': {
            'Copper': {'PVC': 0.10, 'Steel': 0.10},
            'Aluminum': {'PVC': 0.16, 'Steel': 0.16}
        }
    },
    '3/0': {
        'XL': {'PVC': 0.042, 'Steel': 0.052},
        'R': {
            'Copper': {'PVC': 0.077, 'Steel': 0.079},
            'Aluminum': {'PVC': 0.13, 'Steel': 0.13}
        }
    },
    '4/0': {
        'XL': {'PVC': 0.041, 'Steel': 0.051},
        'R': {
            'Copper': {'PVC': 0.062, 'Steel': 0.063},
            'Aluminum': {'PVC': 0.10, 'Steel': 0.10}
        }
    },
    '250': {
        'XL': {'PVC': 0.041, 'Steel': 0.052},
        'R': {
            'Copper': {'PVC': 0.052, 'Steel': 0.054},
            'Aluminum': {'PVC': 0.085, 'Steel': 0.086}
        }
    },
    '300': {
        'XL': {'PVC': 0.041, 'Steel': 0.051},
        'R': {
            'Copper': {'PVC': 0.044, 'Steel': 0.045},
            'Aluminum': {'PVC': 0.071, 'Steel': 0.072}
        }
    },
    '350': {
        'XL': {'PVC': 0.040, 'Steel': 0.050},
        'R': {
            'Copper': {'PVC': 0.038, 'Steel': 0.039},
            'Aluminum': {'PVC': 0.061, 'Steel': 0.063}
        }
    },
    '400': {
        'XL': {'PVC': 0.040, 'Steel': 0.049},
        'R': {
            'Copper': {'PVC': 0.033, 'Steel': 0.035},
            'Aluminum': {'PVC': 0.054, 'Steel': 0.055}
        }
    },
    '500': {
        'XL': {'PVC': 0.039, 'Steel': 0.048},
        'R': {
            'Copper': {'PVC': 0.027, 'Steel': 0.029},
            'Aluminum': {'PVC': 0.043, 'Steel': 0.045}
        }
    },
    '600': {
        'XL': {'PVC': 0.039, 'Steel': 0.048},
        'R': {
            'Copper': {'PVC': 0.023, 'Steel': 0.025},
            'Aluminum': {'PVC': 0.036, 'Steel': 0.038}
        }
    },
    '750': {
        'XL': {'PVC': 0.038, 'Steel': 0.048},
        'R': {
            'Copper': {'PVC': 0.019, 'Steel': 0.021},
            'Aluminum': {'PVC': 0.029, 'Steel': 0.031}
        }
    },
    '1000': {
        'XL': {'PVC': 0.037, 'Steel': 0.046},
        'R': {
            'Copper': {'PVC': 0.015, 'Steel': 0.018},
            'Aluminum': {'PVC': 0.023, 'Steel': 0.025}
        }
    }
}

CONDUCTOR_AREA_TABLE = {
    '14': {'cmil': 4110, 'area': {'THHN': 0.0097, 'THWN': 0.0097, 'THWN-2': 0.0097, 'RHW': 0.0293, 'XHHW': 0.0139, 'XHHW-2': 0.0139, 'XHH': 0.0139}},
    '12': {'cmil': 6530, 'area': {'THHN': 0.0133, 'THWN': 0.0133, 'THWN-2': 0.0133, 'RHW': 0.0353, 'XHHW': 0.0181, 'XHHW-2': 0.0181, 'XHH': 0.0181}},
    '10': {'cmil': 10380, 'area': {'THHN': 0.0211, 'THWN': 0.0211, 'THWN-2': 0.0211, 'RHW': 0.0437, 'XHHW': 0.0243, 'XHHW-2': 0.0243, 'XHH': 0.0243}},
    '8': {'cmil': 16510, 'area': {'THHN': 0.0366, 'THWN': 0.0366, 'THWN-2': 0.0366, 'RHW': 0.0835, 'XHHW': 0.0437, 'XHHW-2': 0.0437, 'XHH': 0.0437}},
    '6': {'cmil': 26240, 'area': {'THHN': 0.0507, 'THWN': 0.0507, 'THWN-2': 0.0507, 'RHW': 0.1041, 'XHHW': 0.0590, 'XHHW-2': 0.0590, 'XHH': 0.0590}},
    '4': {'cmil': 41740, 'area': {'THHN': 0.0824, 'THWN': 0.0824, 'THWN-2': 0.0824, 'RHW': 0.1333, 'XHHW': 0.0814, 'XHHW-2': 0.0814, 'XHH': 0.0814}},
    '3': {'cmil': 52620, 'area': {'THHN': 0.0973, 'THWN': 0.0973, 'THWN-2': 0.0973, 'RHW': 0.1521, 'XHHW': 0.0962, 'XHHW-2': 0.0962, 'XHH': 0.0962}},
    '2': {'cmil': 66360, 'area': {'THHN': 0.1158, 'THWN': 0.1158, 'THWN-2': 0.1158, 'RHW': 0.1750, 'XHHW': 0.1146, 'XHHW-2': 0.1146, 'XHH': 0.1146}},
    '1': {'cmil': 83690, 'area': {'THHN': 0.1562, 'THWN': 0.1562, 'THWN-2': 0.1562, 'RHW': 0.2660, 'XHHW': 0.1534, 'XHHW-2': 0.1534, 'XHH': 0.1534}},
    '1/0': {'cmil': 105600, 'area': {'THHN': 0.1855, 'THWN': 0.1855, 'THWN-2': 0.1855, 'RHW': 0.3039, 'XHHW': 0.1825, 'XHHW-2': 0.1825, 'XHH': 0.1825}},
    '2/0': {'cmil': 133100, 'area': {'THHN': 0.2223, 'THWN': 0.2223, 'THWN-2': 0.2223, 'RHW': 0.3505, 'XHHW': 0.2190, 'XHHW-2': 0.2190, 'XHH': 0.2190}},
    '3/0': {'cmil': 167800, 'area': {'THHN': 0.2679, 'THWN': 0.2679, 'THWN-2': 0.2679, 'RHW': 0.4072, 'XHHW': 0.2642, 'XHHW-2': 0.2642, 'XHH': 0.2642}},
    '4/0': {'cmil': 211600, 'area': {'THHN': 0.3237, 'THWN': 0.3237, 'THWN-2': 0.3237, 'RHW': 0.4754, 'XHHW': 0.3197, 'XHHW-2': 0.3197, 'XHH': 0.3197}},
    '250': {'cmil': 250000, 'area': {'THHN': 0.397, 'THWN': 0.397, 'THWN-2': 0.397, 'RHW': 0.6291, 'XHHW': 0.3904, 'XHHW-2': 0.3904, 'XHH': 0.3904}},
    '300': {'cmil': 300000, 'area': {'THHN': 0.4608, 'THWN': 0.4608, 'THWN-2': 0.4608, 'RHW': 0.7088, 'XHHW': 0.4536, 'XHHW-2': 0.4536, 'XHH': 0.4536}},
    '350': {'cmil': 350000, 'area': {'THHN': 0.5242, 'THWN': 0.5242, 'THWN-2': 0.5242, 'RHW': 0.787, 'XHHW': 0.5166, 'XHHW-2': 0.5166, 'XHH': 0.5166}},
    '400': {'cmil': 400000, 'area': {'THHN': 0.5863, 'THWN': 0.5863, 'THWN-2': 0.5863, 'RHW': 0.8626, 'XHHW': 0.5782, 'XHHW-2': 0.5782, 'XHH': 0.5782}},
    '500': {'cmil': 500000, 'area': {'THHN': 0.7073, 'THWN': 0.7073, 'THWN-2': 0.7073, 'RHW': 1.0082, 'XHHW': 0.6984, 'XHHW-2': 0.6984, 'XHH': 0.6984}},
    '600': {'cmil': 600000, 'area': {'THHN': 0.8676, 'THWN': 0.8676, 'THWN-2': 0.8676, 'RHW': 1.2135, 'XHHW': 0.8709, 'XHHW-2': 0.8709, 'XHH': 0.8709}},
    '700': {'cmil': 700000, 'area': {'THHN': 0.9887, 'THWN': 0.9887, 'THWN-2': 0.9887, 'RHW': 1.3561, 'XHHW': 0.9923, 'XHHW-2': 0.9923, 'XHH': 0.9923}},
    '750': {'cmil': 750000, 'area': {'THHN': 1.0496, 'THWN': 1.0496, 'THWN-2': 1.0496, 'RHW': 1.4272, 'XHHW': 1.0532, 'XHHW-2': 1.0532, 'XHH': 1.0532}},
    '800': {'cmil': 800000, 'area': {'THHN': 1.1085, 'THWN': 1.1085, 'THWN-2': 1.1085, 'RHW': 1.4957, 'XHHW': 1.1122, 'XHHW-2': 1.1122, 'XHH': 1.1122}},
    '900': {'cmil': 900000, 'area': {'THHN': 1.2311, 'THWN': 1.2311, 'THWN-2': 1.2311, 'RHW': 1.6377, 'XHHW': 1.2351, 'XHHW-2': 1.2351, 'XHH': 1.2351}},
    '1000': {'cmil': 1000000, 'area': {'THHN': 1.3478, 'THWN': 1.3478, 'THWN-2': 1.3478, 'RHW': 1.7719, 'XHHW': 1.3519, 'XHHW-2': 1.3519, 'XHH': 1.3519}}
}

PART_TYPE_MAP = {
    14: "Panelboard",
    15: "Transformer",
    16: "Switchboard",
    17: "Other Panel",
    18: "Equipment Switch"
}

CONDUIT_AREA_TABLE = {
    'non_magnetic': {
        'PVC-40': {
            '1/2"': 0.2850, '3/4"': 0.5080, '1"': 0.8320, '1-1/4"': 1.4530, '1-1/2"': 1.9860,
            '2"': 3.2910, '2-1/2"': 4.6950, '3"': 7.2680, '3-1/2"': 9.7370, '4"': 12.5540,
            '5"': 19.7610, '6"': 28.5670
        },
        'PVC-80': {
            '1/2"': 0.2170, '3/4"': 0.4090, '1"': 0.6880, '1-1/4"': 1.2370, '1-1/2"': 1.7110,
            '2"': 2.8740, '2-1/2"': 4.1190, '3"': 6.4420, '3-1/2"': 8.6880, '4"': 11.2580,
            '5"': 17.8550, '6"': 25.5980
        },
        'ENT': {
            '1/2"': 0.2850, '3/4"': 0.5080, '1"': 0.8320, '1-1/4"': 1.4530, '1-1/2"': 1.9860,
            '2"': 3.2910
        },
        'LFNC-A': {
            '1/2"': 0.3120, '3/4"': 0.5350, '1"': 0.8540, '1-1/4"': 1.5020, '1-1/2"': 2.0180,
            '2"': 3.3430
        },
        'LFNC-B': {
            '1/2"': 0.3140, '3/4"': 0.5410, '1"': 0.8730, '1-1/4"': 1.5280, '1-1/2"': 1.9810,
            '2"': 3.2460
        }
    },
    'magnetic': {
        'EMT': {
            '1/2"': 0.3040, '3/4"': 0.5330, '1"': 0.8640, '1-1/4"': 1.4960, '1-1/2"': 2.0360,
            '2"': 3.3560, '2-1/2"': 5.8580, '3"': 8.8460, '3-1/2"': 11.5450, '4"': 14.7530
        },
        'RMC': {
            '1/2"': 0.3140, '3/4"': 0.5790, '1"': 0.8870, '1-1/4"': 1.5260, '1-1/2"': 2.0710,
            '2"': 3.4080, '2-1/2"': 4.8660, '3"': 7.4990, '3-1/2"': 10.0100, '4"': 12.8820,
            '5"': 20.2120, '6"': 29.1580
        },
        'FMC': {
            '1/2"': 0.3170, '3/4"': 0.5330, '1"': 0.8170, '1-1/4"': 1.2770, '1-1/2"': 1.8580,
            '2"': 3.2690, '2-1/2"': 4.9090, '3"': 7.0690, '3-1/2"': 9.6210, '4"': 12.5660
        },
        'IMC': {
            '1/2"': 0.3420, '3/4"': 0.5860, '1"': 0.9590, '1-1/4"': 1.6470, '1-1/2"': 2.2250,
            '2"': 3.6300, '2-1/2"': 5.1350, '3"': 7.9220, '3-1/2"': 10.5840, '4"': 13.6310
        },
        'LFMC': {
            '1/2"': 0.3140, '3/4"': 0.5410, '1"': 0.8730, '1-1/4"': 1.5280, '1-1/2"': 1.9810,
            '2"': 3.2460, '2-1/2"': 4.8810, '3"': 7.4750, '3-1/2"': 9.7310, '4"': 12.6920
        }
    }
}


class CircuitSettings(object):
    def __init__(self):
        # User-adjustable settings (can be loaded from file or UI later)
        self.min_wire_size = '12'
        self.max_wire_size = '600'
        self.min_breaker_size = 20
        self.max_parallel_size = '500'  # largest wire allowed before parallel
        self.excluded_wire_sizes = []

    def to_dict(self):
        return {
            'min_wire_size': self.min_wire_size,
            'max_wire_size': self.max_wire_size,
            'min_breaker_size': self.min_breaker_size,
            'max_parallel_size': self.max_parallel_size
        }

    def load_from_dict(self, data):
        self.min_wire_size = data.get('min_wire_size', self.min_wire_size)
        self.max_wire_size = data.get('max_wire_size', self.max_wire_size)
        self.min_breaker_size = data.get('min_breaker_size', self.min_breaker_size)
        self.max_parallel_size = data.get('max_parallel_size', self.max_parallel_size)


class CircuitBranch(object):
    def __init__(self, circuit, settings=None):
        self.circuit = circuit
        self.settings = settings if settings else CircuitSettings()
        self.circuit_id = circuit.Id.IntegerValue
        self.name = "{}-{}".format(circuit.BaseEquipment.Name, circuit.CircuitNumber)

        self._wire_info = None  # Lazy-loaded wire info dictionary

        # User overrides (None = no override)
        self._breaker_override = None
        self._hot_wire_override = None
        self._ground_wire_override = None
        self._max_single_wire_size = '500'  # Max size before parallel sets

        self._initial_wire_size = None
        self._initial_wire_ampacity = None
        self._initial_voltage_drop = None

        # Calculated values (set by calculation methods)
        self._calculated_breaker = self.rating
        self._calculated_hot_wire = None
        self._calculated_hot_sets = None
        self._calculated_hot_ampacity = None
        self._calculated_ground_wire = None
        self._calculated_voltage_drop = None  # % final voltage drop
        self._calculated_conduit_size = None  # e.g. '1"', '1-1/4"', etc.
        self._calculated_conduit_fill = None  # total conductor area used (in²)

    # ----------- Classification -----------

    @property
    def is_power_circuit(self):
        return self.circuit.SystemType == ElectricalSystemType.PowerCircuit

    @property
    def is_feeder(self):
        try:
            # Get elements this circuit supplies
            elements = list(self.circuit.Elements)
            for el in elements:
                if isinstance(el, DB.FamilyInstance):
                    family = el.Symbol.Family
                    part_type = family.get_Parameter(DB.BuiltInParameter.FAMILY_CONTENT_PART_TYPE)
                    if part_type and part_type.StorageType == DB.StorageType.Integer:
                        part_value = part_type.AsInteger()
                        if part_value in [14, 15, 16, 17]:
                            return True
        except Exception as e:
            logger.debug("Error in is_feeder: {}".format(str(e)))
        return False

    @property
    def is_spare(self):
        return self.circuit.CircuitType == CircuitType.Spare

    @property
    def is_space(self):
        return self.circuit.CircuitType == CircuitType.Space

    # ----------- Wire Info Dictionary -----------

    @property
    def wire_info(self):
        if not self.is_power_circuit:
            return {}
        if self._wire_info is None:
            try:
                wt = self.circuit.WireType
                self._wire_info = {
                    'material': DB.Element.Name.__get__(wt.WireMaterial),
                    'temperature': DB.Element.Name.__get__(wt.TemperatureRating),
                    'conduit': WireConduitType.Name.__get__(wt.Conduit),
                    'insulation': DB.Element.Name.__get__(wt.Insulation)
                }
            except Exception:
                self._wire_info = {}
        return self._wire_info

    # ----------- Circuit Properties -----------

    @property
    def rating(self):
        try:
            if self.is_power_circuit and not self.is_space:
                return self.circuit.Rating
        except:
            return None

    @property
    def length(self):
        try:
            if self.is_power_circuit and not self.is_spare and not self.is_space:
                return self.circuit.Length
        except:
            return None

    @property
    def voltage(self):
        try:
            return ElectricalSystem.Voltage.__get__(self.circuit)
        except:
            return None

    @property
    def apparent_power(self):
        try:
            return ElectricalSystem.ApparentLoad.__get__(self.circuit)
        except:
            return None

    @property
    def apparent_current(self):
        try:
            return ElectricalSystem.ApparentCurrent.__get__(self.circuit)
        except:
            return None

    @property
    def poles(self):
        try:
            return ElectricalSystem.PolesNumber.__get__(self.circuit)
        except:
            return None

    @property
    def power_factor(self):
        try:
            return ElectricalSystem.PowerFactor.__get__(self.circuit)
        except:
            return None

    # ----------- Override Setters -----------

    def set_breaker_override(self, value):
        self._breaker_override = value

    def set_hot_wire_override(self, wire_size):
        self._hot_wire_override = wire_size

    def set_ground_wire_override(self, wire_size):
        self._ground_wire_override = wire_size

    def set_max_single_wire_size(self, wire_size):
        self._max_single_wire_size = wire_size

    # ----------- Public Access Properties -----------

    @property
    def breaker_rating(self):
        return self._breaker_override if self._breaker_override is not None else self._calculated_breaker

    @property
    def hot_wire_size(self):
        return self._hot_wire_override if self._hot_wire_override is not None else self._calculated_hot_wire

    @property
    def ground_wire_size(self):
        return self._ground_wire_override if self._ground_wire_override is not None else self._calculated_ground_wire

    @property
    def number_of_sets(self):
        return self._calculated_hot_sets

    @property
    def circuit_base_ampacity(self):
        return self._calculated_hot_ampacity

    @property
    def max_voltage_drop(self):
        if not self.is_power_circuit:
            return None
        return 2.0 if self.is_feeder else 3.0

    # ----------- Calculations -----------

    def calculate_breaker_size(self):
        try:
            amps = self.apparent_current
            if amps is not None:
                amps *= 1.25
                if amps < self.settings.min_breaker_size:
                    amps = self.settings.min_breaker_size

                for b in sorted(BREAKER_FRAME_SWITCH_TABLE.keys()):
                    if b >= amps:
                        self._calculated_breaker = b
                        break
        except:
            self._calculated_breaker = None

    def calculate_hot_wire_size(self):
        if self.rating:
            self._breaker_override = self.rating  # Respect Revit breaker setting

        try:
            breaker_amps = self.breaker_rating
            if breaker_amps is None:
                return

            temp = int(self.wire_info.get('temperature', '75').replace('C', '').replace('°', ''))
            material = self.wire_info.get('material', 'Copper')
            wire_set = WIRE_AMPACITY_TABLE.get(material, {}).get(temp, [])

            min_size = self.settings.min_wire_size
            max_size = self.settings.max_parallel_size
            sets = 1

            # Find index of minimum allowed wire size
            start_index = 0

            # Determine initial base wire size from ampacity table (1 set)
            for wire, ampacity in wire_set[start_index:]:
                if ampacity >= breaker_amps:
                    self._initial_wire_size = wire
                    self._initial_wire_ampacity = ampacity
                    self._initial_voltage_drop = self.calculate_voltage_drop(wire, 1)
                    break

            for i, (size, _) in enumerate(wire_set):
                if size == min_size:
                    start_index = i
                    break

            sets = 1

            while sets < 10:
                for wire, ampacity in wire_set[start_index:]:
                    if ampacity * sets >= breaker_amps:
                        vd_percent = self.calculate_voltage_drop(wire, sets)
                        if vd_percent is not None and self.max_voltage_drop is not None:
                            if vd_percent > self.max_voltage_drop:
                                continue  # Try next bigger wire

                        # ✅ Passed both checks — accept wire
                        self._calculated_hot_wire = wire
                        self._calculated_hot_sets = sets
                        self._calculated_hot_ampacity = ampacity * sets
                        return

                    # If you reached max size for current set count — stop and go to more sets
                    if wire == max_size:
                        break  # exit wire loop to increment sets

                # 🔁 Increase sets and restart from start_index
                sets += 1

            # Fallback if no suitable wire was found
            self._calculated_hot_wire = None
            self._calculated_hot_sets = None
            self._calculated_hot_ampacity = None

        except Exception as e:
            self._calculated_hot_wire = None
            self._calculated_hot_sets = None
            self._calculated_hot_ampacity = None

    def calculate_ground_wire_size(self):
        try:
            amps = self.breaker_rating
            if amps is None:
                return

            material = self.wire_info.get('material', 'Copper')
            min_size = self.settings.min_wire_size
            egc_list = EGC_TABLE.get(material, [])

            # Find index of first entry that matches or exceeds min wire size
            start_index = 0
            for i, (_, wire_size) in enumerate(egc_list):
                if wire_size == min_size:
                    start_index = i
                    break

            # Filter from the min wire size forward
            for amp_limit, wire_size in egc_list[start_index:]:
                if amps <= amp_limit:
                    self._calculated_ground_wire = wire_size
                    return

            # If nothing matches
            self._calculated_ground_wire = None

        except Exception as e:
            self._calculated_ground_wire = None

    def calculate_hot_wire_quantity(self):
        return self.poles or 0

    def calculate_neutral_quantity(self):
        # Placeholder for future logic
        return 0

    def calculate_ground_wire_quantity(self):
        return 1

    def calculate_isolated_ground_quantity(self):
        # Placeholder for future logic
        return 0

    def get_downstream_demand_current(self):
        try:
            for el in self.circuit.Elements:
                if isinstance(el, DB.FamilyInstance):
                    param = el.get_Parameter(DB.BuiltInParameter.RBS_ELEC_PANEL_TOTAL_DEMAND_CURRENT_PARAM)
                    if param and param.StorageType == DB.StorageType.Double:
                        self._demand_current = param.AsDouble()
                        return self._demand_current
        except Exception as e:
            self.debug("Error in get_downstream_demand_current: {}".format(str(e)))
        self._demand_current = None
        return None

    def calculate_voltage_drop(self, wire_size, sets=1):
        try:
            length = self.length
            voltage = self.voltage
            poles = self.poles or 2
            pf = self.power_factor or 0.9
            phase = 3 if poles == 3 else 1

            # --- 🔽 Correct amperage based on feeder status ---
            if self.is_feeder:
                amps = self.get_downstream_demand_current()  # new method you define
            else:
                amps = self.apparent_current

            if not amps or not length or not voltage:
                return None

            material = self.wire_info.get('material', 'Copper')
            conduit_type = self.wire_info.get('conduit_type', 'PVC')  # still needs to be set
            impedance = WIRE_IMPEDANCE_TABLE.get(wire_size)
            if not impedance:
                return None

            R = impedance['R'].get(material, {}).get(conduit_type)
            X = impedance['XL'].get(conduit_type)
            if R is None or X is None:
                return None

            R = R / sets
            X = X / sets
            sin_phi = (1 - pf ** 2) ** 0.5

            if phase == 3:
                drop = (1.732 * amps * (R * pf + X * sin_phi) * length) / 1000.0
            else:
                drop = (2 * amps * (R * pf + X * sin_phi) * length) / 1000.0

            percent = (drop / voltage) * 100
            return percent

        except Exception as e:
            self.debug("Error in calculate_voltage_drop: {}".format(str(e)))
            return None

    def print_info(self, include_wire_info=True, include_all_properties=False):
        print("\n=== CircuitBranch: {} (ID: {}) ===".format(self.name, self.circuit_id))

        # Wire Info
        if include_wire_info:
            print("\nWire Info:")
            if self.wire_info:
                for key, val in self.wire_info.items():
                    print("    {}: {}".format(key, val if val else "N/A"))
            else:
                print("    No wire info available.")

        print("Feeder: {}".format(self.is_feeder))
        if self.is_feeder:
            current_source = self.get_downstream_demand_current()
            print("Current used for voltage drop: {:.2f} A (Feeder demand)".format(current_source or 0))
        else:
            print("Current used for voltage drop: {:.2f} A (Circuit apparent)".format(self.apparent_current or 0))

        # Circuit Info
        print("\nCircuit Info:")
        info_fields = [
            'rating', 'voltage', 'length',
            'apparent_power', 'apparent_current',
            'poles', 'power_factor'
        ]
        for attr in info_fields:
            try:
                value = getattr(self, attr)
                print("    {}: {}".format(attr, value if value is not None else "N/A"))
            except:
                print("    {}: [Error]".format(attr))

        if self._calculated_hot_wire and self._calculated_hot_sets:
            # First wire tried: Revit rating, min wire size
            initial_vd = self.calculate_voltage_drop(self.settings.min_wire_size, 1)
            final_vd = self.calculate_voltage_drop(self._calculated_hot_wire, self._calculated_hot_sets)

            print("\nVoltage Drop Calculation:")
            print("    Base breaker size (from Revit): {} A".format(self.rating))
            print("    Initial VD ({} x 1 set): {:.2f}%".format(self.settings.min_wire_size, initial_vd or 0))
            print("    Final wire: {} x {} set(s)".format(self._calculated_hot_wire, self._calculated_hot_sets))
            print("    Final voltage drop: {:.2f}% (Max allowed: {}%)".format(final_vd or 0, self.max_voltage_drop))
        else:
            print("\nVoltage Drop Calculation: [⚠️ Sizing failed or skipped]")

        # Calculation Results
        print("\nCalculated/Resolved Values:")
        print("    Breaker Rating: {}".format(self.breaker_rating or "N/A"))
        print("    Hot Wire Size: {}".format(self.hot_wire_size or "N/A"))
        print("    Number of Sets: {}".format(self.number_of_sets or "N/A"))
        print("    Circuit Base Ampacity: {}".format(self.circuit_base_ampacity or "N/A"))
        print("    Ground Wire Size: {}".format(self.ground_wire_size or "N/A"))

        # Quantities
        print("\nWire Quantities:")
        print("    Hot Conductors: {}".format(self.calculate_hot_wire_quantity()))
        print("    Ground Conductors: {}".format(self.calculate_ground_wire_quantity()))
        print("    Neutral Conductors: {}".format(self.calculate_neutral_quantity()))
        print("    Isolated Ground Conductors: {}".format(self.calculate_isolated_ground_quantity()))


def main():
    test_condition = 1

    if test_condition == 0:
        test_circuit = revit.get_selection()

    else:
        test_circuit = eu.pick_circuits_from_list(doc, select_multiple=True)

    for circuit in test_circuit:
        branch = CircuitBranch(circuit)
        branch.calculate_breaker_size()
        branch.calculate_hot_wire_size()
        branch.calculate_ground_wire_size()

        branch.print_info()
        # branch.debug("Post Calculation")



if __name__ == "__main__":
    main()
