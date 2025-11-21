# -*- coding: utf-8 -*-
"""
CSV utilities shared by the Place Elements command.
"""

import csv
import codecs


def feet_inch_to_inches(value):
    """Parse strings like 5'-6 1/4" into total inches (float)."""
    try:
        if value is None:
            return None
        s = value.strip()
        if not s:
            return None
        s = s.replace('"', "")

        sign = 1.0
        if s.startswith("-"):
            sign = -1.0
            s = s[1:].strip()

        feet = 0.0
        inches = 0.0
        if "'" in s:
            ft_part, rest = s.split("'", 1)
            ft_part = ft_part.strip()
            if ft_part:
                feet = float(ft_part)
            s = rest.strip()
        else:
            s = s.strip()

        if s:
            parts = s.split()
            if len(parts) == 1:
                if "/" in parts[0]:
                    num, den = parts[0].split("/")
                    inches = float(num) / float(den)
                else:
                    inches = float(parts[0])
            elif len(parts) == 2:
                whole = float(parts[0])
                num, den = parts[1].split("/")
                inches = whole + (float(num) / float(den))

        return sign * (feet * 12.0 + inches)
    except Exception:
        return None


def read_xyz_csv(csv_path):
    """
    Reads the CAD CSV and returns (rows, names).
    Skips rows where Count != "1" or Position X is blank.
    """
    xyz_rows = []
    unique_names = set()

    with codecs.open(csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f, delimiter=",")
        for row in reader:
            if row.get("Count", "").strip() != "1":
                continue
            if not row.get("Position X", "").strip():
                continue
            xyz_rows.append(row)
            cad_name = (row.get("Name") or "").strip()
            if cad_name:
                unique_names.add(cad_name)

    return xyz_rows, list(unique_names)


__all__ = ["feet_inch_to_inches", "read_xyz_csv"]
