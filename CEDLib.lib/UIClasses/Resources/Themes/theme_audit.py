import os
import xml.etree.ElementTree as ET

# -----------------------------
# CONFIG
# -----------------------------

ROOT = r"C:\Users\Aevelina\CED_Extensions\CEDLib.lib\UIClasses\Resources\Themes"

THEME_FILES = {
    "Base": "CED.Colors.xaml",
    "Dark": "CEDTheme.Dark.xaml",
    "DarkAlt": "CEDTheme.DarkAlt.xaml"
}

# -----------------------------
# HELPERS
# -----------------------------

def parse_colors(file_path):
    """Return dict of {key: color_value} from a XAML file"""
    colors = {}

    if not os.path.exists(file_path):
        print("Missing file: " + file_path)
        return colors

    try:
        tree = ET.parse(file_path)
        root = tree.getroot()
    except Exception as e:
        print("Failed to parse: " + file_path)
        return colors

    # XAML namespace handling
    for elem in root.iter():
        tag = elem.tag

        # strip namespace if present
        if "}" in tag:
            tag = tag.split("}", 1)[1]

        if tag == "Color":
            key = elem.attrib.get("{http://schemas.microsoft.com/winfx/2006/xaml}Key")
            value = elem.text

            if key:
                colors[key] = value

    return colors


# -----------------------------
# MAIN
# -----------------------------

def build_color_table():
    """Build merged color table across themes"""

    theme_data = {}
    all_keys = set()

    # Parse each theme
    for theme_name, filename in THEME_FILES.items():
        path = os.path.join(ROOT, filename)
        colors = parse_colors(path)

        theme_data[theme_name] = colors
        all_keys.update(colors.keys())

    return theme_data, sorted(all_keys)


def print_table(theme_data, all_keys):
    """Print tab-delimited table for Excel"""

    themes = list(THEME_FILES.keys())

    # Header
    header = ["ColorName"] + themes
    print("\t".join(header))

    # Rows
    for key in all_keys:
        row = [key]

        for theme in themes:
            value = theme_data.get(theme, {}).get(key, "")
            row.append(value if value else "")

        print("\t".join(row))


# -----------------------------
# RUN
# -----------------------------

if __name__ == "__main__":
    theme_data, all_keys = build_color_table()
    print_table(theme_data, all_keys)