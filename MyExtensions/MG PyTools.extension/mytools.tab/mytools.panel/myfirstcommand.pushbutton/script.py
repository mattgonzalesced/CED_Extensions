from pyrevit import revit, DB
import math

# Define a function to create a line in the Revit document
def create_wall(doc, start_point, end_point, wall_type, level):
    """Creates a wall between two points."""
    try:
        line = DB.Line.CreateBound(start_point, end_point)
        wall = DB.Wall.Create(doc, line, wall_type.Id, level.Id, 10, 0, False, False)
        return wall
    except Exception as e:
        print("Error creating wall: {}".format(e))
        return None

# Define the coordinates for each letter in "HELLO"
def generate_hello_geometry(base_x, base_y, spacing):
    """Generates wall geometry for the word HELLO."""
    letters = []

    # "H"
    letters.append([(0, 0), (0, 5)])  # Left vertical
    letters.append([(3, 0), (3, 5)])  # Right vertical
    letters.append([(0, 2.5), (3, 2.5)])  # Horizontal middle

    # "E"
    base_x += spacing
    letters.append([(base_x, 0), (base_x, 5)])  # Vertical
    letters.append([(base_x, 5), (base_x + 3, 5)])  # Top horizontal
    letters.append([(base_x, 2.5), (base_x + 2, 2.5)])  # Middle horizontal
    letters.append([(base_x, 0), (base_x + 3, 0)])  # Bottom horizontal

    # "L"
    base_x += spacing
    letters.append([(base_x, 0), (base_x, 5)])  # Vertical
    letters.append([(base_x, 0), (base_x + 3, 0)])  # Bottom horizontal

    # "L"
    base_x += spacing
    letters.append([(base_x, 0), (base_x, 5)])  # Vertical
    letters.append([(base_x, 0), (base_x + 3, 0)])  # Bottom horizontal

    # "O"
    base_x += spacing
    letters.append([(base_x, 0), (base_x, 5)])  # Left vertical
    letters.append([(base_x + 3, 0), (base_x + 3, 5)])  # Right vertical
    letters.append([(base_x, 5), (base_x + 3, 5)])  # Top horizontal
    letters.append([(base_x, 0), (base_x + 3, 0)])  # Bottom horizontal

    return letters

# Main function
def main():
    doc = revit.doc
    base_point_x = 0
    base_point_y = 0
    spacing_between_letters = 6  # Space between letters

    # Get the active level and default wall type
    level = DB.FilteredElementCollector(doc).OfClass(DB.Level).FirstElement()
    wall_type = DB.FilteredElementCollector(doc).OfClass(DB.WallType).FirstElement()

    if not level or not wall_type:
        print("Could not find an active level or default wall type.")
        return

    # Generate wall geometry for "HELLO"
    hello_geometry = generate_hello_geometry(base_point_x, base_point_y, spacing_between_letters)

    # Start a transaction
    with revit.Transaction("Create Walls for 'HELLO'"):
        for segment in hello_geometry:
            start = DB.XYZ(segment[0][0], segment[0][1], 0)
            end = DB.XYZ(segment[1][0], segment[1][1], 0)
            create_wall(doc, start, end, wall_type, level)

    print("Walls created for 'HELLO'!")

# Run the script
main()
