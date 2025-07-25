# -*- coding: utf-8 -*-
import Autodesk.Revit.DB.Plumbing as DBP
from pyrevit import revit, DB, script, forms

output = script.get_output()
doc = revit.doc

def to_internal_inches(inches):
    return DB.UnitUtils.ConvertToInternalUnits(inches, DB.UnitTypeId.Inches)

def create_pipe(pipe_type, system_type_id, level_id, start, end, diameter):
    pipe = DBP.Pipe.Create(doc, system_type_id, pipe_type.Id, level_id, start, end)
    param = pipe.get_Parameter(DB.BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
    if param and not param.IsReadOnly:
        param.Set(diameter)
    return pipe

def find_connector_at_point(pipe, point):
    for c in pipe.ConnectorManager.Connectors:
        if c.Origin.IsAlmostEqualTo(point):
            return c
    return None

def create_union(pipe1, pt, system_type_id, level, start_point, diameter):
    mid = DB.XYZ(start_point.X + 5.0, start_point.Y, start_point.Z)
    pipe2 = create_pipe(pt, system_type_id, level.Id, start_point, mid, diameter)
    conn1 = find_connector_at_point(pipe1, start_point)
    conn2 = find_connector_at_point(pipe2, start_point)
    if conn1 and conn2:
        doc.Create.NewUnionFitting(conn1, conn2)
        output.print_md("Union fitting placed at {}.".format(start_point))
    return pipe2, mid

def create_reducer(pipe1, pt, system_type_id, level, start_point):
    diameter_increased = to_internal_inches(3.0)
    mid = DB.XYZ(start_point.X + 5.0, start_point.Y, start_point.Z)
    pipe2 = create_pipe(pt, system_type_id, level.Id, start_point, mid, diameter_increased)
    conn1 = find_connector_at_point(pipe1, start_point)
    conn2 = find_connector_at_point(pipe2, start_point)
    if conn1 and conn2:
        doc.Create.NewTransitionFitting(conn1, conn2)
        output.print_md("Reducer fitting placed at {}.".format(start_point))
    return pipe2, mid

def create_tee(pipe1, pt, system_type_id, level, start_point):
    mid = DB.XYZ(start_point.X + 5.0, start_point.Y, start_point.Z)
    pipe2 = create_pipe(pt, system_type_id, level.Id, start_point, mid, to_internal_inches(3.0))
    branch_end = DB.XYZ(start_point.X, start_point.Y + 5.0, start_point.Z)
    branch_pipe = create_pipe(pt, system_type_id, level.Id, start_point, branch_end, to_internal_inches(3.0))
    conn1 = find_connector_at_point(pipe1, start_point)
    conn2 = find_connector_at_point(pipe2, start_point)
    conn3 = find_connector_at_point(branch_pipe, start_point)
    if conn1 and conn2 and conn3:
        doc.Create.NewTeeFitting(conn1, conn2, conn3)
        output.print_md("Tee fitting placed at {}.".format(start_point))
    return pipe2, mid

def create_cross(pipe_in, pt, system_type_id, level, start_point):
    forward_x = DB.XYZ(start_point.X + 2.5, start_point.Y, start_point.Z)
    pipe_x = create_pipe(pt, system_type_id, level.Id, start_point, forward_x, to_internal_inches(3.0))
    branch_up = DB.XYZ(start_point.X, start_point.Y + 2.5, start_point.Z)
    branch_down = DB.XYZ(start_point.X, start_point.Y - 2.5, start_point.Z)
    pipe_y_up = create_pipe(pt, system_type_id, level.Id, start_point, branch_up, to_internal_inches(3.0))
    pipe_y_down = create_pipe(pt, system_type_id, level.Id, branch_down, start_point, to_internal_inches(3.0))
    conn1 = find_connector_at_point(pipe_in, start_point)
    conn2 = find_connector_at_point(pipe_x, start_point)
    conn3 = find_connector_at_point(pipe_y_up, start_point)
    conn4 = find_connector_at_point(pipe_y_down, start_point)
    if conn1 and conn2 and conn3 and conn4:
        doc.Create.NewCrossFitting(conn1, conn2, conn3, conn4)
        output.print_md("Cross fitting placed at {}.".format(start_point))
    return pipe_x, forward_x

def create_elbow(pipe1, pt, system_type_id, level, start_point):
    end = DB.XYZ(start_point.X, start_point.Y, start_point.Z + 5.0)
    pipe2 = create_pipe(pt, system_type_id, level.Id, start_point, end, to_internal_inches(3.0))
    conn1 = find_connector_at_point(pipe1, start_point)
    conn2 = find_connector_at_point(pipe2, start_point)
    if conn1 and conn2:
        doc.Create.NewElbowFitting(conn1, conn2)
        output.print_md("Elbow fitting placed at {}.".format(start_point))
    return pipe2, end

def run_for_pipe_type(pt, system_type_id, level, offset_y):
    start = DB.XYZ(50.0, offset_y, 0.0)
    p2_start = DB.XYZ(start.X + 5.0, start.Y, start.Z)

    # Start with 2" pipes and union
    pipe1 = create_pipe(pt, system_type_id, level.Id, start, p2_start, to_internal_inches(2.0))
    pipe2, p2_end = create_union(pipe1, pt, system_type_id, level, p2_start, to_internal_inches(2.0))

    # Reducer to 3"
    pipe3, p3_end = create_reducer(pipe2, pt, system_type_id, level, p2_end)

    # Tee, cross, elbow
    pipe4, p4_end = create_tee(pipe3, pt, system_type_id, level, p3_end)
    pipe5, p5_end = create_cross(pipe4, pt, system_type_id, level, p4_end)
    pipe6, _ = create_elbow(pipe5, pt, system_type_id, level, p5_end)

def main_run():
    level = DB.FilteredElementCollector(doc).OfClass(DB.Level).FirstElement()
    system_type = DB.FilteredElementCollector(doc).OfClass(DBP.PipingSystemType).FirstElement()
    pipe_types = DB.FilteredElementCollector(doc).OfClass(DBP.PipeType).ToElements()
    if not level or not system_type or not pipe_types:
        forms.alert("Missing required elements.")
        script.exit()

    offset_y = 50
    for pt in pipe_types:
        pt_name = DB.Element.Name.__get__(pt)
        output.print_md("## Starting for Pipe Type: {}".format(pt_name))
        run_for_pipe_type(pt, system_type.Id, level, offset_y)
        offset_y += 10.0

with revit.Transaction("Draw Pipe Runs with All Fittings (All Pipe Types)"):
    main_run()

forms.alert("Finished creating pipe runs for all pipe types, offset by 10ft in Y.")
output.close_others()
