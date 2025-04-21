def debug_element_parameters(element):
    print("Parameters for Equipment ID {}:".format(element.Id.IntegerValue))
    for p in element.Parameters:
        name = p.Definition.Name
        raw_val = p.AsString() if p.StorageType == DB.StorageType.String else None
        disp_val = p.AsValueString()
        print("  {}: raw='{}', display='{}'".format(name, raw_val, disp_val))
