# Revit MCP Server – Tool Reference

Reference for all tools exposed by the **user-Revit MCP Server**. Use this so the AI (or other consumers) knows what is available and how to call it.

---

## Model & status

### get_revit_status
**Description:** Check if the Revit MCP API is active and responding.

**Arguments:** None.

**Returns:** Result string (e.g. status/health).

---

### get_revit_model_info
**Description:** Get comprehensive information about the current Revit model (project info, levels, element counts, documentation stats, linked models, warnings, etc.).

**Arguments:** None.

**Returns:** Result string with model summary.

---

## Views

### get_current_view_info
**Description:** Get detailed information about the currently active view (name, type, ID, scale, detail level, crop box, view family type, discipline, template status).

**Arguments:** None.

**Returns:** Result string with view details.

---

### get_current_view_elements
**Description:** Get all elements visible in the currently active view. Returns element ID, name, type, category, level, location, and summary statistics by category.

**Arguments:** None.

**Returns:** Result string with element list and stats.

---

### list_revit_views
**Description:** Get a list of all exportable views in the current Revit model.

**Arguments:** None.

**Returns:** Result string (list of views).

---

### get_revit_view
**Description:** Export a specific Revit view as an image.

**Arguments:**
| Name        | Type   | Required | Description              |
|-------------|--------|----------|--------------------------|
| view_name   | string | Yes      | Name of the view to export |

**Returns:** Result string (e.g. path or status).

---

## Levels & families

### list_levels
**Description:** Get a list of all levels in the current Revit model.

**Arguments:** None.

**Returns:** Result string (list of levels).

---

### list_family_categories
**Description:** Get a list of all family categories in the current Revit model.

**Arguments:** None.

**Returns:** Result string (list of categories).

---

### list_families
**Description:** Get a flat list of available family types in the current Revit model.

**Arguments:**
| Name     | Type    | Required | Default | Description                          |
|----------|---------|----------|---------|--------------------------------------|
| contains | string  | No       | null    | Filter names containing this string  |
| limit    | integer | No       | 50      | Max number of results to return      |

**Returns:** Result string (list of family types).

---

### place_family
**Description:** Place a family instance at a specified location in the Revit model.

**Arguments:**
| Name         | Type   | Required | Default | Description                    |
|--------------|--------|----------|---------|--------------------------------|
| family_name  | string | Yes      | —       | Name of the family             |
| type_name    | string | No       | null    | Type name (if multiple types)  |
| x            | number | No       | 0       | X coordinate                   |
| y            | number | No       | 0       | Y coordinate                   |
| z            | number | No       | 0       | Z coordinate                   |
| rotation     | number | No       | 0       | Rotation (e.g. degrees)        |
| level_name   | string | No       | null    | Level to place on               |
| properties   | object | No       | null    | Optional instance properties   |

**Returns:** Result string (e.g. element ID or status).

---

## Coloring & parameters

### list_category_parameters
**Description:** Get available parameters for elements in a category. Use to discover what parameters exist for coloring or filtering (e.g. "Walls", "Doors").

**Arguments:**
| Name           | Type   | Required | Description                          |
|----------------|--------|----------|--------------------------------------|
| category_name  | string | Yes      | Category to list parameters for     |

**Returns:** Result string (parameters with types and sample values).

---

### color_splash
**Description:** Color elements in a category by parameter values. Same parameter value gets the same color.

**Arguments:**
| Name           | Type    | Required | Default | Description                              |
|----------------|---------|----------|---------|------------------------------------------|
| category_name  | string  | Yes      | —       | Category to color (e.g. Walls, Doors)     |
| parameter_name | string  | Yes      | —       | Parameter to drive colors (e.g. Mark)    |
| use_gradient   | boolean | No       | false   | Use gradient instead of distinct colors   |
| custom_colors  | array   | No       | null    | Optional hex colors, e.g. ["#FF0000"]    |

**Returns:** Result string (stats and color assignments).

---

### clear_colors
**Description:** Clear color overrides for elements in a category, restoring default appearance.

**Arguments:**
| Name           | Type   | Required | Description                        |
|----------------|--------|----------|------------------------------------|
| category_name  | string | Yes      | Category to clear (e.g. Walls)     |

**Returns:** Result string (e.g. count of elements processed).

---

## Code execution & scripting

### execute_revit_code
**Description:** Execute IronPython code in Revit with strict JSON I/O and error handling. Code runs in the Revit process; modifying the document (transactions) may not be allowed depending on context—use the ExternalEvent pattern to run commands that need transactions (see [mcp-call-pyrevit-button.md](mcp-call-pyrevit-button.md)).

**Arguments:**
| Name        | Type   | Required | Default           | Description              |
|-------------|--------|----------|-------------------|--------------------------|
| code        | string | Yes      | —                 | IronPython code to run   |
| description | string | No       | "Code execution"  | Short label for the run  |

**Returns:** Result string (output or error message).

---

### crystallize_script
**Description:** Turn a Python script into a pyRevit pushbutton. Creates a new pushbutton in the MG_Tools panel (script saved as a `.pushbutton` folder with `script.py`). Code must be IronPython 2.7–compatible (no f-strings, async/await, type hints; simple ASCII). If the folder exists, a numeric suffix is added (e.g. `MyScript_2.pushbutton`).

**Arguments:**
| Name        | Type   | Required | Description                                              |
|-------------|--------|----------|----------------------------------------------------------|
| script_name | string | Yes      | Pushbutton name without `.pushbutton` (e.g. "RenumberDoors") |
| script_code | string | Yes      | Full IronPython 2.7–compatible script content            |

**Returns:** Result string (e.g. success message and path to `script.py`).

---

## Summary table

| Tool                      | Purpose                                      |
|---------------------------|----------------------------------------------|
| get_revit_status          | Health check                                 |
| get_revit_model_info      | Full model summary                           |
| get_current_view_info     | Active view details                          |
| get_current_view_elements | Elements in active view                      |
| list_revit_views          | All exportable views                         |
| get_revit_view            | Export view as image                         |
| list_levels               | All levels                                   |
| list_family_categories    | All family categories                        |
| list_families             | Family types (optional filter/limit)         |
| place_family              | Place family instance at x,y,z               |
| list_category_parameters  | Parameters for a category                    |
| color_splash              | Color category by parameter                  |
| clear_colors              | Clear category color overrides               |
| execute_revit_code        | Run IronPython in Revit                      |
| crystallize_script        | Save script as pyRevit pushbutton            |
