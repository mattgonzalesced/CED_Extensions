# UIClasses Framework

`UIClasses` now includes reusable foundations for new pyRevit UIs:

- `resource_loader.py`
  - Loads centralized dictionaries from `UIClasses/Resources`.
  - Applies theme (`light`, `dark`, `dark_alt`) and accent (`blue`, `red`, `green`, `neutral`).
- `ui_bases.py`
  - `CEDWindowBase`: base for modeless/modal windows.
  - `CEDPanelBase`: base for dockable panels.
- `Resources/Templates/WindowChrome.xaml`
  - Base styles for window/page chrome and section cards.
- `Resources/Templates/ControlPrimitives.xaml`
  - Common icon, separator, empty-state, and overlay primitives.

## Minimal usage

```python
from UIClasses.ui_bases import CEDWindowBase

class MyWindow(CEDWindowBase):
    theme_aware = True
    auto_wire_textboxes = True
    text_select_all_on_click = True
    text_select_all_on_focus = True

    def __init__(self):
        CEDWindowBase.__init__(
            self,
            xaml_source="MyWindow.xaml",  # optional when class-name xaml exists
        )
```

### What is automatic now

- Resolves module-relative XAML path (e.g. `"MyWindow.xaml"`).
- Infers XAML automatically (`<ClassName>.xaml` / `<ModuleName>.xaml`) if `xaml_source` is omitted.
- Resolves workspace/lib/resources paths and appends `CEDLib.lib` to `sys.path`.
- Loads theme/accent from `AE-pyTools-Theme` config when `theme_aware = True`.
- Falls back to Light/Blue when `theme_aware = False`.
- Optional shift+mousewheel horizontal scroll handling.
- Optional textbox select-all behavior (click/focus) via class flags.

For existing tools, migration can stay incremental: move to these base classes first, then remove duplicated local path/theme/input wiring per tool.
