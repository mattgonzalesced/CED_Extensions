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
    def __init__(self, xaml_path, resources_root):
        CEDWindowBase.__init__(
            self,
            xaml_path=xaml_path,
            resources_root=resources_root,
            theme_mode="light",
            accent_mode="blue",
        )
```

For existing tools, migration can be incremental: keep tool logic as-is and move only resource/theme wiring into these bases first.
