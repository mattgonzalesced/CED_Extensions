# -*- coding: utf-8 -*-
# Edit Element Linker profiles (CadBlockProfile / TypeConfig) at runtime

from pyrevit import script, forms

from Element_Linker import CAD_BLOCK_PROFILES
from ProfileEditorWindow import ProfileEditorWindow


def main():
    xaml_path = script.get_bundle_file('ProfileEditorWindow.xaml')
    if not xaml_path:
        forms.alert(
            "ProfileEditorWindow.xaml not found in the bundle.",
            title="Edit Element Linker Profiles"
        )
        return

    window = ProfileEditorWindow(xaml_path, CAD_BLOCK_PROFILES)
    window.show_dialog()
    # Changes are applied in-memory to CAD_BLOCK_PROFILES via the window


if __name__ == "__main__":
    main()