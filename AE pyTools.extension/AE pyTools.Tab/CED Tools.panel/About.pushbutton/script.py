# -*- coding: utf-8 -*-
"""CED About window."""

import os

from pyrevit import forms, script, versionmgr

TITLE = "About AE pyTools"
UPDATE_AEPYTOOLS_URL = "https://coolsysinc.sharepoint.com/:u:/s/Teams-CED-Admin/IQDes9bTY0rqQqoDO87G_HuqAS_KIZ-hlQZvO_Jof3FGlcQ?e=JK8jrC"
UPDATE_PYREVIT_URL = "https://coolsysinc.sharepoint.com/:f:/s/Teams-CED-Admin/IgCCzGCyEpCERar1EUegCQtZAdYnru-wKl7y5QixPsg45qo?e=grdKjl"
REPORT_BUGS_URL = "https://coolsysinc.sharepoint.com/sites/Teams-Coolsys-ToolDevelopment/_layouts/15/listforms.aspx?cid=NGFkYTZkYzUtMmI3Mi00NmQyLTlmZDYtNzczNjk2Y2ZmYjM5&nav=ZTBkMWYyOTgtMjk0ZS00ZmQzLThjOWItYTlhZjI1YmFlMzRm"
REQUEST_FEATURES_URL = "https://coolsysinc.sharepoint.com/sites/Teams-Coolsys-ToolDevelopment/_layouts/15/listforms.aspx?cid=MTliNjNhOTktM2JkMC00NDk3LWI0NzYtYTdiNTkxZWNjNzM3&nav=NjczZmU2MzYtMGNhNC00NDI1LWE4ODYtY2U5ODE4N2U3NDli"
DESIGNER_HUB_URL = "https://coolsysinc.sharepoint.com/sites/ToolDevelopment-Public"
TRAINING_URL = "https://productivitynow.imaginit.com/peak/libraries/cc3ea5f2-0cb5-4957-8489-2f80f08f78e8"


def _read_about_yaml(yaml_path):
    version = None
    authors = []
    try:
        with open(yaml_path, "r") as yaml_file:
            for raw_line in yaml_file:
                stripped = str(raw_line or "").strip()
                if not stripped or stripped.startswith("#"):
                    continue
                if stripped.startswith("toolbar_version:"):
                    version = stripped.split(":", 1)[1].strip()
                    continue
                if stripped.startswith("- "):
                    authors.append(stripped[2:].strip())
    except Exception:
        pass
    return version, authors


def _format_toolbar_version(value):
    text = str(value or "").strip()
    if not text:
        return "Unknown"
    if text.lower().startswith("v"):
        return text
    return "v{0}".format(text)


def _get_pyrevit_version_text():
    try:
        pyrvt_ver = versionmgr.get_pyrevit_version()
        if pyrvt_ver:
            return "v{0}".format(pyrvt_ver.get_formatted())
    except Exception:
        pass
    try:
        cli_ver = str(versionmgr.get_pyrevit_cli_version() or "").strip()
        if cli_ver:
            return "v{0}".format(cli_ver.lstrip("vV"))
    except Exception:
        pass
    return "Unknown"


class AboutWindow(forms.WPFWindow):
    def __init__(self, xaml_path, logo_path, toolbar_version, authors):
        forms.WPFWindow.__init__(self, xaml_path)
        self.toolbar_version_tb.Text = _format_toolbar_version(toolbar_version)
        self.pyrevit_version_tb.Text = _get_pyrevit_version_text()
        self.authors_tb.Text = ", ".join(list(authors or [])) or "Not listed"
        if logo_path and os.path.exists(logo_path):
            self.set_image_source(self.ced_logo_img, logo_path)

    def open_update_ae_pytools(self, sender, args):
        script.open_url(UPDATE_AEPYTOOLS_URL)

    def open_update_pyrevit(self, sender, args):
        script.open_url(UPDATE_PYREVIT_URL)

    def open_report_bugs(self, sender, args):
        script.open_url(REPORT_BUGS_URL)

    def open_request_features(self, sender, args):
        script.open_url(REQUEST_FEATURES_URL)

    def open_designer_hub(self, sender, args):
        script.open_url(DESIGNER_HUB_URL)

    def open_training_videos(self, sender, args):
        script.open_url(TRAINING_URL)

    def close_window(self, sender, args):
        self.Close()


def main():
    this_dir = os.path.abspath(os.path.dirname(__file__))
    xaml_path = os.path.join(this_dir, "AboutWindow.xaml")
    logo_path = os.path.join(this_dir, "CEDLogo.png")
    yaml_path = os.path.join(this_dir, "about.yaml")

    if not os.path.exists(xaml_path):
        forms.alert("Could not find AboutWindow.xaml.", title=TITLE, ok=True)
        return

    toolbar_version, authors = _read_about_yaml(yaml_path)
    dlg = AboutWindow(xaml_path, logo_path, toolbar_version, authors)
    dlg.ShowDialog()


if __name__ == "__main__":
    main()
