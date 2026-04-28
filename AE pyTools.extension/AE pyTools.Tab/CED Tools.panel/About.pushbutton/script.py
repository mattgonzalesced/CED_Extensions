# -*- coding: utf-8 -*-
"""CED About window."""

import os
import re

from pyrevit import forms, script, versionmgr

TITLE = "About AE pyTools"
UPDATE_AEPYTOOLS_URL = "https://coolsysinc.sharepoint.com/:u:/s/Teams-CED-Admin/IQDes9bTY0rqQqoDO87G_HuqAS_KIZ-hlQZvO_Jof3FGlcQ?e=JK8jrC"
UPDATE_PYREVIT_URL = "https://coolsysinc.sharepoint.com/:f:/s/Teams-CED-Admin/IgCCzGCyEpCERar1EUegCQtZAdYnru-wKl7y5QixPsg45qo?e=grdKjl"
REPORT_BUGS_URL = "https://coolsysinc.sharepoint.com/:l:/s/Teams-Coolsys-ToolDevelopment/JADFbdpKcivSRp_WdzaWz_s5Ac-juj4VItpGkCr8KghJd4Y?nav=ZTBkMWYyOTgtMjk0ZS00ZmQzLThjOWItYTlhZjI1YmFlMzRm"
REQUEST_FEATURES_URL = "https://coolsysinc.sharepoint.com/:l:/s/Teams-Coolsys-ToolDevelopment/JACZOrYZ0DuXRLR2p7WR7Mc3Aemey6nPdkhK9ghcwegqDzg?nav=NjczZmU2MzYtMGNhNC00NDI1LWE4ODYtY2U5ODE4N2U3NDli"
DESIGNER_HUB_URL = "https://coolsysinc.sharepoint.com/sites/ToolDevelopment-Public"
TRAINING_URL = "https://productivitynow.imaginit.com/peak/libraries/cc3ea5f2-0cb5-4957-8489-2f80f08f78e8"


def _read_about_yaml(yaml_path):
    version = None
    build = None
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
                if stripped.startswith("build:"):
                    build = stripped.split(":", 1)[1].strip()
                    continue
                if stripped.startswith("- "):
                    authors.append(stripped[2:].strip())
    except Exception:
        pass
    return version, build, authors


def _normalize_build(value):
    text = str(value or "").strip()
    if not text:
        return ""
    if re.match(r"^\d{8}\+\d{4}$", text):
        return text
    if re.match(r"^\d{12}$", text):
        return "{0}+{1}".format(text[:8], text[8:])
    return ""


def _format_toolbar_version(version_value, build_value):
    version_text = str(version_value or "").strip()
    if not version_text:
        return "Unknown"
    version_text = version_text.lstrip("vV")
    build_text = _normalize_build(build_value)
    if build_text:
        return "v{0}.{1}".format(version_text, build_text)
    return "v{0}".format(version_text)


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
    def __init__(self, xaml_path, logo_path, toolbar_version, build, authors):
        forms.WPFWindow.__init__(self, xaml_path)
        self.toolbar_version_tb.Text = _format_toolbar_version(toolbar_version, build)
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

    toolbar_version, build, authors = _read_about_yaml(yaml_path)
    dlg = AboutWindow(xaml_path, logo_path, toolbar_version, build, authors)
    dlg.ShowDialog()


if __name__ == "__main__":
    main()
