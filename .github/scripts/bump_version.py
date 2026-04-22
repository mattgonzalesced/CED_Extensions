import os

import yaml

file_path = "AE pyTools.extension/AE pyTools.Tab/CED Tools.panel/About.pushbutton/about.yaml"

with open(file_path) as f:
    data = yaml.safe_load(f)

version = data.get("toolbar_version", "0.0.0")
major, minor, patch = [int(x) for x in version.split(".")]

labels = os.environ.get("PR_LABELS", "")

if "release: major" in labels:
    major += 1
    minor = 0
    patch = 0
elif "release: minor" in labels:
    minor += 1
    patch = 0
elif "release: patch" in labels:
    patch += 1
else:
    raise Exception("No valid release label")

data["toolbar_version"] = "{}.{}.{}".format(major, minor, patch)

with open(file_path, "w") as f:
    yaml.safe_dump(data, f, default_flow_style=False, sort_keys=False)

print("New version:", data["toolbar_version"])