# -*- coding: utf-8 -*-
"""
Schema Viewer
-------------
Displays the current YAML history stored inside Extensible Storage, including a
decoded view of the latest YAML text.
"""

import base64
import json
import zlib

from Autodesk.Revit.DB.ExtensibleStorage import Schema
from System import Guid
from pyrevit import script, forms

SCHEMA_GUID = Guid("9f6633b1-d77f-49ef-9390-5111fbb16d82")


def _decode(text):
    if not text:
        return ""
    return zlib.decompress(base64.b64decode(text.encode("ascii"))).decode("utf-8")


doc = __revit__.ActiveUIDocument.Document
if doc is None:
    forms.alert("No active document detected.", title="Schema Viewer")
    raise SystemExit

schema = Schema.Lookup(SCHEMA_GUID)
if schema is None:
    forms.alert("No CED_YamlHistory schema found. Run Select YAML first.", title="Schema Viewer")
    raise SystemExit

entity = doc.ProjectInformation.GetEntity(schema)
if entity is None or not entity.IsValid():
    forms.alert("ProjectInformation does not contain CED_YamlHistory data.", title="Schema Viewer")
    raise SystemExit

history_field = schema.GetField("HistoryJson")
meta_field = schema.GetField("MetadataJson")

history_json = entity.Get[str](history_field) or "[]"
metadata_json = entity.Get[str](meta_field) or "{}"

history = []
metadata = {}
try:
    history = json.loads(history_json)
except Exception:
    history = []

try:
    metadata = json.loads(metadata_json)
except Exception:
    metadata = {}

output = script.get_output()
output.print_md("### CED_YamlHistory")

if history:
    latest = history[-1]
    try:
        decoded_text = _decode(latest.get("new_content"))
    except Exception:
        decoded_text = "<error decoding new_content>"
    output.print_md("**Latest YAML Text**\n```yaml\n{}\n```".format(decoded_text or "<empty>"))
else:
    output.print_md("**Latest YAML Text**\n```\n<no history entries>\n```")

output.print_md("**HistoryJson (raw)**\n```\n{}\n```".format(history_json))
output.print_md("**MetadataJson (raw)**\n```\n{}\n```".format(metadata_json))
