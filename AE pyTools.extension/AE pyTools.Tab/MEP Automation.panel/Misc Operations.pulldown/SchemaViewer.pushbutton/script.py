import os
import sys
import io

sys.path.append(os.path.join(r"C:\CED_Extensions", "CEDLib.lib"))

from ExtensibleStorage import yaml_store  # noqa: E402

path = r"C:\CED_Extensions\CEDLib.lib\prototypeHEBtest12.yaml"
with io.open(path, "r", encoding="utf-8") as handle:
    raw = handle.read()
san = yaml_store._sanitize_hash_keys(raw)
print("raw_has_hash =", "#" in raw)
print("san_has_hash =", "#" in san)
print("san_equals_raw =", san == raw)
