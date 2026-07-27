import json
d = json.load(open("_probe_status_checkboxes.json", "r", encoding="utf-8"))
for cb in d["checkboxes"]:
    if cb["visible"]:
        print(
            f"id={cb['id']:12s} "
            f"name={cb['name']!r:35s} "
            f"value={cb['value']!r:8s} "
            f"label={cb['labelText'][:60]!r}"
        )
