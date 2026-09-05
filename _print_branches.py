"""Print all branch codes from probe result"""
import json
from pathlib import Path

data = json.loads(Path("_probe_appt_booking_out.json").read_text(encoding="utf-8"))
for r in data.get("captured_api", []):
    if "branch/by-codes" in r["url"]:
        try:
            parsed = json.loads(r["body_preview"])
            if not isinstance(parsed, list):
                # response was truncated, try to close the JSON
                continue
            print(f"branches count: {len(parsed)}")
            for br in parsed:
                bid = br.get("branch_code_id", "")
                bname = br.get("branch_name_th", "")
                print(f"  {bid:20s} | {bname[:80]}")
            break
        except json.JSONDecodeError:
            # response might be truncated (2000 char preview)
            # try to salvage: parse until last complete object
            body = r["body_preview"]
            print(f"[!] JSON parse fail, body {len(body)} chars — truncated")
            # attempt to find "branch_code_id" entries manually
            import re
            for m in re.finditer(r'"branch_code_id":"([^"]+)"[^}]*"branch_name_th":"([^"]+)"', body):
                print(f"  {m.group(1):20s} | {m.group(2)[:80]}")
            break
