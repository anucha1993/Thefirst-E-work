import json
import sys

sys.stdout.reconfigure(encoding='utf-8')
d = json.load(open('step5_probe.json', encoding='utf-8'))
s = d.get('step5_after_5_2', {})

print('=== FULL BODY SAMPLE (5.3) ===')
body = s.get('bodySample', '') or ''
print(body)
print()
print('=== ALL CHECKBOXES (incl hidden) ===')
for c in s.get('checkboxes', []):
    print(f"id={c.get('id','')!r} name={c.get('name','')!r} checked={c.get('checked')} visible={c.get('visible')}")
    print(f"   label: {c.get('label','')[:300]}")
    print(f"   near : {c.get('near','')[:300]}")
    print()
print('=== ALL FILE INPUTS ===')
for f in s.get('fileInputs', []):
    print(f"id={f.get('id','')!r} name={f.get('name','')!r} accept={f.get('accept','')!r} visible={f.get('visible')}")
    for da in f.get('dataAttrs', []):
        print(f"   {da[:140]}")
    print()
print('=== ALL VISIBLE BUTTONS (full) ===')
for b in s.get('buttons', []):
    print(f"  {b.get('text','')!r:60} id={b.get('id','')!r:24} da={b.get('da','')!r:18} type={b.get('type','')!r:8} onclick={b.get('onclick','')[:60]}")
