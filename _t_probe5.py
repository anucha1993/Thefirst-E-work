import json
import sys

sys.stdout.reconfigure(encoding='utf-8')
d = json.load(open('step5_probe.json', encoding='utf-8'))
for k in ['step5_initial', 'step5_after_5_1', 'step5_after_5_2']:
    s = d.get(k, {})
    print(f'=== {k} ===')
    print('  url:', s.get('url', ''))
    print('  bodySample[:800]:', repr((s.get('bodySample', '') or '')[:800]))
    print('  buttons (filtered):')
    for b in s.get('buttons', [])[:50]:
        t = b.get('text', '')
        if any(kw in t for kw in ['ถัด', 'ก่อน', 'ยืนยัน', 'บันทึก', 'อัปโหลด', 'อัพโหลด', 'ภาพ',
                                   'ตรวจ', 'ส่ง', 'รับทราบ', 'ข้าพเจ้า']) or b.get('da') == 'next':
            print(f"    {t!r:50} id={b.get('id','')!r:24} da={b.get('da','')!r:18} cls={b.get('cls','')[:60]}")
    print('  checkboxes total:', len(s.get('checkboxes', [])),
          'visible:', sum(1 for c in s.get('checkboxes', []) if c.get('visible')))
    for c in s.get('checkboxes', []):
        if c.get('visible'):
            print(f"    cb id={c.get('id','')!r} name={c.get('name','')!r} checked={c.get('checked')}")
            print(f"       label={c.get('label','')[:100]}")
            print(f"       near={c.get('near','')[:100]}")
    print('  fileInputs total:', len(s.get('fileInputs', [])),
          'visible:', sum(1 for f in s.get('fileInputs', []) if f.get('visible')))
    for f in s.get('fileInputs', []):
        if f.get('visible'):
            print(f"    file id={f.get('id','')!r} accept={f.get('accept','')[:40]}")
            da = f.get('dataAttrs', [])
            if da:
                print(f"       data: {da[0][:120]}")
    print('  imgs (visible):')
    for im in s.get('imgs', [])[:8]:
        print(f"    img src=...{im.get('src','')} id={im.get('id','')!r} cls={im.get('cls','')[:40]}")
    print()
