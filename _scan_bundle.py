"""Download FE JS bundle and grep for endpoints + keywords."""
import urllib.request, re, sys
from pathlib import Path

u = 'https://queue-fe-uat.doe.go.th/doe/assets/index-Dpwm8sYV.js'
try:
    text = urllib.request.urlopen(u, timeout=30).read().decode('utf-8', errors='replace')
except Exception as e:
    print('download failed:', e); sys.exit(1)
print(f'downloaded {len(text):,} bytes')
Path('_probe_bundle.js').write_text(text, encoding='utf-8')

# scan for url string literals
patterns = [
    r'/api/v\d/[a-zA-Z0-9/_\-{}\$?=&.]+',
    r'`/api/[^`]{0,120}`',
    r'"/doe-booking[^"]{0,120}"',
    r'`[^`]{0,100}\$\{[^}]+\}[^`]{0,60}`',   # template-literal urls
]
found = set()
for p in patterns:
    for m in re.finditer(p, text):
        s = m.group(0)
        if '/api' in s or 'booking' in s:
            found.add(s[:200])
print(f'\nfound {len(found)} url-like strings:')
for x in sorted(found):
    print(' ', x)

# keyword hits (context)
KWS = ['calendar/get-data','get-time','timeslot','time-slot','time_slot',
       '/period','/slot','/hour','booking-time','schedule','reserve',
       'branch/by-codes','round','extra_book']
for kw in KWS:
    matches = [m.start() for m in re.finditer(re.escape(kw), text)]
    if matches:
        print(f'\n[{kw}] {len(matches)} hits')
        for pos in matches[:4]:
            snip = text[max(0,pos-90):pos+140].replace(chr(10),' ')
            print('  ...', snip, '...')
