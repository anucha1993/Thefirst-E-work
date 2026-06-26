import sys
sys.stdout.reconfigure(encoding='utf-8')
from pypdf import PdfReader
import glob

pdfs = glob.glob('Files/*69144400626350*.pdf') + glob.glob('Files/*KHAING*.pdf')
if not pdfs:
    pdfs = glob.glob('Files/*.pdf')[:3]
print('Found:', pdfs)
for p in pdfs[:1]:
    print('=== ', p, ' ===')
    r = PdfReader(p)
    for i, pg in enumerate(r.pages):
        print(f'--- PAGE {i+1} ---')
        print(pg.extract_text())
        print()
