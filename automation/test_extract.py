import sys, glob
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, r'C:\Users\danaya.th\OneDrive - Gulf\Documents\GitHub\gulf-dashboard\automation')
from extract import extract_from_pdf

BASE = r'C:\Users\danaya.th\Gulf\Engineering - Engineering Documents\00 Project Reports\2026'

tests = [
    ('PRJ-002', glob.glob(BASE + r'\25_260624\GSO2026\GOE2-NWT3*.pdf')),
    ('PRJ-009', glob.glob(BASE + r'\24_260617\Wind2027\ALPHA 1*.pdf')),
    ('PRJ-026', [BASE + r'\25_260624\Pak Beng\Pak Beng Hydroelectric Power Project Monthly Report 15_Updated.pdf']),
    ('PRJ-013', glob.glob(BASE + r'\25_260624\iWTE\*GCE*.pdf')),
]

for prj_id, pdfs in tests:
    if not pdfs:
        print(f'{prj_id}: no PDF found')
        continue
    pdf = pdfs[0]
    print(f'=== {prj_id} : {pdf.split(chr(92))[-1]} ===')
    d = extract_from_pdf(pdf, prj_id)
    print(f'  plan={d["plan"]}, actual={d["actual"]}')
    for sc, sv in d.get('scopes', {}).items():
        print(f'  scope [{sc}]: plan={sv["plan"]}, actual={sv["actual"]}')
        for disc, dv in sv.get('disciplines', {}).items():
            print(f'    {disc}: plan={dv["plan"]}, actual={dv["actual"]}')
    for disc, dv in d.get('disciplines', {}).items():
        print(f'  disc [{disc}]: plan={dv["plan"]}, actual={dv["actual"]}')
    print()
