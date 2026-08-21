# ── test_validate.py ── Unit checks for the validation rules ──────────────
#
# These run in milliseconds (no PDFs involved) and pin down the rules that
# matter: the two failures that went unnoticed for weeks in production must
# each be caught, and a healthy week must stay silent.
#
#   python test_validate.py

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

from validate import (FAIL, WARN, looks_like_table_text, validate_project,
                      validate_run, project_history, _frozen_streak)

_failures = []


def check(name, condition, detail=''):
    if condition:
        print(f'  ok   {name}')
    else:
        print(f'  FAIL {name}' + (f' — {detail}' if detail else ''))
        _failures.append(name)


def rules(findings):
    return {f['rule'] for f in findings}


def sev(findings, rule):
    return next((f['severity'] for f in findings if f['rule'] == rule), None)


# ── Fixtures modelled on real extraction output ───────────────────────────

HEALTHY_GMTP = {
    'plan': 26.98, 'actual': 30.65,
    'disciplines': {
        'Engineering':   {'plan': 76.81, 'actual': 84.57},
        'Procurement':   {'plan': 33.00, 'actual': 36.01},
        'Construction':  {'plan': 18.44, 'actual': 22.71},
        'Commissioning': {'plan': 0.0,   'actual': 0.0},
    },
    'scopes': {
        'LNG Tank Area':      {'plan': 34.44, 'actual': 39.99},
        'BOP & Utility Area': {'plan': 28.87, 'actual': 32.53},
        'Marine Area':        {'plan': 7.63,  'actual': 8.24},
    },
    'concerns': [
        'ISB Elevator — Application of elevator in instrument / substation '
        'building should be settled considering construction schedule.',
        'Duct Bank for Future GTG Power Line — Confirmation on duct bank for '
        'future GTG power line installation is required.',
    ],
    'activities': [
        '[LNG Tank] Tank 101: Installation DOKA (outer wall 6th) climbing form',
        '[BOP Area] Piling Work: Pile Driving at Process Area',
    ],
}

# What GMTP actually produced from W28 to W31: the overall % parsed, but the
# 3.1.1 table's scrambled text layer yielded no breakdown, and the generic
# scan filled concerns with the SHE NCR table and the milestone list.
BROKEN_GMTP = {
    'plan': 26.05, 'actual': 30.13,
    'disciplines': {}, 'scopes': {},
    'concerns': [
        '8.4 SHE NCR Status',
        'Doc No. Description Issue date Root cause Status',
        '14 PRO [Elec.] Submission of revised TBE for MV, LV SWGR & MCC',
    ] + [f'filler concern line number {i} with enough length' for i in range(17)],
    'activities': [f'activity line number {i} with enough length'
                   for i in range(20)],
}

# Pak Lay for W31/W32 before the OCR patterns were loosened.
BROKEN_PAKLAY = {
    'plan': None, 'actual': None,
    'disciplines': {}, 'scopes': {},
    'concerns': ['a real concern sentence about the spillway works'],
    'activities': ['a real activity sentence about concrete pouring'],
}

PREV_PAKLAY = {'plan': 4.71, 'actual': 6.83, 'n_disc': 0, 'n_scopes': 0,
               'week_key': 'PRJ-025_W30_2026'}


print('looks_like_table_text')
check('numbered section heading', looks_like_table_text('8.4 SHE NCR Status'))
check('column header row',
      looks_like_table_text('Doc No. Description Issue date Root cause Status'))
check('milestone row',
      looks_like_table_text('14 PRO [Elec.] Submission of revised TBE for MV'))
check('numeric row', looks_like_table_text('Plan 0 29 0 20'))
check('real prose passes', not looks_like_table_text(
    'Confirmation on duct bank for future GTG power line installation is '
    'required.'))
check('area-tagged activity passes', not looks_like_table_text(
    '[BOP Area] Piling Work: Pile Driving at Process Area'))

print('\nhealthy GMTP week')
f = validate_project('PRJ-001', HEALTHY_GMTP)
check('no findings', not f, str(rules(f)))

print('\nGMTP with no breakdown (the W28-W31 regression)')
f = validate_project('PRJ-001', BROKEN_GMTP)
check('missing-disciplines is FAIL', sev(f, 'missing-disciplines') == FAIL)
check('missing-scopes is FAIL', sev(f, 'missing-scopes') == FAIL)
check('content-truncated warned', sev(f, 'content-truncated') == WARN)
check('content-not-prose warned', sev(f, 'content-not-prose') == WARN)

print('\nPak Lay with no percentages (the W31-W32 regression)')
f = validate_project('PRJ-025', BROKEN_PAKLAY)
check('missing-pct is FAIL', sev(f, 'missing-pct') == FAIL)
check('two missing-pct findings',
      sum(1 for x in f if x['rule'] == 'missing-pct') == 2)

print('\nPak Lay against last week — the strongest signal')
f = validate_project('PRJ-025', BROKEN_PAKLAY, prev=PREV_PAKLAY)
check('value-vanished is FAIL', sev(f, 'value-vanished') == FAIL)
check('names the previous week',
      any('W30' in x['message'] for x in f if x['rule'] == 'value-vanished'))

print('\ncumulative actual going backwards')
f = validate_project('PRJ-025', {'plan': 5.0, 'actual': 5.90,
                                 'concerns': ['x' * 40], 'activities': []},
                     prev=PREV_PAKLAY)
check('actual-went-backwards warned',
      sev(f, 'actual-went-backwards') == WARN)
f = validate_project('PRJ-025', {'plan': 5.0, 'actual': 7.10,
                                 'concerns': ['x' * 40], 'activities': []},
                     prev=PREV_PAKLAY)
check('normal progress stays quiet',
      'actual-went-backwards' not in rules(f))

print('\nstructure shrinking')
f = validate_project('PRJ-013', {
    'plan': 50.0, 'actual': 45.0,
    'disciplines': {d: {'plan': 1.0, 'actual': 1.0} for d in
                    ('Engineering', 'Procurement', 'Construction',
                     'Commissioning')},
    'concerns': ['x' * 40], 'activities': ['y' * 40],
}, prev={'plan': 49.0, 'actual': 44.0, 'n_disc': 4, 'n_scopes': 2})
check('scope count drop warned', sev(f, 'structure-shrank') == WARN)

print('\nfrozen value')
rows = [{'actual': 6.83}, {'actual': 6.83}, {'actual': 6.83}]
check('streak counts prior weeks only', _frozen_streak(rows) == 2)
f = validate_project('PRJ-025', {'plan': 5.0, 'actual': 6.83,
                                 'concerns': ['x' * 40], 'activities': []},
                     prev={'plan': 4.9, 'actual': 6.83}, prev_streak=2)
check('value-frozen warned', sev(f, 'value-frozen') == WARN)

print('\nscope missing a percentage')
f = validate_project('PRJ-010', {
    'plan': 46.75, 'actual': 50.43,
    'scopes': {'CBOP': {'plan': 46.75, 'actual': 50.43},
               'TSA':  {'plan': None,  'actual': 25.3}},
    'concerns': ['x' * 40], 'activities': ['y' * 40],
})
check('scope-incomplete warned', sev(f, 'scope-incomplete') == WARN)
check('names the scope',
      any('TSA' in x['message'] for x in f if x['rule'] == 'scope-incomplete'))

print('\nwhole-run report')
results = {
    'PRJ-001': {'found': True,  'data': HEALTHY_GMTP},
    'PRJ-025': {'found': True,  'data': BROKEN_PAKLAY},
    'PRJ-008': {'found': False, 'data': {}},
}
history = {
    'PRJ-025_W30_2026': {'plan': 4.71, 'actual': 6.83, 'n_disc': 0,
                         'n_scopes': 0},
    # Must be ignored as a baseline for its own week.
    'PRJ-025_W32_2026': {'plan': 99.0, 'actual': 99.0, 'n_disc': 0,
                         'n_scopes': 0},
}
rep = validate_run(results, 32, 2026, history=history)
check('missing report not validated', 'PRJ-008' not in rep['by_project'])
check('healthy project not flagged', 'PRJ-001' not in rep['by_project'])
check('broken project flagged', 'PRJ-025' in rep['by_project'])
check('FAIL counted', rep['counts'][FAIL] >= 1)
check('current week excluded from its own history',
      len(project_history(history, 'PRJ-025', before=(2026, 32))) == 1)

print()
if _failures:
    print(f'{len(_failures)} check(s) failed: {_failures}')
    sys.exit(1)
print('All validation checks passed.')
