# ── test_fixtures.py ── Golden-file regression gate ───────────────────────
#
# Run this BEFORE committing any change to extract.py or patterns.py.
#
#   python test_fixtures.py --week 32            compare against the fixture
#   python test_fixtures.py --week 32 --update   record a new fixture
#   python test_fixtures.py --all                compare every fixture there is
#
# The problem it solves: patterns are shared across 27 projects and each new
# report wording has historically been handled by adding another regex. There
# was no way to tell whether a fix for one project quietly broke another - so
# the same mistakes kept coming back. A fixture pins every number a known-good
# week produced; if a pattern change moves any of them, the diff says exactly
# which project and which field.
#
# Fixtures store the figures (which must never change silently) plus the
# shape and first line of the text lists (which are allowed to be reworded,
# but not to revert to table junk).

import argparse
import glob
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

from config import (BASE_REPORT_PATH, FOLDER_PROJECTS, MANUAL_OVERRIDES,
                    PROJECT_NAMES)
from extract import extract_from_pdf
from run import find_pdf, resolve_group_folder

FIXTURE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           'fixtures')


def week_folder(week):
    matches = sorted(glob.glob(os.path.join(BASE_REPORT_PATH, f'{week:02d}_*')))
    return matches[0] if matches else None


def snapshot(data):
    """The part of an extraction result a fixture pins down."""
    def discs(d):
        return {k: [v.get('plan'), v.get('actual'), v.get('wf')]
                for k, v in sorted((d or {}).items())}

    return {
        'plan':          data.get('plan'),
        'actual':        data.get('actual'),
        'disciplines':   discs(data.get('disciplines')),
        'scopes': {
            name: {
                'plan':        sc.get('plan'),
                'actual':      sc.get('actual'),
                'disciplines': discs(sc.get('disciplines')),
            }
            for name, sc in sorted((data.get('scopes') or {}).items())
        },
        'n_concerns':    len(data.get('concerns') or []),
        'n_activities':  len(data.get('activities') or []),
        'first_concern':  (data.get('concerns') or [None])[0],
        'first_activity': (data.get('activities') or [None])[0],
    }


def extract_week(week, only=None):
    """Re-extract a whole week. Returns {prj_id: snapshot}."""
    folder = week_folder(week)
    if not folder:
        print(f'  no report folder for week {week}')
        return {}
    print(f'  folder: {folder}')

    out = {}
    for folder_name, prj_ids in FOLDER_PROJECTS.items():
        group = resolve_group_folder(folder, folder_name)
        if not group:
            continue
        for prj_id in prj_ids:
            if only and prj_id not in only:
                continue
            pdf = find_pdf(group, prj_id)
            if not pdf:
                continue
            if prj_id in MANUAL_OVERRIDES:
                data = MANUAL_OVERRIDES[prj_id]
            else:
                data = extract_from_pdf(pdf, prj_id, search_dir=group)
            out[prj_id] = snapshot(data)
            print(f'    {prj_id} {PROJECT_NAMES.get(prj_id, "")}'.ljust(34)
                  + f"plan={out[prj_id]['plan']} actual={out[prj_id]['actual']}")
    return out


def _diff(path, expected, actual, diffs):
    """Recursive compare that reports a dotted path per mismatch."""
    if isinstance(expected, dict) and isinstance(actual, dict):
        for key in sorted(set(expected) | set(actual)):
            if key not in expected:
                diffs.append((f'{path}.{key}', '(absent)', actual[key]))
            elif key not in actual:
                diffs.append((f'{path}.{key}', expected[key], '(absent)'))
            else:
                _diff(f'{path}.{key}', expected[key], actual[key], diffs)
    elif expected != actual:
        diffs.append((path, expected, actual))


def compare(week, fixture, current):
    exp, cur = fixture['projects'], current
    diffs = []

    for prj_id in sorted(set(exp) | set(cur)):
        if prj_id not in cur:
            diffs.append((f'{prj_id}', 'extracted', 'NOT EXTRACTED'))
        elif prj_id not in exp:
            print(f'  note: {prj_id} is new since the fixture was recorded')
        else:
            _diff(prj_id, exp[prj_id], cur[prj_id], diffs)

    if not diffs:
        print(f'\n  ✓ W{week}: {len(cur)} projects match the fixture exactly.')
        return True

    print(f'\n  ✗ W{week}: {len(diffs)} difference(s) vs the fixture')
    for path, before, after in diffs:
        print(f'      {path}')
        print(f'        fixture: {before!r}')
        print(f'        now:     {after!r}')
    return False


def fixture_path(week, year):
    return os.path.join(FIXTURE_DIR, f'W{week}_{year}.json')


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--week', type=int, action='append', default=[],
                    help='week number to check (repeatable)')
    ap.add_argument('--year', type=int, default=2026)
    ap.add_argument('--all', action='store_true',
                    help='check every recorded fixture')
    ap.add_argument('--update', action='store_true',
                    help='overwrite the fixture with the current output')
    ap.add_argument('--only', action='append', default=[],
                    help='limit to these project ids (repeatable)')
    ap.add_argument('--from-dump', metavar='PATH', default=None,
                    help='use a results JSON written by "run.py --dump" '
                         'instead of extracting again — a full week takes '
                         'several minutes, and one dump can feed both this '
                         'and a validation re-check')
    args = ap.parse_args()

    os.makedirs(FIXTURE_DIR, exist_ok=True)

    weeks = list(args.week)
    if args.all:
        for name in sorted(os.listdir(FIXTURE_DIR)):
            if name.startswith('W') and name.endswith('.json'):
                weeks.append(int(name[1:].split('_')[0]))
    if not weeks:
        ap.error('give --week N or --all')

    ok = True
    for week in sorted(set(weeks)):
        path = fixture_path(week, args.year)
        print(f'\n=== W{week}/{args.year} ===')
        if args.from_dump:
            with open(args.from_dump, encoding='utf-8') as f:
                dumped = json.load(f)
            only = set(args.only) or None
            current = {pid: snapshot(res.get('data') or {})
                       for pid, res in dumped.items()
                       if res.get('found') and (not only or pid in only)}
            print(f'  from dump: {len(current)} projects '
                  f'({os.path.basename(args.from_dump)})')
        else:
            current = extract_week(week, only=set(args.only) or None)
        if not current:
            ok = False
            continue

        if args.update:
            existing = {}
            if args.only and os.path.exists(path):
                with open(path, encoding='utf-8') as f:
                    existing = json.load(f).get('projects', {})
            existing.update(current)
            with open(path, 'w', encoding='utf-8') as f:
                json.dump({'week': week, 'year': args.year,
                           'projects': existing}, f, indent=1, sort_keys=True)
            print(f'\n  recorded {len(existing)} projects → '
                  f'{os.path.relpath(path)}')
            continue

        if not os.path.exists(path):
            print(f'  no fixture at {os.path.relpath(path)} — run with '
                  f'--update to record one')
            ok = False
            continue
        with open(path, encoding='utf-8') as f:
            fixture = json.load(f)
        ok = compare(week, fixture, current) and ok

    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()
