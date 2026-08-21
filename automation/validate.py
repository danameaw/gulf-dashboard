# ── validate.py ── Post-extraction quality gate ────────────────────────────
#
# Why this exists: run.py used to print "OK" for every project it found a PDF
# for, regardless of what came out of it. Pak Lay reported plan=null/actual=null
# for two consecutive weeks and GMTP reported discs=0 for five, both under an
# "OK" line, so nobody knew until someone opened the dashboard and noticed.
#
# This module takes the extraction results and answers one question per
# project: does this look like a successful read? It knows two kinds of thing:
#
#   1. What each project type is SUPPOSED to yield (an iWTE report always has
#      four disciplines; a GMTP report always has four disciplines and three
#      areas; Solar/Wind always have at least two scopes).
#   2. What the project reported LAST week — the single best oracle for a
#      silently broken parser. A value that existed last week and is null now
#      is a parser regression, not news from site.
#
# Pure data in, findings out: no PDF reading, no file writing, so it can be
# unit-tested and called from run.py, test_extract.py or an audit script.

import json
import os
import re

from config import PROJECT_NAMES
from patterns import AUTO_EXTRACT_PROJECTS

HISTORY_PATH = os.path.join(os.path.dirname(__file__), 'history.json')

# extract_from_pdf truncates both lists at 20. Hitting exactly the cap almost
# never means "there really were 20 concerns" - it means a generic text scan
# matched something table-shaped and got cut off mid-stream.
CONTENT_CAP = 20

FAIL = 'FAIL'
WARN = 'WARN'

# What a successful read looks like per extraction type.
#   min_disc   - project-wide EPCC breakdown entries expected
#   min_scopes - separately-reported scopes/areas expected
#   needs_pct  - project must report an overall plan and actual
EXPECTATIONS = {
    'gmtp':          {'min_disc': 4, 'min_scopes': 3, 'needs_pct': True},
    'iwte':          {'min_disc': 4, 'min_scopes': 0, 'needs_pct': True},
    'solar':         {'min_disc': 0, 'min_scopes': 2, 'needs_pct': True},
    'wind':          {'min_disc': 0, 'min_scopes': 2, 'needs_pct': True},
    'hydro_pakbeng': {'min_disc': 0, 'min_scopes': 0, 'needs_pct': True},
    'hydro_paklay':  {'min_disc': 0, 'min_scopes': 0, 'needs_pct': True},
}

# A concern line that is really a table header or a row of column labels.
# Two or more of these words in one short line is the giveaway - real concern
# text is prose ("Confirmation on duct bank ... is required").
_HEADER_WORDS = re.compile(
    r'\b(description|remark|status|root cause|doc no|issue date|plan|actual|'
    r'total|accum|discipline|week|target|action by)\b', re.I)
# "8.4 SHE NCR Status" - a numbered section heading swept up as a concern.
_SECTION_HEADING = re.compile(r'^\d+(?:\.\d+)+\s+\S')
# "14 PRO [Elec.] Submission of ..." - a milestone table row.
_MILESTONE_ROW = re.compile(r'^\d+\s+[A-Z]{3}\s*\[')
# A line that is nothing but numbers and punctuation is a table row.
_NUMERIC_ROW = re.compile(r'^[\W\d\s%.,\-]+$')
# ...as is one where the numbers outnumber the words ("Plan 0 29 0 20",
# "General 35 35 34" - both real GMTP misreads). Report prose almost never
# runs three bare figures together; a dated activity line ("Pouring Concrete
# for outer wall 6th (19.Aug)") has none, since its tokens carry letters.
_NUM_TOKEN = re.compile(r'^[\d.,%()+\-]+$')


def _numeric_dominant(s):
    toks = s.split()
    if len(toks) < 4:
        return False
    nums = sum(1 for t in toks if _NUM_TOKEN.match(t))
    return nums >= 3 and nums * 2 >= len(toks)

# How far actual may fall below last week's before it looks like a misread
# rather than a re-baseline.
REGRESSION_TOLERANCE = 0.5
# Consecutive weeks of an unchanged actual before it reads as stale.
FROZEN_WEEKS = 3
# Variance beyond this is worth a look even when everything parsed.
VARIANCE_ALERT = 25.0


# ── Findings ──────────────────────────────────────────────────────────────

def _finding(prj_id, severity, rule, message, detail=None):
    return {
        'project':  prj_id,
        'name':     PROJECT_NAMES.get(prj_id, prj_id),
        'severity': severity,
        'rule':     rule,
        'message':  message,
        'detail':   detail,
    }


def _pct_ok(v):
    return isinstance(v, (int, float)) and 0.0 <= v <= 100.0


def looks_like_table_text(line):
    """
    True when a concern/activity line is a table header, a numbered section
    heading, a milestone row or a bare row of numbers rather than prose.
    """
    s = (line or '').strip()
    if not s:
        return True
    if _SECTION_HEADING.match(s) or _MILESTONE_ROW.match(s):
        return True
    if _NUMERIC_ROW.match(s) or _numeric_dominant(s):
        return True
    # Column-label words. Three or more is conclusive ("Plan Actual Diff.
    # Plan Actual Diff. Acc.Plan"). Two is only conclusive in a line too
    # short to be a sentence - "Area Work Description Target Date Problem" is
    # a header row, while "Plan to pouring concrete for WTG-03 next week."
    # also has two and is perfectly good report content.
    labels = len(_HEADER_WORDS.findall(s))
    if labels >= 3:
        return True
    if labels >= 2 and len(s.split()) <= 6:
        return True
    return False


# ── Per-project checks ────────────────────────────────────────────────────

def validate_project(prj_id, data, prev=None, prev_streak=0):
    """
    Check one project's extraction result.
      data        - the dict from extract_from_pdf (or a manual override)
      prev        - the same project's history entry for the previous week
                    it reported, or None
      prev_streak - how many consecutive prior weeks reported the same actual
    Returns a list of finding dicts, most serious first.
    """
    out = []
    kind = AUTO_EXTRACT_PROJECTS.get(prj_id)
    exp = EXPECTATIONS.get(kind, {'min_disc': 0, 'min_scopes': 0,
                                  'needs_pct': False})

    plan   = data.get('plan')
    actual = data.get('actual')
    discs  = data.get('disciplines') or {}
    scopes = data.get('scopes') or {}
    concerns   = data.get('concerns') or []
    activities = data.get('activities') or []

    # ── Overall percentages ───────────────────────────────────────────────
    if exp['needs_pct']:
        for label, v in (('plan', plan), ('actual', actual)):
            if v is None:
                out.append(_finding(
                    prj_id, FAIL, 'missing-pct',
                    f'overall {label} not found in the report'))
            elif not _pct_ok(v):
                out.append(_finding(
                    prj_id, FAIL, 'pct-out-of-range',
                    f'overall {label} = {v} is outside 0-100'))

    # ── Structure ─────────────────────────────────────────────────────────
    if len(discs) < exp['min_disc']:
        out.append(_finding(
            prj_id, FAIL, 'missing-disciplines',
            f'{len(discs)} of {exp["min_disc"]} expected disciplines '
            f'({kind} reports always break EPCC down)'))
    if len(scopes) < exp['min_scopes']:
        out.append(_finding(
            prj_id, FAIL, 'missing-scopes',
            f'{len(scopes)} of {exp["min_scopes"]} expected scopes'))

    for name, sc in scopes.items():
        if sc.get('plan') is None or sc.get('actual') is None:
            out.append(_finding(
                prj_id, WARN, 'scope-incomplete',
                f'scope "{name}" is missing plan or actual',
                f"plan={sc.get('plan')}, actual={sc.get('actual')}"))

    # ── Content quality ───────────────────────────────────────────────────
    # One finding per list. Junk text is the actionable half, so it wins the
    # rule name when both apply; a bare cap hit still gets reported because a
    # list cut at exactly the limit may have dropped real items.
    for label, items in (('concerns', concerns), ('activities', activities)):
        junk = [it for it in items if looks_like_table_text(it)]
        capped = len(items) >= CONTENT_CAP
        if junk:
            note = (f' and the list is capped at {CONTENT_CAP}, so real items '
                    f'may also be missing') if capped else ''
            out.append(_finding(
                prj_id, WARN, 'content-not-prose',
                f'{len(junk)} of {len(items)} {label} are table text rather '
                f'than report content{note}',
                junk[0][:120]))
        elif capped:
            out.append(_finding(
                prj_id, WARN, 'content-truncated',
                f'{label} stopped at the {CONTENT_CAP}-item cap - anything '
                f'past that was dropped'))

    if not concerns and not activities:
        out.append(_finding(
            prj_id, WARN, 'content-empty',
            'no concerns and no activities extracted'))

    # ── Sanity ────────────────────────────────────────────────────────────
    if _pct_ok(plan) and _pct_ok(actual):
        if abs(actual - plan) > VARIANCE_ALERT:
            out.append(_finding(
                prj_id, WARN, 'variance-large',
                f'variance {actual - plan:+.2f}% exceeds '
                f'{VARIANCE_ALERT:.0f}% - confirm this is real'))

    # ── Week-over-week ────────────────────────────────────────────────────
    if prev:
        for label, now, before in (('plan', plan, prev.get('plan')),
                                   ('actual', actual, prev.get('actual'))):
            if now is None and before is not None:
                when = prev.get('week_key', 'the previous week')
                out.append(_finding(
                    prj_id, FAIL, 'value-vanished',
                    f'{label} was {before}% in {when} and is missing now '
                    f'- the pattern has most likely stopped matching'))

        if _pct_ok(actual) and _pct_ok(prev.get('actual')):
            drop = prev['actual'] - actual
            if drop > REGRESSION_TOLERANCE:
                out.append(_finding(
                    prj_id, WARN, 'actual-went-backwards',
                    f'cumulative actual fell {drop:.2f}% '
                    f'({prev["actual"]}% -> {actual}%) - a misread, or a '
                    f're-baseline worth noting'))

        for label, key, now_n in (('disciplines', 'n_disc', len(discs)),
                                  ('scopes', 'n_scopes', len(scopes))):
            before_n = prev.get(key) or 0
            if now_n < before_n:
                out.append(_finding(
                    prj_id, WARN, 'structure-shrank',
                    f'{now_n} {label} now vs {before_n} last time'))

        if (prev_streak + 1) >= FROZEN_WEEKS and _pct_ok(actual) \
                and actual == prev.get('actual'):
            out.append(_finding(
                prj_id, WARN, 'value-frozen',
                f'actual has been {actual}% for {prev_streak + 1} weeks '
                f'running - stale report, or the same figure being re-read'))

    order = {FAIL: 0, WARN: 1}
    return sorted(out, key=lambda f: order.get(f['severity'], 9))


# ── Whole-run validation ──────────────────────────────────────────────────

def summarize(prj_id, data):
    """The subset of an extraction result worth keeping as history."""
    return {
        'plan':         data.get('plan'),
        'actual':       data.get('actual'),
        'n_disc':       len(data.get('disciplines') or {}),
        'n_scopes':     len(data.get('scopes') or {}),
        'n_concerns':   len(data.get('concerns') or []),
        'n_activities': len(data.get('activities') or []),
    }


def _week_sort_key(week_key):
    """'PRJ-001_W32_2026' -> (2026, 32) for chronological ordering."""
    m = re.search(r'_W(\d+)_(\d+)$', week_key)
    return (int(m.group(2)), int(m.group(1))) if m else (0, 0)


def project_history(history, prj_id, before=None):
    """
    Every recorded week for one project, oldest first. `before` is a
    (year, week) tuple; when given, only earlier weeks are returned.
    """
    rows = []
    for key, entry in history.items():
        if not key.startswith(prj_id + '_W'):
            continue
        sk = _week_sort_key(key)
        if before and sk >= before:
            continue
        rows.append((sk, key, entry))
    rows.sort()
    return [dict(entry, week_key=key) for _, key, entry in rows]


def _frozen_streak(rows):
    """
    How many consecutive most-recent weeks share the same actual, not
    counting the latest one itself (so 0 means "the latest value is new").
    """
    if len(rows) < 2:
        return 0
    latest = rows[-1].get('actual')
    if latest is None:
        return 0
    streak = 0
    for row in reversed(rows[:-1]):
        if row.get('actual') == latest:
            streak += 1
        else:
            break
    return streak


def validate_run(results, week, year, history=None):
    """
    Validate a whole weekly run.
      results - {prj_id: {'found': bool, 'data': {...}}} as run.py builds it
      history - loaded history.json, or None to skip week-over-week checks
    Returns {'findings': [...], 'counts': {...}, 'by_project': {...}}.
    """
    history = history or {}
    findings = []

    for prj_id, res in results.items():
        if not res.get('found'):
            continue          # a missing report is reported separately
        rows = project_history(history, prj_id, before=(year, week))
        prev = rows[-1] if rows else None
        streak = _frozen_streak(rows) if rows else 0
        findings.extend(validate_project(prj_id, res.get('data') or {},
                                         prev=prev, prev_streak=streak))

    by_project = {}
    for f in findings:
        by_project.setdefault(f['project'], []).append(f)

    counts = {
        FAIL: sum(1 for f in findings if f['severity'] == FAIL),
        WARN: sum(1 for f in findings if f['severity'] == WARN),
        'projects_flagged': len(by_project),
    }
    return {'findings': findings, 'counts': counts, 'by_project': by_project}


# ── History file ──────────────────────────────────────────────────────────

def load_history(path=HISTORY_PATH):
    try:
        with open(path, encoding='utf-8') as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def save_history(history, path=HISTORY_PATH):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(history, f, indent=1, sort_keys=True)


def record_run(history, results, week, year):
    """Fold this run's results into the history dict (mutates and returns)."""
    for prj_id, res in results.items():
        if not res.get('found'):
            continue
        history[f'{prj_id}_W{week}_{year}'] = summarize(prj_id,
                                                        res.get('data') or {})
    return history


# ── Printing ──────────────────────────────────────────────────────────────

def format_findings(report, indent='  '):
    """Plain-text block for the console and the run log."""
    findings = report['findings']
    if not findings:
        return f'{indent}✓ Validation: no issues found.'

    lines = []
    c = report['counts']
    lines.append(f'{indent}Validation: {c[FAIL]} FAIL, {c[WARN]} WARN '
                 f'across {c["projects_flagged"]} project(s)')
    for prj_id, items in report['by_project'].items():
        name = items[0]['name']
        lines.append(f'{indent}[{prj_id}] {name}')
        for f in items:
            lines.append(f'{indent}   {f["severity"]:4s} {f["rule"]}: '
                         f'{f["message"]}')
            if f['detail']:
                lines.append(f'{indent}        → {f["detail"]}')
    return '\n'.join(lines)


def findings_html(report):
    """Findings table for the weekly email; '' when there is nothing to say."""
    findings = report['findings']
    if not findings:
        return (
            "<p style='margin:0 0 24px;color:#059669'>"
            "Automated data checks passed with no issues.</p>")

    rows = ''.join(
        f"<tr style='background:{'#fef2f2' if f['severity'] == FAIL else '#fffbeb'};"
        f"border-bottom:1px solid #e5e7eb'>"
        f"<td style='padding:6px 12px;color:"
        f"{'#b91c1c' if f['severity'] == FAIL else '#b45309'};"
        f"font-weight:600'>{f['severity']}</td>"
        f"<td style='padding:6px 12px;color:#374151'>{f['project']}</td>"
        f"<td style='padding:6px 12px;color:#374151'>{f['message']}</td></tr>"
        for f in findings
    )
    c = report['counts']
    return f"""
<p style="font-size:12px;font-weight:600;letter-spacing:0.06em;color:#6b7280;margin:0 0 8px">DATA QUALITY CHECKS</p>
<p style="margin:0 0 6px">The automated checks raised <b>{c[FAIL]} error(s)</b> and
<b>{c[WARN]} warning(s)</b> on {c['projects_flagged']} project(s). Figures for
those projects should be confirmed against the source report before use.</p>
<table cellspacing="0" cellpadding="0" style="width:100%;border-collapse:collapse;margin-bottom:24px;border:1px solid #e5e7eb;border-radius:6px;overflow:hidden">
  <tr style="background:#1f3864">
    <th style="padding:8px 12px;text-align:left;color:#fff;font-weight:500;font-size:12px;width:60px">Level</th>
    <th style="padding:8px 12px;text-align:left;color:#fff;font-weight:500;font-size:12px;width:90px">Project</th>
    <th style="padding:8px 12px;text-align:left;color:#fff;font-weight:500;font-size:12px">Finding</th>
  </tr>
  {rows}
</table>"""


# ── Bootstrapping history from index.html ─────────────────────────────────
#
# The dashboard has carried every week's seed data since W24; that is the only
# record of what past runs extracted, so the first history.json is built from
# it rather than by re-running extraction over months of PDFs.

def _balanced(text, start):
    """
    Given text[start] as an opening '{' or '[', return (inner, end) where
    inner is the text between the delimiters and end indexes just past the
    closing one. Skips over single-quoted JS strings and their escapes.
    """
    pairs = {'{': '}', '[': ']'}
    open_ch = text[start]
    close_ch = pairs[open_ch]
    depth = 0
    i = start
    while i < len(text):
        ch = text[i]
        if ch == '\\':
            i += 2
            continue
        if ch == "'":
            i += 1
            while i < len(text):
                if text[i] == '\\':
                    i += 2
                    continue
                if text[i] == "'":
                    break
                i += 1
        elif ch == open_ch:
            depth += 1
        elif ch == close_ch:
            depth -= 1
            if depth == 0:
                return text[start + 1:i], i + 1
        i += 1
    return '', start + 1


def _top_level_items(inner):
    """Split a balanced body on its top-level commas."""
    items, depth, buf, i = [], 0, [], 0
    while i < len(inner):
        ch = inner[i]
        if ch == '\\':
            buf.append(inner[i:i + 2])
            i += 2
            continue
        if ch == "'":
            j = i + 1
            while j < len(inner):
                if inner[j] == '\\':
                    j += 2
                    continue
                if inner[j] == "'":
                    break
                j += 1
            buf.append(inner[i:j + 1])
            i = j + 1
            continue
        if ch in '{[':
            depth += 1
        elif ch in '}]':
            depth -= 1
        if ch == ',' and depth == 0:
            items.append(''.join(buf))
            buf = []
        else:
            buf.append(ch)
        i += 1
    if ''.join(buf).strip():
        items.append(''.join(buf))
    return [it.strip() for it in items if it.strip()]


def _top_level_value(body, key, open_ch):
    """
    The balanced body of `key: <open_ch> ... >` at the TOP level of a seed.
    Searching the whole text with a plain regex would find a nested key first -
    a Solar seed's `scopes` entries each carry their own `disciplines`, so a
    naive search reported those as if the project had a project-wide
    breakdown. Returns None when the key is not present at the top level.
    """
    for item in _top_level_items(body):
        m = re.match(re.escape(key) + r'\s*:\s*' + re.escape(open_ch), item)
        if m:
            inner, _ = _balanced(item, m.end() - 1)
            return inner
    return None


def _count_block(body, key):
    """Number of entries in the top-level `key: { ... }` of a seed body."""
    inner = _top_level_value(body, key, '{')
    return len(_top_level_items(inner)) if inner is not None else 0


def _count_list(body, key):
    """Number of items in the top-level `key: [ ... ]` of a seed body."""
    inner = _top_level_value(body, key, '[')
    return len(_top_level_items(inner)) if inner is not None else 0


def build_history_from_html(html_path):
    """
    Read every seedIfEmpty() block out of index.html into a history dict.
    Blocks are injected newest-first and seedIfEmpty ignores a key that is
    already set, so for a duplicated key the FIRST block in the file is the
    one the dashboard actually uses - and the one recorded here.
    """
    with open(html_path, encoding='utf-8') as f:
        html = f.read()

    history = {}
    for m in re.finditer(r"seedIfEmpty\(\s*'([^']+)'\s*,\s*\{", html):
        key = m.group(1)
        if key in history:
            continue
        body, _ = _balanced(html, m.end() - 1)
        plan   = re.search(r'\bplan\s*:\s*(null|[\d.]+)', body)
        actual = re.search(r'\bactual\s*:\s*(null|[\d.]+)', body)

        def num(mm):
            if not mm or mm.group(1) == 'null':
                return None
            try:
                return float(mm.group(1))
            except ValueError:
                return None

        history[key] = {
            'plan':         num(plan),
            'actual':       num(actual),
            'n_disc':       _count_block(body, 'disciplines'),
            'n_scopes':     _count_block(body, 'scopes'),
            'n_concerns':   _count_list(body, 'concerns'),
            'n_activities': _count_list(body, 'activities'),
        }
    return history


if __name__ == '__main__':
    import argparse
    import sys
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

    from config import DASHBOARD_HTML

    ap = argparse.ArgumentParser(
        description='Validation rules and history for the weekly extraction.')
    ap.add_argument('--bootstrap', action='store_true',
                    help='(re)build history.json from index.html seed data')
    ap.add_argument('--show', metavar='PRJ-XXX',
                    help='print the recorded history for one project')
    args = ap.parse_args()

    if args.bootstrap:
        hist = build_history_from_html(DASHBOARD_HTML)
        save_history(hist)
        weeks = sorted({_week_sort_key(k) for k in hist})
        print(f'history.json: {len(hist)} entries across {len(weeks)} weeks '
              f'({weeks[0][1]}/{weeks[0][0]} - {weeks[-1][1]}/{weeks[-1][0]})')
    elif args.show:
        for row in project_history(load_history(), args.show):
            print(f"  {row['week_key']:22s} plan={row['plan']} "
                  f"actual={row['actual']} disc={row['n_disc']} "
                  f"scopes={row['n_scopes']} concerns={row['n_concerns']} "
                  f"activities={row['n_activities']}")
    else:
        ap.print_help()
