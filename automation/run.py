# ── run.py ── Gulf Dashboard Weekly Automation ────────────────────────────
# Schedule: every Wednesday via Windows Task Scheduler
# Usage: python run.py  (or python run.py --week 25 --year 2026)
import sys, os, re, json, glob, shutil, subprocess, argparse, time
from datetime import datetime
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

sys.path.insert(0, os.path.dirname(__file__))
from config import (BASE_REPORT_PATH, DASHBOARD_HTML, EXCEL_DIR,
                    GIT_REPO, DASHBOARD_URL,
                    FOLDER_PROJECTS, PROJECT_KEYWORDS, PROJECT_NAMES,
                    MANUAL_OVERRIDES, REPORT_CADENCE, DEFAULT_CADENCE,
                    CADENCE_GRACE_WEEKS)
from extract import extract_from_pdf
from validate import (findings_html, format_findings, load_history,
                      project_history,
                      record_run, save_history, validate_run, FAIL)


# ── 1. Find latest week folder ─────────────────────────────────────────────
def _parse_folder_date(name):
    """Folder name is 'WEEK_YYMMDD' (e.g. '28_260715') — return the date
    part as an ISO string ('2026-07-15'), or None if it doesn't parse."""
    parts = name.split('_')
    if len(parts) > 1 and len(parts[1]) >= 6:
        try:
            return (f"{2000 + int(parts[1][:2])}-"
                    f"{int(parts[1][2:4]):02d}-{int(parts[1][4:6]):02d}")
        except ValueError:
            pass
    return None


def find_week_folder(year=None):
    pattern = os.path.join(BASE_REPORT_PATH, r"[0-9]*_[0-9]*")
    folders = sorted(glob.glob(pattern))
    if not folders:
        raise FileNotFoundError(f"No week folders found in {BASE_REPORT_PATH}")
    latest = folders[-1]
    name   = os.path.basename(latest)           # e.g. "25_260701"
    parts  = name.split('_')
    week   = int(parts[0])
    yr     = 2000 + int(parts[1][:2]) if len(parts) > 1 else datetime.now().year
    return latest, week, yr, _parse_folder_date(name)


def _norm_key(name):
    """Reduce a folder name to letters+digits for tolerant comparison."""
    return re.sub(r'[^a-z0-9]', '', name.lower())


def resolve_group_folder(week_folder, folder_name):
    """
    Locate a project group's folder inside the week folder.
    These folders are named by hand and the spelling drifts week to week:
    'iWTE' / 'IWTE' / 'Iwte' (harmless, Windows paths are case-insensitive)
    but also 'Wind2027' vs 'Wind 2027' (as in 33_260819), which an exact
    os.path.join misses — reporting FOLDER NOT FOUND for all five Wind
    projects even though the reports were sitting right there. Compare on
    letters and digits only so case, spacing and punctuation don't matter.
    Returns the real path, or None if no folder matches.
    """
    exact = os.path.join(week_folder, folder_name)
    if os.path.isdir(exact):
        return exact
    want = _norm_key(folder_name)
    try:
        for entry in sorted(os.listdir(week_folder)):
            cand = os.path.join(week_folder, entry)
            if os.path.isdir(cand) and _norm_key(entry) == want:
                return cand
    except OSError:
        pass
    return None


# ── 2. Find PDF for a project ──────────────────────────────────────────────
def find_pdf(folder_path, prj_id):
    keywords = PROJECT_KEYWORDS.get(prj_id, [])
    pdfs = glob.glob(os.path.join(folder_path, '*.pdf'))
    pdfs += glob.glob(os.path.join(folder_path, '**', '*.pdf'), recursive=True)
    matches = [pdf for pdf in pdfs
               if any(kw.upper() in os.path.basename(pdf).upper() for kw in keywords)]
    if not matches:
        return None
    # iWTE's regular weekly construction report is always named
    # "Project_<ID>_Construction_weekly_report_...". A one-off report (e.g.
    # a special "Back energize" report covering multiple projects) can
    # legitimately contain a project's keyword too and was winning just by
    # sorting first — prefer the canonically-named file when one exists.
    canonical = [pdf for pdf in matches
                 if os.path.basename(pdf).upper().startswith('PROJECT_')]
    return canonical[0] if canonical else matches[0]


def is_report_due(history, prj_id, week, year):
    """
    Should this project have a report this week?
    Returns (due, note). A weekly project is always due. A monthly or
    biweekly one is due only once its grace period since the last report it
    actually filed has elapsed, so it is not chased every week in between.
    A project with no history at all is treated as due - better to ask than
    to silently stop expecting a report.
    """
    cadence = REPORT_CADENCE.get(prj_id, DEFAULT_CADENCE)
    if cadence == DEFAULT_CADENCE:
        return True, None

    rows = project_history(history, prj_id, before=(year, week))
    if not rows:
        return True, f'{cadence}, no prior report on record'

    last = rows[-1]['week_key']
    m = re.search(r'_W(\d+)_(\d+)$', last)
    if not m:
        return True, cadence
    last_week, last_year = int(m.group(1)), int(m.group(2))
    elapsed = (year - last_year) * 52 + (week - last_week)
    grace = CADENCE_GRACE_WEEKS.get(cadence, 1)
    if elapsed < grace:
        return False, (f'{cadence}, last reported W{last_week}/{last_year} '
                       f'({elapsed} week(s) ago)')
    return True, (f'{cadence}, {elapsed} week(s) since W{last_week}/{last_year}')


# ── 3. Build JS seed snippet ───────────────────────────────────────────────
def js_escape(s):
    return s.replace('\\', '\\\\').replace("'", "\\'").replace('\n', ' ')

def _js_num(v):
    """Format a number (or null) for JS."""
    return 'null' if v is None else str(round(v, 2))

def _build_disciplines_js(discs, indent='    '):
    """Render disciplines dict as JS object literal."""
    if not discs:
        return '{}'
    parts = []
    for disc, vals in discs.items():
        p = _js_num(vals.get('plan'))
        a = _js_num(vals.get('actual'))
        # Weight factor, where the report states one (GMTP's area tables do).
        wf = vals.get('wf')
        wf_js = f", wf: {_js_num(wf)}" if wf is not None else ''
        parts.append(
            f"{indent}  '{disc}': {{ plan: {p}, actual: {a}{wf_js} }}")
    return '{\n' + ',\n'.join(parts) + f'\n{indent}}}'

def _build_scopes_js(scopes):
    """Render multi-scope dict as JS object literal."""
    if not scopes:
        return '{}'
    parts = []
    for scope_name, sc in scopes.items():
        p = _js_num(sc.get('plan'))
        a = _js_num(sc.get('actual'))
        discs = sc.get('disciplines', {})
        if discs:
            disc_js = _build_disciplines_js(discs, indent='      ')
            parts.append(
                f"    '{js_escape(scope_name)}': {{\n"
                f"      plan: {p}, actual: {a},\n"
                f"      disciplines: {disc_js}\n"
                f"    }}"
            )
        else:
            parts.append(
                f"    '{js_escape(scope_name)}': {{ plan: {p}, actual: {a} }}"
            )
    return '{\n' + ',\n'.join(parts) + '\n  }'

def build_seed(prj_id, week, year, data):
    concerns    = ',\n    '.join(f"'{js_escape(c)}'" for c in data['concerns'])
    activities  = ',\n    '.join(f"'{js_escape(a)}'" for a in data['activities'])
    plan        = _js_num(data.get('plan'))
    actual      = _js_num(data.get('actual'))

    # Scopes (Solar/Wind) take priority over flat disciplines (iWTE/GMTP)
    scopes = data.get('scopes', {})
    discs  = data.get('disciplines', {})

    # A project can have both: GMTP reports per-area progress (LNG Tank /
    # BOP & Utility / Marine) *and* a project-wide EPCC breakdown. Emitting
    # only `scopes` dropped the disciplines from the dashboard entirely.
    structure_line = ''
    if scopes:
        structure_line += f"  scopes: {_build_scopes_js(scopes)},\n"
    if discs:
        structure_line += f"  disciplines: {_build_disciplines_js(discs)},\n"

    return (
        f"// ── {PROJECT_NAMES.get(prj_id, prj_id)} ({prj_id}) Week {week} — auto-extracted\n"
        f"seedIfEmpty('{prj_id}_W{week}_{year}', {{\n"
        f"  plan: {plan}, actual: {actual},\n"
        f"{structure_line}"
        f"  concerns: [{concerns}],\n"
        f"  activities: [{activities}],\n"
        f"}});\n"
    )


# ── 4. Update index.html ───────────────────────────────────────────────────
SEED_MARKER          = '// ── SEED DATA (pre-loaded from PDF reports) ──'
MISSING_MARKER_START = '// ── AUTO: MISSING REPORTS ──'
MISSING_MARKER_END   = '// ── END MISSING REPORTS ──'
DATE_MARKER_START    = '// ── AUTO: REPORT DATES ──'
DATE_MARKER_END      = '// ── END REPORT DATES ──'
WEEK_MARKER_START    = '// ── AUTO: CURRENT WEEK ──'
WEEK_MARKER_END      = '// ── END CURRENT WEEK ──'

def _find_seed_inject_point(html):
    """
    Offset just past seedIfEmpty()'s closing brace, where generated seed
    blocks are injected. Anchored on the closing brace rather than the
    function body so edits inside seedIfEmpty can't silently break the match.
    Returns -1 when the marker or the function can't be found.
    """
    marker_pos = html.find(SEED_MARKER)
    if marker_pos == -1:
        return -1
    fn_pos = html.find('function seedIfEmpty(key, data) {', marker_pos)
    if fn_pos == -1:
        return -1
    close_pos = html.find('\n}\n', fn_pos)
    if close_pos == -1:
        return -1
    return close_pos + len('\n}\n')


def _week_seed_header(week, year):
    return f'// ── Week {week}/{year} — AUTO-GENERATED ──'


def _drop_week_seed_block(html, week_header):
    """
    Remove any existing seed block for one week: from its header up to the
    next week header (or the end of the seed region). Re-running a week used
    to insert a SECOND block for the same keys instead of replacing the
    first; the newest block does win (it is injected at the top, and
    seedIfEmpty skips keys already set), but every stale copy stayed in the
    file forever - index.html is already 1.5MB - and made "what is actually
    seeded for W32" impossible to answer by reading the file. Dropping the
    old block first makes a corrective re-run idempotent.
    """
    while True:
        start = html.find(week_header)
        if start == -1:
            return html
        end = html.find('// ── Week ', start + len(week_header))
        if end == -1:
            # Blocks are injected newest-first, so the week being re-run
            # normally has an older week's header after it. Without one
            # this is the oldest block in the file and the only stop point
            # left is </script> - everything between would go with it.
            # Leave it alone and fall back to the old append behaviour.
            print(f'  [warn] no block after {week_header.strip()} - '
                  'leaving the existing one in place')
            return html
        html = html[:start] + html[end:]


def bump_seed_version(html):
    """
    seedIfEmpty() refuses to overwrite a key that already holds real
    plan/actual numbers in a visitor's browser, so re-extracting a week only
    reaches people who have never opened the dashboard - everyone else keeps
    seeing the old, wrong numbers forever. index.html carries a SEED_VERSION
    force-refresh guard for exactly this case, but nothing ever bumped it: it
    still read '2026-07-22.2' while five weekly runs shipped on top of it.
    Stamp it on every run so a correction always reaches every visitor.

    This drops each visitor's cached progressData wholesale, manual dashboard
    edits included, for any week whose seed block does not carry them.
    """
    m = re.search(r"const SEED_VERSION = '([^']*)';", html)
    if not m:
        print("  [warn] SEED_VERSION not found - visitors may keep stale data")
        return html
    stamp = datetime.now().strftime('%Y-%m-%d.%H%M')
    print(f"  ✓ SEED_VERSION {m.group(1)} → {stamp} (forces client refresh)")
    return html[:m.start()] + f"const SEED_VERSION = '{stamp}';" + html[m.end():]


def update_html(week, year, seeds_js, missing_ids, report_date=None):
    with open(DASHBOARD_HTML, encoding='utf-8') as f:
        html = f.read()

    # Inject seed data right after the seedIfEmpty() function that follows
    # SEED_MARKER — anchored on the function's closing brace, not its body.
    # (A previous version matched a hardcoded copy of the function body; once
    # seedIfEmpty's internals were edited in index.html the match silently
    # stopped working and every run fell back to appending a duplicate block
    # at the end of the file. Because seedIfEmpty only overwrites a key that
    # is still null, those appended duplicates could permanently shadow any
    # later correction for a key that already had a non-null value.)
    week_header = _week_seed_header(week, year)
    html = _drop_week_seed_block(html, week_header)
    inject_point = _find_seed_inject_point(html)
    if inject_point != -1:
        new_seeds = '\n\n' + week_header + '\n' + '\n'.join(seeds_js)
        html = html[:inject_point] + new_seeds + html[inject_point:]
    else:
        print("  [warn] SEED_MARKER not found, appending seeds before </script>")
        html = html.replace('</script>', '\n'.join(seeds_js) + '\n</script>', 1)

    # Update missing reports — merge into MISSING_REPORTS_BY_WEEK dict
    week_key = f"W{week}_{year}"
    by_week_re = re.compile(
        r'(// ── AUTO: MISSING REPORTS ──\n)'
        r'const MISSING_REPORTS_BY_WEEK = (\{.*?\});'
        r'(\n// ── END MISSING REPORTS ──)',
        re.DOTALL)
    m = by_week_re.search(html)
    if m:
        existing = json.loads(m.group(2))
        existing[week_key] = missing_ids
        new_block = (
            f"{MISSING_MARKER_START}\n"
            f"const MISSING_REPORTS_BY_WEEK = {json.dumps(existing, separators=(',', ':'))};\n"
            f"{MISSING_MARKER_END}"
        )
        html = html[:m.start()] + new_block + html[m.end():]
    else:
        new_block = (
            f"{MISSING_MARKER_START}\n"
            f"const MISSING_REPORTS_BY_WEEK = {json.dumps({week_key: missing_ids}, separators=(',', ':'))};\n"
            f"{MISSING_MARKER_END}"
        )
        if MISSING_MARKER_START in html:
            start = html.index(MISSING_MARKER_START)
            end   = html.index(MISSING_MARKER_END) + len(MISSING_MARKER_END)
            html  = html[:start] + new_block + html[end:]
        else:
            html = html.replace('<script>', new_block + '\n<script>', 1)

    # Update real report date — merge into REPORT_DATES_BY_WEEK dict
    if report_date:
        date_re = re.compile(
            r'(// ── AUTO: REPORT DATES ──.*?)'
            r'const REPORT_DATES_BY_WEEK = (\{.*?\});'
            r'(.*?// ── END REPORT DATES ──)',
            re.DOTALL)
        m = date_re.search(html)
        if m:
            existing = json.loads(m.group(2))
            existing[week_key] = report_date
            new_block = (
                m.group(1) +
                f"const REPORT_DATES_BY_WEEK = {json.dumps(existing, separators=(',', ':'))};" +
                m.group(3)
            )
            html = html[:m.start()] + new_block + html[m.end():]
        else:
            print("  [warn] REPORT_DATES_BY_WEEK marker not found, skipping report date")

    # Update current week displayed in dashboard
    week_js = (
        f"{WEEK_MARKER_START}\n"
        f"currentWeek = {week}; currentYear = {year}; saveData();\n"
        f"{WEEK_MARKER_END}"
    )
    if WEEK_MARKER_START in html:
        start = html.index(WEEK_MARKER_START)
        end   = html.index(WEEK_MARKER_END) + len(WEEK_MARKER_END)
        html  = html[:start] + week_js + html[end:]
    else:
        # Replace any existing hardcoded currentWeek line
        import re as _re
        html = _re.sub(
            r'currentWeek\s*=\s*\d+;\s*currentYear\s*=\s*\d+;\s*saveData\(\);',
            week_js, html)

    html = bump_seed_version(html)

    with open(DASHBOARD_HTML, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"  ✓ index.html updated (Week {week}/{year})")


# ── 5. Update / create Excel ───────────────────────────────────────────────
def update_excel(week, year, results):
    try:
        import openpyxl
        from openpyxl.styles import PatternFill, Font, Alignment
    except ImportError:
        print("  [skip] openpyxl not installed")
        return

    xl_path = os.path.join(EXCEL_DIR, f'Gulf_Dashboard_W{week}_{year}.xlsx')

    # Load existing or create new
    if os.path.exists(xl_path):
        wb = openpyxl.load_workbook(xl_path)
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

    sheet_name = f'W{week}'
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    # Header
    headers = ['PRJ ID', 'Project Name', 'PDF Found', 'Concerns', 'Activities']
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor='1F3864')
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for ri, (prj_id, info) in enumerate(sorted(results.items()), 2):
        ws.cell(ri, 1, prj_id)
        ws.cell(ri, 2, PROJECT_NAMES.get(prj_id, prj_id))
        ws.cell(ri, 3, 'YES' if info['found'] else 'NO')
        ws.cell(ri, 4, '\n'.join(info['data']['concerns']))
        ws.cell(ri, 5, '\n'.join(info['data']['activities']))
        for ci in range(1, 6):
            ws.cell(ri, ci).alignment = Alignment(wrap_text=True, vertical='top')
        if not info['found']:
            for ci in range(1, 6):
                ws.cell(ri, ci).fill = PatternFill('solid', fgColor='FFE0E0')

    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 60
    ws.column_dimensions['E'].width = 60

    wb.save(xl_path)
    print(f"  ✓ Excel saved: {xl_path}")


# ── 6. Git commit & push ───────────────────────────────────────────────────
def git_push(week, year):
    try:
        subprocess.run(['git', 'add', '-A'],
                       cwd=GIT_REPO, check=True, capture_output=True)
        subprocess.run(['git', 'commit', '-m',
                        f'Auto: Week {week}/{year} data extracted from PDFs'],
                       cwd=GIT_REPO, check=True, capture_output=True)
        subprocess.run(['git', 'push'],
                       cwd=GIT_REPO, check=True, capture_output=True)
        print("  ✓ Pushed to GitHub")
    except subprocess.CalledProcessError as e:
        print(f"  [git error] {e.stderr.decode(errors='replace') if e.stderr else e}")


# ── 7. Send Email via Outlook (local desktop app) ─────────────────────────
EMAIL_FROM   = "danaya.th@gulf.co.th"
EMAIL_TO     = ["purachet.am@gulf.co.th", "chalong@gulf.co.th"]

def _build_html_body(week, year, found, missing, report=None):
    found_rows = ''.join(
        f"<tr style='border-bottom:1px solid #e5e7eb'>"
        f"<td style='padding:6px 12px;color:#374151'>{pid}</td>"
        f"<td style='padding:6px 12px;color:#374151'>{PROJECT_NAMES.get(pid, pid)}</td></tr>"
        for pid in sorted(found)
    )
    missing_rows = ''.join(
        f"<tr style='background:#fffbeb;border-bottom:1px solid #e5e7eb'>"
        f"<td style='padding:6px 12px;color:#374151'>{pid}</td>"
        f"<td style='padding:6px 12px;color:#374151'>{PROJECT_NAMES.get(pid, pid)}</td></tr>"
        for pid in missing
    ) if missing else (
        "<tr><td colspan='2' style='padding:6px 12px;color:#059669'>All projects reported — none missing.</td></tr>"
    )
    missing_note = (
        f"<p style='margin:0 0 6px'>Please be advised that the following <b>{len(missing)} project(s)</b> "
        f"have not yet submitted their weekly progress reports. "
        f"Kindly ensure the relevant PDF files are uploaded to the ShareDrive at the earliest convenience.</p>"
    ) if missing else ""

    findings_block = findings_html(report) if report else ''

    return f"""
<html><body style="font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#111827;line-height:1.7;max-width:640px;margin:0 auto;padding:24px">
<p style="margin:0 0 4px">Dear P'Tee and P'Hall,</p>
<p style="margin:0 0 20px">
  Please be informed that the <b>Gulf Engineering Dashboard — Week {week}/{year}</b>
  has been successfully updated and is now available for review.
</p>
<table cellspacing="0" cellpadding="0" style="background:#f3f4f6;border-radius:8px;padding:16px 20px;margin-bottom:24px;width:100%">
  <tr><td colspan="2" style="font-size:12px;font-weight:600;letter-spacing:0.06em;color:#6b7280;padding-bottom:10px">SUMMARY</td></tr>
  <tr>
    <td style="padding:2px 0;color:#374151">Projects Updated</td>
    <td style="padding:2px 0;color:#059669;font-weight:600;text-align:right">{len(found)} projects</td>
  </tr>
  <tr>
    <td style="padding:2px 0;color:#374151">Reports Not Yet Received</td>
    <td style="padding:2px 0;color:#d97706;font-weight:600;text-align:right">{len(missing)} projects</td>
  </tr>
</table>
{findings_block}
<p style="font-size:12px;font-weight:600;letter-spacing:0.06em;color:#6b7280;margin:0 0 8px">UPDATED PROJECTS</p>
<table cellspacing="0" cellpadding="0" style="width:100%;border-collapse:collapse;margin-bottom:24px;border:1px solid #e5e7eb;border-radius:6px;overflow:hidden">
  <tr style="background:#1f3864">
    <th style="padding:8px 12px;text-align:left;color:#fff;font-weight:500;font-size:12px;width:100px">Project ID</th>
    <th style="padding:8px 12px;text-align:left;color:#fff;font-weight:500;font-size:12px">Project Name</th>
  </tr>
  {found_rows}
</table>
<p style="font-size:12px;font-weight:600;letter-spacing:0.06em;color:#6b7280;margin:0 0 8px">REPORTS NOT YET RECEIVED</p>
{missing_note}
<table cellspacing="0" cellpadding="0" style="width:100%;border-collapse:collapse;margin-bottom:24px;border:1px solid #e5e7eb;border-radius:6px;overflow:hidden">
  <tr style="background:#92400e">
    <th style="padding:8px 12px;text-align:left;color:#fef3c7;font-weight:500;font-size:12px;width:100px">Project ID</th>
    <th style="padding:8px 12px;text-align:left;color:#fef3c7;font-weight:500;font-size:12px">Project Name</th>
  </tr>
  {missing_rows}
</table>
<p style="margin:0 0 20px">
  The full dashboard is accessible via the link below:<br>
  <a href="{DASHBOARD_URL}" style="color:#1d4ed8">{DASHBOARD_URL}</a>
</p>
<p style="margin:0 0 4px">Best Regards,</p>
<p style="margin:0 0 2px;font-weight:500">Danaya</p>
<p style="margin:0;font-size:11px;color:#9ca3af">Gulf Engineering — Project Management &amp; Control</p>
</body></html>"""

def _ensure_outlook_running(wait_seconds=20):
    """Launch Outlook if it's not already running, and give it time to finish
    starting up. A cold Dispatch() launch from inside the COM call is what
    causes the automation to hang, so we start the process separately first."""
    try:
        running = subprocess.run(
            ['tasklist', '/FI', 'IMAGENAME eq OUTLOOK.EXE'],
            capture_output=True, text=True
        )
        if 'OUTLOOK.EXE' not in running.stdout.upper():
            subprocess.Popen(['outlook.exe'], shell=True)
            time.sleep(wait_seconds)
    except Exception as e:
        print(f"  [email] Could not check/launch Outlook: {e}")


def send_email(week, year, results, missing_ids, xl_path, timeout=90,
               report=None):
    found   = {k: v for k, v in results.items() if v['found']}
    missing = missing_ids

    subject   = (f"[Gulf Dashboard] W{week}/{year} Update — "
                 f"{len(found)} Projects Updated, {len(missing)} Reports Pending")
    html_body = _build_html_body(week, year, found, missing,
                                report=report)

    _ensure_outlook_running()

    payload = {
        'to': EMAIL_TO,
        'subject': subject,
        'html_body': html_body,
        'attachment': os.path.abspath(xl_path) if xl_path and os.path.exists(xl_path) else None,
    }
    payload_path = os.path.join(os.path.dirname(__file__), '_email_payload.json')
    sender_script = os.path.join(os.path.dirname(__file__), '_send_outlook_mail.py')
    with open(payload_path, 'w', encoding='utf-8') as f:
        json.dump(payload, f)

    try:
        # Runs in its own process so a stuck Outlook COM call (e.g. waiting on
        # a dialog) can be killed by the timeout instead of hanging run.py.
        subprocess.run(
            [sys.executable, sender_script, payload_path],
            timeout=timeout, check=True, capture_output=True, text=True,
        )
        print(f"  [OK] Email sent via Outlook to: {', '.join(EMAIL_TO)}")
    except subprocess.TimeoutExpired:
        print(f"  [email error] Outlook did not respond within {timeout}s "
              f"(may be showing a dialog or still starting up) — email not sent")
    except subprocess.CalledProcessError as e:
        print(f"  [email error] {e.stderr.strip() if e.stderr else e}")
    except Exception as e:
        print(f"  [email error] {e}")
    finally:
        if os.path.exists(payload_path):
            os.remove(payload_path)


# ── 8. Windows Notification ────────────────────────────────────────────────
def notify(week, year, found_count, missing_count):
    msg = (f"Week {week}/{year} — {found_count} projects updated"
           + (f", {missing_count} missing reports" if missing_count else ""))
    url = DASHBOARD_URL
    try:
        # Try win10toast first
        from win10toast import ToastNotifier
        ToastNotifier().show_toast(
            "Gulf Dashboard Updated", msg,
            duration=10, threaded=True)
    except Exception:
        pass
    # Also open browser to dashboard
    try:
        import webbrowser
        webbrowser.open(url)
    except Exception:
        pass
    # Fallback: PowerShell balloon
    try:
        ps_cmd = (
            f'Add-Type -AssemblyName System.Windows.Forms;'
            f'$n = New-Object System.Windows.Forms.NotifyIcon;'
            f'$n.Icon = [System.Drawing.SystemIcons]::Information;'
            f'$n.Visible = $true;'
            f'$n.ShowBalloonTip(8000,"Gulf Dashboard Updated","{msg}",'
            f'[System.Windows.Forms.ToolTipIcon]::Info);'
            f'Start-Sleep 9; $n.Dispose()'
        )
        subprocess.Popen(['powershell', '-Command', ps_cmd])
    except Exception:
        pass


# ── MAIN ───────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--week', type=int, default=None)
    parser.add_argument('--year', type=int, default=None)
    parser.add_argument('--dry-run', action='store_true',
                        help='Extract only, do not update files')
    parser.add_argument('--no-email', action='store_true',
                        help='Skip sending email notification')
    parser.add_argument('--dump', metavar='PATH', default=None,
                        help='Write the raw extraction results to a JSON file '
                             '(a full run takes minutes, so this lets the '
                             'validation rules be re-checked without one)')
    parser.add_argument('--strict', action='store_true',
                        help='Stop before updating any file if the data '
                             'checks report a FAIL')
    parser.add_argument('--no-push', action='store_true',
                        help='Update index.html/Excel locally but do not '
                             'commit or push (for reviewing a correction '
                             'before it goes live)')
    args = parser.parse_args()

    print("=" * 60)
    print("Gulf Dashboard Weekly Automation")
    print("=" * 60)

    # Find week folder
    if args.week and args.year:
        # Manual override: find folder matching week number
        pattern = os.path.join(BASE_REPORT_PATH, f"{args.week:02d}_*")
        matches = glob.glob(pattern)
        if not matches:
            print(f"  [error] No folder found for week {args.week}")
            sys.exit(1)
        week_folder = matches[0]
        week, year = args.week, args.year
        report_date = _parse_folder_date(os.path.basename(week_folder))
    else:
        week_folder, week, year, report_date = find_week_folder()

    print(f"  Week: {week}/{year}")
    print(f"  Folder: {week_folder}")
    print()

    results    = {}   # prj_id → {found, data}
    seeds_js   = []
    missing    = []
    not_due    = []

    # Loaded up front: the main loop needs it to tell a late report from one
    # that simply is not due yet.
    history = load_history()

    # Process each project group
    for folder_name, prj_ids in FOLDER_PROJECTS.items():
        group_path = resolve_group_folder(week_folder, folder_name)
        folder_exists = group_path is not None

        for prj_id in prj_ids:
            name = PROJECT_NAMES.get(prj_id, prj_id)
            print(f"  [{prj_id}] {name}", end=' ... ')

            pdf = find_pdf(group_path, prj_id) if folder_exists else None
            if not pdf:
                reason = ('FOLDER NOT FOUND (%s)' % folder_name
                          if not folder_exists else 'PDF NOT FOUND')
                due, note = is_report_due(history, prj_id, week, year)
                suffix = f' — {note}' if note else ''
                print(f"{'' if due else 'NOT DUE — '}{reason}{suffix}")
                results[prj_id] = {'found': False,
                                   'data': {'concerns': [], 'activities': []}}
                (missing if due else not_due).append(prj_id)
                continue

            if prj_id in MANUAL_OVERRIDES:
                print(f"OK → {os.path.basename(pdf)} (manual override)")
                data = MANUAL_OVERRIDES[prj_id]
                results[prj_id] = {'found': True, 'data': data}
                seeds_js.append(build_seed(prj_id, week, year, data))
                continue

            print(f"OK → {os.path.basename(pdf)}")
            data = extract_from_pdf(pdf, prj_id, search_dir=group_path)
            plan_s   = f"{data['plan']}%" if data['plan'] is not None else 'null'
            actual_s = f"{data['actual']}%" if data['actual'] is not None else 'null'
            scopes_n = len(data.get('scopes', {}))
            discs_n  = len(data.get('disciplines', {}))
            struct   = f"scopes={scopes_n}" if scopes_n else f"discs={discs_n}"
            print(f"       concerns={len(data['concerns'])}, activities={len(data['activities'])}, plan={plan_s}, actual={actual_s}, {struct}")
            results[prj_id] = {'found': True, 'data': data}
            seeds_js.append(build_seed(prj_id, week, year, data))

    print()
    print(f"  Found:   {sum(1 for r in results.values() if r['found'])} projects")
    print(f"  Missing: {len(missing)} projects: {missing}")
    if not_due:
        print(f"  Not due: {len(not_due)} projects: {not_due}")

    # Quality gate. Every project that produced a PDF used to print "OK"
    # whatever came out of it, so a pattern that stopped matching stayed
    # invisible until someone opened the dashboard - Pak Lay ran null for two
    # weeks and GMTP for five that way. Check the shape of each result against
    # what its report type always yields, and against what the project itself
    # reported last week.
    if args.dump:
        with open(args.dump, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=1, ensure_ascii=False)
        print(f"  [dump] raw results written to {args.dump}")

    report = validate_run(results, week, year, history=history)
    print()
    print(format_findings(report))

    if args.dry_run:
        print("\n  [dry-run] Skipping file updates.")
        return

    if args.strict and report['counts'][FAIL]:
        print(f"\n  [strict] {report['counts'][FAIL]} FAIL finding(s) - "
              f"stopping before any file is written.")
        sys.exit(1)

    # Update files
    print("\n  Updating dashboard...")
    update_html(week, year, seeds_js, missing, report_date=report_date)
    update_excel(week, year, results)
    if args.no_push:
        print("  [no-push] Skipping git commit/push.")
    else:
        git_push(week, year)
    # Recorded after validation so this week is never its own baseline.
    save_history(record_run(history, results, week, year))

    xl_path = os.path.join(EXCEL_DIR, f'Gulf_Dashboard_W{week}_{year}.xlsx')
    if not args.no_email:
        send_email(week, year, results, missing, xl_path, report=report)
    notify(week, year,
           sum(1 for r in results.values() if r['found']),
           len(missing))

    print("\n  Done! Dashboard URL:", DASHBOARD_URL)
    print("=" * 60)


if __name__ == '__main__':
    main()
