# ── extract.py ── Extract concerns, activities & progress % from PDF ───────
import sys, os, re, glob
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

import pdfplumber
from patterns import (
    AUTO_EXTRACT_PROJECTS,
    IWTE_PAGE_INCLUDE, IWTE_PAGE_EXCLUDE, IWTE_ROWS,
    IWTE_CURR_PLAN_COL, IWTE_CURR_ACTUAL_COL,
    GMTP_PAGE_INCLUDE, GMTP_ROWS,
    SOLAR_THIS_WEEK, SOLAR_PREV_WEEK, SOLAR_PROGRESS,
    SOLAR_SCOPE_GUE, SOLAR_SCOPE_SIE, SOLAR_SCOPE_TL, SOLAR_SCOPE_COMM,
    SOLAR_PROGRESS_ALT, SOLAR_PROGRESS_SUMMARY, SOLAR_EXEC_SUMMARY_PROJECTS,
    WIND_SCURVE_PAGE, WIND_DISC_ROW, WIND_SCOPE_MAP,
    PAKBENG_PLAN_ACTUAL, PAKBENG_CUM_PLANNED, PAKBENG_CUM_ACTUAL,
    PAKLAY_PROGRESS,
)

# ── Regex helpers ────────────────────────────────────────────────────────────
_CONCERN_HEADERS  = re.compile(r'concern|area of concern|issue|delay', re.I)
_ACTIVITY_HEADERS = re.compile(
    r'next period|next week|planned activity|activity plan|upcoming', re.I)
_BULLET   = re.compile(r'^[\-•\*\d]+[\.\)]\s*')
_PAGE_NUM = re.compile(r'^\d+\s*$')
_SKIP_LINE = re.compile(
    r'^(action|table of content|page|note:|concern to achieve|next milestone)', re.I)
_PCT = r'(\d{1,3}(?:\.\d{1,2})?)'


# ── Text utilities ────────────────────────────────────────────────────────────
def _clean(lines):
    out = []
    for ln in lines:
        ln = ln.strip()
        if not ln or _PAGE_NUM.match(ln):
            continue
        if _SKIP_LINE.match(ln):
            continue
        if len(ln) < 8:
            continue
        ln = _BULLET.sub('', ln).strip()
        if ln:
            out.append(ln)
    return out


def _extract_section(text, header_re, stop_re=None, max_lines=30):
    lines = text.split('\n')
    collecting, result, blank_streak = False, [], 0
    for ln in lines:
        if not collecting:
            if header_re.search(ln):
                collecting = True
            continue
        if stop_re and stop_re.search(ln):
            break
        if not ln.strip():
            blank_streak += 1
            if blank_streak >= 3:
                break
        else:
            blank_streak = 0
            result.append(ln)
        if len(result) >= max_lines:
            break
    return _clean(result)


# ── Progress extraction helpers ───────────────────────────────────────────────

def _parse_pct(s):
    """Return float or None from a string that should be a percentage."""
    try:
        v = float(str(s).strip().rstrip('%').strip())
        if 0 <= v <= 100:
            return v
    except (ValueError, AttributeError):
        pass
    return None


def _find_two_pcts_in_line(line):
    """Return (plan, actual) as floats if line contains exactly 2 pct-like numbers."""
    nums = re.findall(r'\d{1,3}(?:\.\d{1,2})?', line)
    nums = [_parse_pct(n) for n in nums]
    nums = [n for n in nums if n is not None]
    if len(nums) >= 2:
        return nums[0], nums[1]
    return None, None


def _is_multicol_row(row, min_cells=3):
    """
    True if a pdfplumber table row looks like a real multi-column row.
    Some pages make pdfplumber emit a spurious extra "row" that dumps the
    entire page's text into a single cell with every other cell None (an
    artifact of its table-grid detection) — that blob can accidentally
    contain header/section keywords as substrings and get mistaken for the
    real header/data row. Real rows in these Exec Summary tables always have
    several distinct populated cells (Topic + one per scope column).
    """
    if not row:
        return False
    return sum(1 for c in row if c not in (None, '')) >= min_cells


def _table_rows_to_disc(table):
    """
    Given a pdfplumber table (list of rows), try to extract disc plan/actual.
    Looks for rows whose first cell matches a discipline name.
    Returns dict: {disc_name: (plan, actual)} and overall (plan, actual).
    """
    discs = {}
    overall = (None, None)
    DISC_MAP = {
        'engineering': 'Engineering',
        'procurement': 'Procurement',
        'construction': 'Construction',
        'commissioning': 'Commissioning',
    }
    for row in table:
        if not row or row[0] is None:
            continue
        label = str(row[0]).strip().lower()
        # Find which discipline this row is
        disc = None
        for key, name in DISC_MAP.items():
            if key in label:
                disc = name
                break
        is_overall = 'overall' in label

        if disc or is_overall:
            # Collect all numeric cells
            nums = []
            for cell in row[1:]:
                if cell is None:
                    continue
                v = _parse_pct(str(cell).strip())
                if v is not None:
                    nums.append(v)
            if len(nums) >= 2:
                # If 4 numbers → This Week Plan, This Week Actual, Cumul Plan, Cumul Actual
                # We want cumulative (index 2, 3 if 4 nums; index 0, 1 if 2 nums)
                if len(nums) >= 4:
                    p, a = nums[2], nums[3]
                else:
                    p, a = nums[0], nums[1]
                if disc:
                    discs[disc] = (p, a)
                elif is_overall:
                    overall = (p, a)
    return discs, overall


# ── iWTE extractor ────────────────────────────────────────────────────────────

def extract_progress_iwte(pdf):
    """
    iWTE weekly report: find the page with the Weekly Progress table.
    Returns { plan, actual, disciplines: {disc: {plan, actual}} }.

    Page identification: contains 'weekly' or 'week NN' but NOT 'monthly'.
    Table has: Engineering / Procurement / Construction / Commissioning / Overall rows.
    """
    with pdfplumber.open(pdf) as doc:
        for page in doc.pages:
            text = page.extract_text() or ''
            tl = text.lower()

            # Must look like a weekly page, must NOT be a monthly table
            if not IWTE_PAGE_INCLUDE.search(tl):
                continue
            if IWTE_PAGE_EXCLUDE.search(tl):
                continue
            # Must have at least one discipline name
            if not any(d.lower() in tl for d in
                       ['engineering', 'procurement', 'construction', 'commissioning']):
                continue

            # ── Try table extraction (iWTE: cols 5=curr_plan, 6=curr_actual) ─
            tables = page.extract_tables()
            for tbl in tables:
                if not tbl:
                    continue
                discs = {}
                overall_plan = overall_actual = None
                DISC_MAP = {'engineering': 'Engineering', 'procurement': 'Procurement',
                            'construction': 'Construction', 'commissioning': 'Commissioning'}
                for row in tbl:
                    if not row or row[0] is None:
                        continue
                    label = str(row[0]).strip().lower()
                    disc = DISC_MAP.get(label)
                    is_overall = 'overall' in label and disc is None

                    if disc or is_overall:
                        # Try current-week columns first (9-col format)
                        if len(row) >= IWTE_CURR_ACTUAL_COL + 1:
                            p = _parse_pct(str(row[IWTE_CURR_PLAN_COL] or ''))
                            a = _parse_pct(str(row[IWTE_CURR_ACTUAL_COL] or ''))
                        else:
                            # Fallback: collect all numeric cells
                            nums = [_parse_pct(str(c)) for c in row[1:] if c]
                            nums = [n for n in nums if n is not None]
                            p, a = (nums[0], nums[1]) if len(nums) >= 2 else (None, None)

                        if p is not None and a is not None:
                            if is_overall:
                                overall_plan, overall_actual = p, a
                            else:
                                discs[disc] = {'plan': p, 'actual': a}

                if len(discs) >= 3:
                    if overall_plan is None and discs:
                        weights = {'Engineering': 0.10, 'Procurement': 0.40,
                                   'Construction': 0.40, 'Commissioning': 0.10}
                        wp = wa = wt = 0
                        for d, v in discs.items():
                            w = weights.get(d, 0)
                            wp += v['plan'] * w; wa += v['actual'] * w; wt += w
                        if wt > 0:
                            overall_plan = round(wp / wt, 2)
                            overall_actual = round(wa / wt, 2)
                    disciplines = {d: v for d, v in discs.items()}
                    return {'plan': overall_plan, 'actual': overall_actual,
                            'disciplines': disciplines}

            # ── Fallback: text regex (older iWTE format with 2-col table) ───
            discs = {}
            overall_plan = overall_actual = None
            for disc, pattern in IWTE_ROWS.items():
                m = pattern.search(text)
                if m:
                    p, a = _parse_pct(m.group(1)), _parse_pct(m.group(2))
                    if disc == 'Overall':
                        overall_plan, overall_actual = p, a
                    elif p is not None and a is not None:
                        discs[disc] = {'plan': p, 'actual': a}

            if len(discs) >= 2:
                if overall_plan is None and discs:
                    weights = {'Engineering': 0.10, 'Procurement': 0.40,
                               'Construction': 0.40, 'Commissioning': 0.10}
                    wp = wa = wt = 0
                    for d, v in discs.items():
                        w = weights.get(d, 0)
                        wp += v['plan'] * w; wa += v['actual'] * w; wt += w
                    if wt > 0:
                        overall_plan = round(wp / wt, 2)
                        overall_actual = round(wa / wt, 2)
                return {'plan': overall_plan, 'actual': overall_actual,
                        'disciplines': discs}

    return {'plan': None, 'actual': None, 'disciplines': {}}


# ── GMTP extractor ─────────────────────────────────────────────────────────────

_GMTP_OVERALL_ROW = re.compile(r'^overall\s+progress\b', re.I)


def _parse_gmtp_overall_row(text):
    """
    GMTP Section 3.1.1 'Overall Progress' table: a per-package breakdown
    (LNG Tank / BOP / Marine / Commissioning / Overall Progress) where the
    'Overall Progress' row has columns:
      [Eng P,A][Proc P,A][Constr P,A][Comm P,A][Subtotal-prev P,A]
      [Subtotal-curr P,A][Variance][Weight]
    Packages missing a scope (e.g. no Commissioning work) render that cell as
    '-', which just drops out of the number list — so we index from the END
    (stable regardless of missing cells): the last 4 numbers before
    Variance/Weight are [PrevPlan, PrevActual, CurrPlan, CurrActual].
    """
    for line in text.split('\n'):
        if _GMTP_OVERALL_ROW.match(line.strip()) and '%' in line:
            nums = [_parse_pct(n) for n in re.findall(r'\d{1,3}(?:\.\d{1,2})?', line)]
            nums = [n for n in nums if n is not None]
            if len(nums) >= 6:
                plan, actual = nums[-4], nums[-3]
                discs = {}
                for i, label in enumerate(
                        ['Engineering', 'Procurement', 'Construction', 'Commissioning']):
                    lead = nums[:-6]
                    if len(lead) >= (i + 1) * 2:
                        discs[label] = {'plan': lead[i * 2], 'actual': lead[i * 2 + 1]}
                return plan, actual, discs
    return None, None, {}


def extract_progress_gmtp(pdf):
    """
    GMTP/LNG report: find the Overall Progress table (text-based).
    Looks for page with discipline breakdown and plan/actual headers.
    Returns { plan, actual, disciplines: {disc: {plan, actual}} }.
    """
    with pdfplumber.open(pdf) as doc:
        for page in doc.pages:
            text = page.extract_text() or ''
            tl = text.lower()

            if not GMTP_PAGE_INCLUDE.search(tl):
                continue
            if not any(d.lower() in tl for d in
                       ['engineering', 'procurement', 'construction']):
                continue

            # ── Try 'Overall Progress' per-package breakdown row first ───────
            plan, actual, discs = _parse_gmtp_overall_row(text)
            if plan is not None and actual is not None:
                return {'plan': plan, 'actual': actual, 'disciplines': discs}

            # ── Try table extraction ─────────────────────────────────────────
            tables = page.extract_tables()
            for tbl in tables:
                if not tbl:
                    continue
                discs, overall = _table_rows_to_disc(tbl)
                if len(discs) >= 2:
                    plan, actual = overall
                    disciplines = {d: {'plan': p, 'actual': a}
                                   for d, (p, a) in discs.items()}
                    return {'plan': plan, 'actual': actual,
                            'disciplines': disciplines}

            # ── Fallback: text regex ─────────────────────────────────────────
            discs = {}
            overall_plan = overall_actual = None
            for disc, pattern in GMTP_ROWS.items():
                m = pattern.search(text)
                if m:
                    p, a = _parse_pct(m.group(1)), _parse_pct(m.group(2))
                    if disc == 'Overall':
                        overall_plan, overall_actual = p, a
                    elif p is not None and a is not None:
                        discs[disc] = {'plan': p, 'actual': a}

            if len(discs) >= 2:
                return {'plan': overall_plan, 'actual': overall_actual,
                        'disciplines': discs}

    return {'plan': None, 'actual': None, 'disciplines': {}}


# ── Solar extractor ───────────────────────────────────────────────────────────

def _detect_solar_scope(text):
    """Return scope name based on page title/content, or None.
    Siemens/Substation must be checked before 115kV T/L: substation page
    titles often contain "115kV" too (e.g. "115kV Substation (Siemens)"),
    which would otherwise be misclassified as the T/L scope.
    """
    if SOLAR_SCOPE_GUE.search(text):
        return 'GUE'
    if SOLAR_SCOPE_SIE.search(text):
        return 'Siemens'
    if SOLAR_SCOPE_TL.search(text):
        return '115kV T/L'
    if SOLAR_SCOPE_COMM.search(text):
        return 'Comm/Fiber'
    return None

def _parse_solar_progress_from_text(text):
    """
    Extract discipline plan/actual from a Solar 'This week activities' page.
    Tries both NWT3-style ("stands at X% against planned Y%") and
    PTN-style ("X% compared to plan at Y%") patterns.
    Returns (plan, actual, disciplines_dict).

    NOTE: some reports (e.g. STN) lay Previous-week and This-week out as two
    columns that pdfplumber linearizes onto the SAME text line, so each label
    (Overall/Engineering/...) appears twice per line — previous week first,
    this week second. We deliberately let each new match overwrite the
    previous one (no "already set" guard) so the LAST occurrence — this
    week's column — wins. Single-column reports only ever produce one match
    per label, so this is a no-op for them.
    """
    overall_plan = overall_actual = None
    discs = {}

    # Try all patterns; prefer 'overall' for the top-level; 'construction' as fallback
    construction_plan = construction_actual = None
    for pattern in (SOLAR_PROGRESS, SOLAR_PROGRESS_ALT, SOLAR_PROGRESS_SUMMARY):
        for m in pattern.finditer(text):
            label = m.group(1).lower()
            actual_v = _parse_pct(m.group(2))
            plan_v   = _parse_pct(m.group(3))
            if actual_v is None or plan_v is None:
                continue
            if label == 'overall':
                overall_plan, overall_actual = plan_v, actual_v
            elif label == 'construction':
                construction_plan, construction_actual = plan_v, actual_v
                discs['Construction'] = {'plan': plan_v, 'actual': actual_v}
            elif label in ('engineering', 'procurement', 'commissioning'):
                discs[label.capitalize()] = {'plan': plan_v, 'actual': actual_v}

    # Use construction as proxy for overall if no explicit overall line
    if overall_plan is None and construction_plan is not None:
        overall_plan, overall_actual = construction_plan, construction_actual

    return overall_plan, overall_actual, discs

def extract_progress_solar(pdf, prj_id=None):
    """
    Solar report: find 'This week activities' pages per contractor section.
    Builds scopes dict: { scope_name: { plan, actual, disciplines? } }
    Returns { plan, actual, scopes }.

    For prj_id in SOLAR_EXEC_SUMMARY_PROJECTS, the Executive Summary 'Site
    Progress' row is authoritative (see patterns.py) and is merged in after
    the scan below: its plan/actual wins per scope, but a scope's
    discipline breakdown from this scan is kept if Exec Summary has none.
    """
    scopes = {}

    with pdfplumber.open(pdf) as doc:
        for page in doc.pages:
            text = page.extract_text() or ''

            if not SOLAR_THIS_WEEK.search(text):
                continue

            # Slice to "This week activities" section only
            # (some pages have both "Previous week" and "This week" sections)
            tw_match = SOLAR_THIS_WEEK.search(text)
            this_week_text = text[tw_match.start():]

            # Detect scope from the page TITLE (first line) only — not the full
            # page text. Some Comm/Fiber pages also mention "Transmission Line"
            # further down (e.g. a cable-length tracking table), which used to
            # make the full-text search misclassify them as '115kV T/L' and
            # silently overwrite/skip the real T/L scope's data.
            title = text.split('\n', 1)[0]
            scope = _detect_solar_scope(title)
            if not scope:
                continue

            plan_v, actual_v, discs = _parse_solar_progress_from_text(this_week_text)

            if plan_v is not None or discs:
                entry = {'plan': plan_v, 'actual': actual_v}
                if scope == 'GUE' and discs:
                    entry['disciplines'] = discs
                scopes[scope] = entry

    if prj_id in SOLAR_EXEC_SUMMARY_PROJECTS:
        # Exec Summary is authoritative: merge its plan/actual per scope over
        # this scan's (fills scopes this scan missed entirely, e.g. NWT3's
        # Comm/Fiber), but keep this scan's discipline breakdown when Exec
        # Summary's cell has none (it only ever shows Overall + Construction).
        exec_scopes = _extract_solar_from_exec_summary(pdf)
        for scope_name, exec_entry in exec_scopes.items():
            tw_entry = scopes.get(scope_name)
            if tw_entry and tw_entry.get('disciplines'):
                # Union, exec's per-discipline values win on key collision
                # (e.g. Construction) — tw-only disciplines (e.g. GUE's
                # Engineering/Procurement, absent from Exec Summary) survive.
                merged_disc = dict(tw_entry['disciplines'])
                merged_disc.update(exec_entry.get('disciplines') or {})
                exec_entry['disciplines'] = merged_disc
            scopes[scope_name] = exec_entry
    elif not scopes:
        # Fallback: scan Executive Summary table "Site Progress" row
        scopes = _extract_solar_from_exec_summary(pdf)

    if not scopes:
        return {'plan': None, 'actual': None, 'scopes': {}}

    # Use GUE overall as project overall (largest scope)
    overall_plan = overall_actual = None
    gue = scopes.get('GUE')
    if gue and gue.get('plan') is not None:
        overall_plan  = gue['plan']
        overall_actual = gue['actual']

    return {'plan': overall_plan, 'actual': overall_actual, 'scopes': scopes}


def _extract_solar_from_exec_summary(pdf):
    """
    Fallback: extract from Executive Summary table 'Site Progress' row.
    Scope columns are identified by header row (GUE, Siemens, T/L).
    Each cell has lines like 'Construction progress Y% /Plan at Z%'.
    """
    scopes = {}
    # NOTE: no bare 'substation' → 'Siemens' mapping. Some reports (e.g. SDCE)
    # have a SEPARATE "GIS Substation (PEA&TCCL)" column in addition to the
    # real "Substation (SIEMENS)" column; a generic 'substation' keyword
    # matches both headers and the later column silently overwrote the real
    # Siemens data. 'siemens' alone is specific enough for the real column.
    SCOPE_COL_NAMES = {
        'gue': 'GUE', 'pvfarm': 'GUE', 'pv farm': 'GUE',
        'siemens': 'Siemens',
        '115kv': '115kV T/L', 'transmission': '115kV T/L',
        'fiber': 'Comm/Fiber', 'communication': 'Comm/Fiber',
        'gis substation': 'GIS Substation',
        'add bay': 'Add Bay',
    }

    with pdfplumber.open(pdf) as doc:
        for page in doc.pages[:10]:  # exec summary is always early
            text = page.extract_text() or ''
            if 'executive summary' not in text.lower() and 'site progress' not in text.lower():
                continue
            tables = page.extract_tables()
            for tbl in tables:
                if not tbl:
                    continue
                # Find header row (contains scope names) and Site Progress row
                header_row = None
                site_progress_row = None
                for row in tbl:
                    if not _is_multicol_row(row):
                        continue
                    joined = ' '.join(str(c or '') for c in row).lower()
                    if any(k in joined for k in ['gue', 'pvfarm', 'pv farm', 'siemens']):
                        if header_row is None:
                            header_row = row
                    if 'site progress' in joined or 'siteprogress' in joined.replace(' ', ''):
                        site_progress_row = row

                if header_row is None or site_progress_row is None:
                    continue

                # Map column index → scope name
                col_scope = {}
                for ci, cell in enumerate(header_row):
                    if not cell:
                        continue
                    cl = str(cell).lower()
                    for key, scope in SCOPE_COL_NAMES.items():
                        if key in cl:
                            col_scope[ci] = scope
                            break

                # Extract plan/actual from each scope cell in Site Progress row
                for ci, cell in enumerate(site_progress_row):
                    scope = col_scope.get(ci)
                    if not scope or not cell:
                        continue
                    cell_text = str(cell)
                    plan_v = actual_v = None
                    construction_v = None
                    discs = {}
                    # First match per label wins. Cells often follow the
                    # headline "Overall progress X% against plan Y%" line with
                    # bullet sub-items that reuse the same phrasing for a
                    # single piece of equipment (e.g. "Transformer station
                    # overall progress 98.50% vs plan 100.00%") — overwriting
                    # on every match let that sub-item clobber the true
                    # headline Overall value.
                    # Try all cell phrasings seen across reports: SSE/LNE/SDCE
                    # ("/Plan at" / "vs plan" / "against plan"), NWT3-style
                    # ("stands at X% against the planned progress of Y%"), and
                    # PTN/STN-style ("progress X% compared to plan at Y%").
                    for pattern in (SOLAR_PROGRESS_SUMMARY, SOLAR_PROGRESS, SOLAR_PROGRESS_ALT):
                        for m in pattern.finditer(cell_text):
                            label = m.group(1).lower()
                            a = _parse_pct(m.group(2))
                            p = _parse_pct(m.group(3))
                            if a is None or p is None:
                                continue
                            if label == 'overall':
                                if actual_v is None:
                                    actual_v, plan_v = a, p
                            elif label == 'construction':
                                if construction_v is None:
                                    construction_v = (p, a)
                                    discs['Construction'] = {'plan': p, 'actual': a}
                            elif label in ('engineering', 'procurement'):
                                if label.capitalize() not in discs:
                                    discs[label.capitalize()] = {'plan': p, 'actual': a}
                    # Use construction as proxy for overall if no explicit overall line
                    if plan_v is None and construction_v is not None:
                        plan_v, actual_v = construction_v
                    # Some columns state Overall with no plan comparison at all
                    # (e.g. "Transmission line Overall progress is 100.00%")
                    # — treat the single figure as both plan and actual rather
                    # than losing the scope entirely.
                    if plan_v is None:
                        m = re.search(
                            r'overall\s*progress\s*(?:is\s+)?(\d{1,3}(?:\.\d{1,2})?)\s*%',
                            cell_text, re.I)
                        if m:
                            plan_v = actual_v = _parse_pct(m.group(1))
                    if plan_v is not None:
                        entry = {'plan': plan_v, 'actual': actual_v}
                        if scope == 'GUE' and discs:
                            entry['disciplines'] = discs
                        scopes[scope] = entry

    return scopes


def extract_solar_concerns_from_exec_summary(pdf):
    """
    Read the 'Concern need management attention' row of the same Executive
    Summary table used by _extract_solar_from_exec_summary, instead of the
    generic keyword-based bullet scan used elsewhere in extract_from_pdf.
    Each non-N/A cell is prefixed with its column header for context.
    Returns a list of concern strings (possibly empty).
    """
    concerns = []
    with pdfplumber.open(pdf) as doc:
        for page in doc.pages[:10]:
            text = page.extract_text() or ''
            if 'executive summary' not in text.lower() and 'site progress' not in text.lower():
                continue
            for tbl in page.extract_tables():
                if not tbl:
                    continue
                header_row = None
                concern_row = None
                for row in tbl:
                    if not _is_multicol_row(row):
                        continue
                    joined = ' '.join(str(c or '') for c in row).lower()
                    if any(k in joined for k in ['gue', 'pvfarm', 'pv farm', 'siemens']):
                        if header_row is None:
                            header_row = row
                    if 'concern' in joined.replace(' ', ''):
                        concern_row = row

                if header_row is None or concern_row is None:
                    continue

                bullet_re = re.compile(r'^[•▪❑\-]\s*')
                for ci, cell in enumerate(concern_row):
                    header = str(header_row[ci]) if ci < len(header_row) and header_row[ci] else None
                    if not header or not cell or header.strip().lower() == 'topic':
                        continue
                    header_label = ' '.join(header.split())
                    # A single bullet's text wraps across multiple lines in
                    # the PDF (narrow column) — only lines starting with a
                    # bullet marker begin a new item; other lines continue
                    # the previous one.
                    items, current = [], None
                    for line in str(cell).split('\n'):
                        line = line.strip()
                        if not line:
                            continue
                        if bullet_re.match(line):
                            if current:
                                items.append(current)
                            current = bullet_re.sub('', line)
                        elif current is not None:
                            current += ' ' + line
                        else:
                            current = line
                    if current:
                        items.append(current)
                    for it in items:
                        # The last column's cell often has the page number
                        # bleed into its text (e.g. "N/A 4") since it sits at
                        # the page's bottom-right corner — strip a trailing
                        # bare number before the N/A check.
                        it_check = re.sub(r'\s*\d+\s*$', '', it).strip()
                        if not it_check or it_check.upper() in ('N/A', 'NA'):
                            continue
                        concerns.append(f'[{header_label}] {it}')
    return concerns


# ── Wind extractor ─────────────────────────────────────────────────────────────

def extract_progress_wind(pdf):
    """
    Wind report: find 'Overall Progress S-Curve (SCOPE)' pages.
    Each page has a table: Description | Plan% | Actual% | Ahead/Delay
    Rows: Engineering, Procurement, Construction, Overall
    Returns { plan, actual, scopes }.
    """
    scopes = {}

    with pdfplumber.open(pdf) as doc:
        for page in doc.pages:
            text = page.extract_text() or ''

            m = WIND_SCURVE_PAGE.search(text)
            if not m:
                continue

            raw_scope = m.group(1).strip()
            # Normalize scope name
            scope = WIND_SCOPE_MAP.get(raw_scope.lower(), raw_scope.upper())

            # Skip "Construction Progress S-Curve" pages (only one row, less useful)
            if re.search(r'^construction progress s-?curve', text.strip(), re.I):
                continue

            # Parse the discipline rows
            discs = {}
            overall_plan = overall_actual = None

            for dm in WIND_DISC_ROW.finditer(text):
                label = dm.group(1).lower()
                plan_v   = _parse_pct(dm.group(2))
                actual_v = _parse_pct(dm.group(3))
                if plan_v is None or actual_v is None:
                    continue
                if label == 'overall':
                    overall_plan, overall_actual = plan_v, actual_v
                elif label in ('engineering', 'procurement', 'construction', 'commissioning'):
                    discs[label.capitalize()] = {'plan': plan_v, 'actual': actual_v}

            # Also try table extraction as fallback
            if not discs and not overall_plan:
                tables = page.extract_tables()
                for tbl in tables:
                    d, ov = _table_rows_to_disc(tbl)
                    if d or ov[0] is not None:
                        discs = {k: {'plan': p, 'actual': a} for k, (p, a) in d.items()}
                        overall_plan, overall_actual = ov
                        break

            if overall_plan is not None or discs:
                entry = {'plan': overall_plan, 'actual': overall_actual}
                if discs:
                    entry['disciplines'] = discs
                scopes[scope] = entry

    if not scopes:
        return {'plan': None, 'actual': None, 'scopes': {}}

    # Use CBOP overall as project overall (largest scope)
    cbop = scopes.get('CBOP')
    overall_plan = cbop['plan'] if cbop else None
    overall_actual = cbop['actual'] if cbop else None

    return {'plan': overall_plan, 'actual': overall_actual, 'scopes': scopes}


# ── Hydro extractors ───────────────────────────────────────────────────────────

def extract_progress_hydro_pakbeng(pdf):
    """
    Pak Beng monthly report.
    Pattern 1: 'Plan/Actual：X.XX%/Y.YY%（Z.ZZ%）'
    Pattern 2: 'Cum. Progress -Planned ... X.XX%' and 'Cum. Progress -Actual ... Y.YY%'
    Returns { plan, actual, disciplines: {} }.
    """
    with pdfplumber.open(pdf) as doc:
        full_text = '\n'.join(p.extract_text() or '' for p in doc.pages)

    # Pattern 1: Plan/Actual label
    m = PAKBENG_PLAN_ACTUAL.search(full_text)
    if m:
        plan_v   = _parse_pct(m.group(1))
        actual_v = _parse_pct(m.group(2))
        if plan_v is not None and actual_v is not None:
            return {'plan': plan_v, 'actual': actual_v, 'disciplines': {}}

    # Pattern 2: Cumulative Progress rows — take last value (current month)
    mp = PAKBENG_CUM_PLANNED.findall(full_text)
    ma = PAKBENG_CUM_ACTUAL.findall(full_text)
    if mp and ma:
        plan_v   = _parse_pct(mp[-1])
        actual_v = _parse_pct(ma[-1])
        if plan_v is not None and actual_v is not None:
            return {'plan': plan_v, 'actual': actual_v, 'disciplines': {}}

    return {'plan': None, 'actual': None, 'disciplines': {}}


def extract_progress_hydro_paklay(pdf, search_dir=None):
    """
    Pak Lay: EPC weekly report is primarily issue/design tracking and rarely
    states an overall %. Tries 'Physical Progress: Plan X% / Actual Y%' in
    the EPC report first; if that's absent and search_dir is given, falls
    back to OCR-reading the Owner deck's 'Progress S-Curve from EPC' slide
    (see _ocr_paklay_scurve — that data exists only as a screenshot image).
    Returns { plan, actual, disciplines: {} }.
    """
    try:
        with pdfplumber.open(pdf) as doc:
            full_text = '\n'.join(p.extract_text() or '' for p in doc.pages[:20])

        m = PAKLAY_PROGRESS.search(full_text)
        if m:
            plan_v   = _parse_pct(m.group(1))
            actual_v = _parse_pct(m.group(2))
            if plan_v is not None and actual_v is not None:
                return {'plan': plan_v, 'actual': actual_v, 'disciplines': {}}
    except Exception as e:
        print(f"  [paklay error] {e}")

    if search_dir:
        ocr_result = _ocr_paklay_scurve(search_dir)
        if ocr_result['plan'] is not None or ocr_result['actual'] is not None:
            return {'plan': ocr_result['plan'], 'actual': ocr_result['actual'],
                    'disciplines': {}}

    return {'plan': None, 'actual': None, 'disciplines': {}}


def _ocr_pct_from_band(page_img, arr, target_rgb, y_limit, tol=10, pad=8):
    """
    Locate a solid-color highlight band (by RGB) in a rendered page image,
    restricted to the top y_limit px (the bands sit above the actual chart),
    crop it with padding, upscale, and OCR the single line of text inside
    for a '= X.XX%' value. Returns float or None — never raises.
    """
    import numpy as np
    import pytesseract
    from PIL import Image

    search = arr[:y_limit]
    diff = np.abs(search.astype(int) - np.array(target_rgb)).sum(axis=2)
    ys, xs = np.where(diff < tol)
    if len(xs) == 0:
        return None
    x0, y0, x1, y1 = int(xs.min()), int(ys.min()), int(xs.max()), int(ys.max())
    crop = page_img.crop((max(x0 - pad, 0), max(y0 - pad, 0), x1 + pad, y1 + pad))
    crop = crop.resize((crop.width * 3, crop.height * 3), Image.LANCZOS)
    text = pytesseract.image_to_string(crop, config='--psm 7')
    m = re.search(r'=\s*(\d{1,3}(?:\.\d{1,2})?)\s*%', text)
    return _parse_pct(m.group(1)) if m else None


def _ocr_paklay_scurve(search_dir):
    """
    Pak Lay's Owner 'Progress of works' deck has a slide titled 'Progress
    S-Curve from EPC' where the Cumulative Plan/Achieved % are baked into a
    pasted screenshot image — not real PDF text — inside two highlighted
    bands: a light-blue 'Accumulative Planned Progress ... =X%' line and a
    light-orange 'Accumulative Achived Progress ... =Y%' line. OCR is the
    only way to read them. Searches every PDF under search_dir (the file
    isn't always named the same week to week) for that slide.
    Returns {'plan': float|None, 'actual': float|None} — never raises;
    any failure (Tesseract missing, slide not found, OCR miss) yields Nones
    so a bad OCR read can never break the weekly run.
    """
    try:
        import numpy as np
        import pytesseract
        tess_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        if os.path.exists(tess_path):
            pytesseract.pytesseract.tesseract_cmd = tess_path

        for pdf_path in glob.glob(os.path.join(search_dir, '**', '*.pdf'), recursive=True):
            try:
                with pdfplumber.open(pdf_path) as doc:
                    target_page = None
                    for page in doc.pages:
                        t = (page.extract_text() or '').lower()
                        if 's-curve from epc' in t or 's curve from epc' in t:
                            target_page = page
                            break
                    if target_page is None:
                        continue

                    page_img = target_page.to_image(resolution=300).original.convert('RGB')
                    arr = np.array(page_img)
                    y_limit = int(arr.shape[0] * 0.4)

                    plan_v   = _ocr_pct_from_band(page_img, arr, (167, 204, 242), y_limit)
                    actual_v = _ocr_pct_from_band(page_img, arr, (247, 182, 150), y_limit)
                    return {'plan': plan_v, 'actual': actual_v}
            except Exception as e:
                print(f"  [paklay OCR error] {os.path.basename(pdf_path)}: {e}")
    except Exception as e:
        print(f"  [paklay OCR error] {e}")

    return {'plan': None, 'actual': None}


# ── Progress dispatcher ───────────────────────────────────────────────────────

def extract_progress(pdf_path, prj_id, search_dir=None):
    """
    Dispatch to the correct extractor based on project type.
    Returns dict with: plan, actual, and either disciplines or scopes.
    """
    extract_type = AUTO_EXTRACT_PROJECTS.get(prj_id)
    try:
        if extract_type == 'iwte':
            return extract_progress_iwte(pdf_path)
        elif extract_type == 'gmtp':
            return extract_progress_gmtp(pdf_path)
        elif extract_type == 'solar':
            return extract_progress_solar(pdf_path, prj_id=prj_id)
        elif extract_type == 'wind':
            return extract_progress_wind(pdf_path)
        elif extract_type == 'hydro_pakbeng':
            return extract_progress_hydro_pakbeng(pdf_path)
        elif extract_type == 'hydro_paklay':
            return extract_progress_hydro_paklay(pdf_path, search_dir=search_dir)
    except Exception as e:
        print(f"  [progress extract error] {e}")
    return {'plan': None, 'actual': None, 'disciplines': {}}


# ── Main entry point ──────────────────────────────────────────────────────────

def extract_from_pdf(pdf_path, prj_id=None, search_dir=None):
    """
    Returns dict:
      { concerns: [...], activities: [...],
        plan: float|None, actual: float|None,
        disciplines: {disc: {plan, actual}} }
    search_dir (optional): the project's report folder, used by extractors
    that must look at a second PDF beyond the one already matched for this
    project (e.g. Pak Lay's OCR fallback lives in a sibling 'Owner' file).
    """
    concerns   = []
    activities = []
    seen_c = set()
    seen_a = set()

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                tl   = text.lower()

                if 'table of content' in tl:
                    continue

                # iWTE "Concern to Achieve Milestone" pages open with a
                # "Next Milestone :" list BEFORE the real "Concern (delay)"
                # section. _CONCERN_HEADERS matches broadly on the word
                # "concern" anywhere, so on these pages it was matching the
                # page title ("Concern to Achieve Milestone") and collecting
                # the Milestone list as if it were the concerns — the actual
                # delay concerns just below were skipped. Skip the generic
                # pass here entirely and let the precise pass below handle it.
                is_iwte_concern_page = 'concern (delay)' in tl or 'concern(delay)' in tl

                # Concerns
                if not is_iwte_concern_page and _CONCERN_HEADERS.search(tl) and '%' not in text:
                    items = _extract_section(
                        text, _CONCERN_HEADERS,
                        stop_re=_ACTIVITY_HEADERS)
                    for it in items:
                        if it not in seen_c and len(it) > 10:
                            seen_c.add(it)
                            concerns.append(it)

                # iWTE-style "Concern (Delay)" section
                if is_iwte_concern_page:
                    items = _extract_section(
                        text,
                        re.compile(r'concern.*delay', re.I),
                        stop_re=re.compile(r'^action', re.I))
                    for it in items:
                        if it not in seen_c and len(it) > 10:
                            seen_c.add(it)
                            concerns.append(it)

                # Next period activities
                if _ACTIVITY_HEADERS.search(tl):
                    items = _extract_section(text, _ACTIVITY_HEADERS)
                    for it in items:
                        if it not in seen_a and len(it) > 10:
                            seen_a.add(it)
                            activities.append(it)

    except Exception as e:
        print(f"  [extract error] {e}")

    # SOLAR_EXEC_SUMMARY_PROJECTS (SDCE, and NWT3/PTN/STN as of W27): the
    # generic bullet scan above picks up noisy/incorrect text elsewhere in
    # the report; the Executive Summary's own 'Concern need management
    # attention' row is the authoritative source for these projects.
    if prj_id in SOLAR_EXEC_SUMMARY_PROJECTS:
        try:
            exec_concerns = extract_solar_concerns_from_exec_summary(pdf_path)
            if exec_concerns:
                concerns = exec_concerns
        except Exception as e:
            print(f"  [sdce concerns error] {e}")

    # Extract progress %
    progress = extract_progress(pdf_path, prj_id, search_dir=search_dir) if prj_id else \
               {'plan': None, 'actual': None, 'disciplines': {}}

    result = {
        'concerns':    concerns[:20],
        'activities':  activities[:20],
        'plan':        progress.get('plan'),
        'actual':      progress.get('actual'),
        'disciplines': progress.get('disciplines', {}),
    }
    # Multi-scope projects (Solar/Wind) — pass scopes through
    if 'scopes' in progress:
        result['scopes'] = progress['scopes']
    return result
