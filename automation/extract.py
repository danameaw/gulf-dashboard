# ── extract.py ── Extract concerns, activities & progress % from PDF ───────
import sys, re
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

import pdfplumber
from patterns import (
    AUTO_EXTRACT_PROJECTS,
    IWTE_PAGE_INCLUDE, IWTE_PAGE_EXCLUDE, IWTE_ROWS,
    IWTE_CURR_PLAN_COL, IWTE_CURR_ACTUAL_COL,
    GMTP_PAGE_INCLUDE, GMTP_ROWS,
    SOLAR_THIS_WEEK, SOLAR_PREV_WEEK, SOLAR_PROGRESS,
    SOLAR_SCOPE_GUE, SOLAR_SCOPE_SIE, SOLAR_SCOPE_TL, SOLAR_SCOPE_COMM,
    SOLAR_PROGRESS_ALT, SOLAR_PROGRESS_SUMMARY,
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

def extract_progress_solar(pdf):
    """
    Solar report: find 'This week activities' pages per contractor section.
    Builds scopes dict: { scope_name: { plan, actual, disciplines? } }
    Returns { plan, actual, scopes }.
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

            scope = _detect_solar_scope(text)  # scope from full page title
            if not scope:
                continue

            plan_v, actual_v, discs = _parse_solar_progress_from_text(this_week_text)

            if plan_v is not None or discs:
                entry = {'plan': plan_v, 'actual': actual_v}
                if scope == 'GUE' and discs:
                    entry['disciplines'] = discs
                scopes[scope] = entry

    # Fallback: scan Executive Summary table "Site Progress" row
    if not scopes:
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
                    if not row:
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
                    for m in SOLAR_PROGRESS_SUMMARY.finditer(cell_text):
                        label = m.group(1).lower()
                        a = _parse_pct(m.group(2))
                        p = _parse_pct(m.group(3))
                        if a is None or p is None:
                            continue
                        if label == 'overall':
                            actual_v, plan_v = a, p
                        elif label == 'construction':
                            construction_v = (p, a)
                            discs['Construction'] = {'plan': p, 'actual': a}
                        elif label in ('engineering', 'procurement'):
                            discs[label.capitalize()] = {'plan': p, 'actual': a}
                    # Use construction as proxy for overall if no explicit overall line
                    if plan_v is None and construction_v is not None:
                        plan_v, actual_v = construction_v
                    if plan_v is not None:
                        entry = {'plan': plan_v, 'actual': actual_v}
                        if scope == 'GUE' and discs:
                            entry['disciplines'] = discs
                        scopes[scope] = entry

    return scopes


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


def extract_progress_hydro_paklay(pdf):
    """
    Pak Lay EPC weekly report.
    Report is primarily issue/design tracking; overall progress % may not exist in text.
    Tries: 'Physical Progress: Plan X% / Actual Y%' pattern.
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

    return {'plan': None, 'actual': None, 'disciplines': {}}


# ── Progress dispatcher ───────────────────────────────────────────────────────

def extract_progress(pdf_path, prj_id):
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
            return extract_progress_solar(pdf_path)
        elif extract_type == 'wind':
            return extract_progress_wind(pdf_path)
        elif extract_type == 'hydro_pakbeng':
            return extract_progress_hydro_pakbeng(pdf_path)
        elif extract_type == 'hydro_paklay':
            return extract_progress_hydro_paklay(pdf_path)
    except Exception as e:
        print(f"  [progress extract error] {e}")
    return {'plan': None, 'actual': None, 'disciplines': {}}


# ── Main entry point ──────────────────────────────────────────────────────────

def extract_from_pdf(pdf_path, prj_id=None):
    """
    Returns dict:
      { concerns: [...], activities: [...],
        plan: float|None, actual: float|None,
        disciplines: {disc: {plan, actual}} }
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

    # Extract progress %
    progress = extract_progress(pdf_path, prj_id) if prj_id else \
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
