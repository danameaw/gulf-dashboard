# ── extract.py ── Extract concerns, activities & progress % from PDF ───────
import sys, re
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

import pdfplumber
from patterns import (
    AUTO_EXTRACT_PROJECTS,
    IWTE_PAGE_INCLUDE, IWTE_PAGE_EXCLUDE, IWTE_ROWS,
    GMTP_PAGE_INCLUDE, GMTP_ROWS,
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
        v = float(s.strip())
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

            # ── Try table extraction first ───────────────────────────────────
            tables = page.extract_tables()
            for tbl in tables:
                if not tbl:
                    continue
                discs, overall = _table_rows_to_disc(tbl)
                if len(discs) >= 3:  # at least 3 disciplines found
                    plan, actual = overall
                    # Compute overall from disciplines if not found
                    if (plan is None or actual is None) and discs:
                        weights = {'Engineering': 0.10, 'Procurement': 0.40,
                                   'Construction': 0.40, 'Commissioning': 0.10}
                        wp = wa = wt = 0
                        for d, (p, a) in discs.items():
                            w = weights.get(d, 0)
                            wp += p * w; wa += a * w; wt += w
                        if wt > 0:
                            plan = round(wp / wt, 2)
                            actual = round(wa / wt, 2)
                    disciplines = {d: {'plan': p, 'actual': a}
                                   for d, (p, a) in discs.items()}
                    return {'plan': plan, 'actual': actual,
                            'disciplines': disciplines}

            # ── Fallback: text regex ─────────────────────────────────────────
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


# ── Progress dispatcher ───────────────────────────────────────────────────────

def extract_progress(pdf_path, prj_id):
    """
    Dispatch to the correct extractor based on project type.
    Returns { plan, actual, disciplines }.
    Projects whose S-curves are IMAGE-based return plan=None, actual=None.
    """
    extract_type = AUTO_EXTRACT_PROJECTS.get(prj_id)
    try:
        if extract_type == 'iwte':
            return extract_progress_iwte(pdf_path)
        elif extract_type == 'gmtp':
            return extract_progress_gmtp(pdf_path)
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

                # Concerns
                if _CONCERN_HEADERS.search(tl) and '%' not in text:
                    items = _extract_section(
                        text, _CONCERN_HEADERS,
                        stop_re=_ACTIVITY_HEADERS)
                    for it in items:
                        if it not in seen_c and len(it) > 10:
                            seen_c.add(it)
                            concerns.append(it)

                # iWTE-style "Concern (Delay)" section
                if 'concern (delay)' in tl or 'concern(delay)' in tl:
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

    return {
        'concerns':    concerns[:20],
        'activities':  activities[:20],
        'plan':        progress['plan'],
        'actual':      progress['actual'],
        'disciplines': progress['disciplines'],
    }
