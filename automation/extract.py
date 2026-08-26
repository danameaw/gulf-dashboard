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
    WIND_EXEC_PLAN_THEN_ACTUAL, WIND_EXEC_ACTUAL_THEN_PLAN,
    WIND_EXEC_DELTA, WIND_EXEC_INLINE, WIND_EXEC_ACTUAL_NEWLINE_PLAN,
    WIND_EXEC_PROGRESS_DELTA, WIND_EXEC_SCOPE_COL_NAMES,
    PAKBENG_PLAN_ACTUAL, PAKBENG_CUM_PLANNED, PAKBENG_CUM_ACTUAL,
    TTT_SCURVE_PAGE, TTT_AREA_ROW, TTT_OVERALL_ROW, TTT_PLAN_ACTUAL,
    TTT_DISC_CANON,
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

def _is_number(text):
    """True for a bare figure like '82.23' or '100.00' (no letters)."""
    return bool(re.fullmatch(r'\d{1,3}(?:\.\d{1,2})?\s*%?', (text or '').strip()))


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


# ── iWTE Executive Summary concerns ───────────────────────────────────────────
# The iWTE Executive Summary page is a two-column layout: a Plan/Actual
# progress table on the left, free-text sections (Key Achievement, Key Delay
# & Impact, Area of Concern, ...) on the right. pdfplumber's default
# extract_text() linearizes both columns onto the same lines, which
# corrupts wrapped bullets across the column boundary (a bullet's 2nd line
# gets merged with the unrelated left-column text on that same output line).
# _extract_right_column_text reconstructs just the right column from word
# positions so the "Area of Concern" section can be read cleanly.

_IWTE_AREA_OF_CONCERN = re.compile(r'area\s+of\s+concern\s*:?', re.I)
_IWTE_CONCERN_STOP    = re.compile(r'issue\s+to\s+request|ho\s+support', re.I)
_IWTE_NUM_BULLET      = re.compile(r'^\d+[\.\)]\s*')
# The KPI box and the neighbouring "Issue to request HO support" label share
# text lines with the concern items on these two-column slides; drop them
# rather than treating them as the end of the section.
_IWTE_HO_SUPPORT   = re.compile(r'issue\s+to\s+request\s+ho\s+support\s*:?.*$',
                                re.I)
_IWTE_FIELD_NOISE  = re.compile(
    r'^(\(?accumulative\)?|trir\b.*|total\s+man-?hours.*|n/?a\.?|none\.?'
    r'|[\d,.\s%]+)$', re.I)
_IWTE_FIELD_END    = re.compile(
    r'^(key\s+\w+|catch\s*[- ]?up|pre-?conditions|progress\b|construction\b'
    r'|commissioning\b|total\s+progress|manpower|organization|milestone)', re.I)
# "Area of Concern: Refer to page 15" / "AREA OF CONCERN : Ref. page 14"
_IWTE_CROSS_REF    = re.compile(r'^(refer|ref\.?)\s*(to\s*)?page\s*\d+', re.I)
_IWTE_CONCERN_TITLE = re.compile(r'^\s*areas?\s+of\s+concern\s*$', re.I | re.M)
# "Key Lookahead milestone:" - the Executive Summary field stating the work
# coming up, as opposed to the "Next Week Activities" slide which only ever
# lists safety routine.
_IWTE_LOOKAHEAD = re.compile(r'key\s+look\s*-?\s*ahead(?:\s+milestones?)?\s*:?', re.I)
# Items in these fields are numbered on some reports and dash-bulleted on
# others (GSE/TSE), so both start a new item.
_IWTE_ITEM_START = re.compile(r'^(?:\d+[.)]|[-\u2013\u2022\u25aa])\s*')
# Terminators for a field that is followed by other summary fields. The
# concern reader deliberately does not use this: "Issue to request HO
# support" sits ABOVE its items in reading order there, so it has to be
# stripped rather than stopped on.
_IWTE_NEXT_FIELD = re.compile(
    r'^(key\s+\w+|catch\s*[- ]?up|pre-?conditions|area\s+of\s+concern'
    r'|issue\s+to\s+request|progress\b|construction\b|commissioning\b'
    r'|total\s+progress|manpower|organization|milestone\s+tracking)', re.I)


def _page_lines(page, min_x0=None):
    """
    The page as visual lines: [{'top', 'x0', 'text'}], ordered top to bottom,
    words within a line ordered left to right. Keeping x0 is what allows a
    column to be isolated by where a heading actually sits, instead of by a
    fixed fraction of the page width that differs between contractors.

    min_x0 filters by WORD, not by line, which is the only thing that works
    on these slides: a wrapped concern continuation shares its visual line
    with the KPI box on the left, so the line begins at x=53 with "TRIR 0"
    and only then carries the text belonging to the right-hand column.
    Judging the line by its leftmost word throws that continuation away.
    """
    words = page.extract_words()
    if min_x0 is not None:
        words = [w for w in words if w['x0'] >= min_x0]
    lines = []
    for w in sorted(words, key=lambda w: (w['top'], w['x0'])):
        for ln in lines:
            if abs(ln['top'] - w['top']) <= 3:
                ln['words'].append(w)
                break
        else:
            lines.append({'top': w['top'], 'words': [w]})
    out = []
    for ln in sorted(lines, key=lambda l: l['top']):
        ws = sorted(ln['words'], key=lambda w: w['x0'])
        out.append({'top': ln['top'], 'x0': ws[0]['x0'],
                    'text': ' '.join(w['text'] for w in ws)})
    return out


def _extract_right_column_text(page, split_frac=0.45):
    """Reconstruct the right-hand column of a two-column page from word
    x-positions: words with x0 past split_frac*page width are grouped into
    lines by y-position (clustered within a few px), then lines are sorted
    top-to-bottom and words within each line left-to-right."""
    width = page.width
    right_words = [w for w in page.extract_words() if w['x0'] >= width * split_frac]
    right_words.sort(key=lambda w: w['top'])
    lines = []
    for w in right_words:
        for line in lines:
            if abs(line['top'] - w['top']) <= 3:
                line['words'].append(w)
                break
        else:
            lines.append({'top': w['top'], 'words': [w]})
    lines.sort(key=lambda l: l['top'])
    return '\n'.join(
        ' '.join(w['text'] for w in sorted(l['words'], key=lambda w: w['x0']))
        for l in lines)


def _iwte_field_items(lines, start_idx, col_x, stop_re=None,
                     item_re=None, strip_ho=True):
    """
    Collect the list items of an Executive Summary field: the lines below
    `start_idx` that belong to the column starting at `col_x`, grouped into
    items on bullet or number starts, with wrapped lines folded into the item
    above. Returns (items, cross_ref) - cross_ref is True when the field only
    points at another page ("Refer to page 15").
    """
    stop_re = stop_re or _IWTE_FIELD_END
    item_re = item_re or _IWTE_NUM_BULLET
    items, current, cross_ref = [], None, False
    for ln in lines[start_idx + 1:]:
        if ln['x0'] < col_x:
            continue
        text = ln['text']
        if strip_ho:
            text = _IWTE_HO_SUPPORT.sub('', text)
        text = text.strip()
        if not text or _IWTE_FIELD_NOISE.match(text):
            continue
        if _IWTE_CROSS_REF.match(text):
            cross_ref = True
            continue
        if stop_re.match(text):
            break
        if item_re.match(text):
            if current:
                items.append(current)
            current = item_re.sub('', text).strip()
        elif current is not None:
            current += ' ' + text
        else:
            current = text
    if current:
        items.append(current)
    return items, cross_ref


def _iwte_exec_summary_page(doc, anchor_re):
    """
    The Executive Summary slide, plus the index of `anchor_re` on it and the
    column-filtered lines to read from.
    A report's Table of Contents lists "Executive Summary" as an entry, so
    matching on that phrase alone finds the contents page first - which is
    why this requires the field itself to be present before accepting a page.
    Returns (lines, idx, anchor_x) or None.
    """
    for page in doc.pages[:8]:
        text = page.extract_text() or ''
        if 'executive summary' not in text.lower() or not anchor_re.search(text):
            continue
        base = _page_lines(page)
        idx = next((i for i, ln in enumerate(base)
                    if anchor_re.search(ln['text'])), None)
        if idx is None:
            continue
        anchor_x = base[idx]['x0']
        # Re-read keeping only words from the heading's column onward, so a
        # wrapped continuation is not discarded for sharing its visual line
        # with the left-hand KPI box.
        lines = _page_lines(page, min_x0=anchor_x - 10)
        idx = next((i for i, ln in enumerate(lines)
                    if anchor_re.search(ln['text'])), None)
        if idx is None:
            continue
        return page, lines, idx, anchor_x
    return None


def extract_iwte_lookahead(pdf_path):
    """
    The Executive Summary's "Key Lookahead milestone" field - the work coming
    up next period. Returns a list, or None when the field is absent so the
    caller can tell "nothing planned stated" from "could not read it".
    """
    try:
        with pdfplumber.open(pdf_path) as doc:
            found = _iwte_exec_summary_page(doc, _IWTE_LOOKAHEAD)
            if not found:
                return None
            _, lines, idx, anchor_x = found
            head = _IWTE_LOOKAHEAD.sub('', lines[idx]['text'], count=1).strip(' :')
            # The heading line can also carry a stray KPI figure from the box
            # to its left ("TRIR 0"), and a plural label spelling used to
            # leave a lone "s" behind. Only keep a remainder substantial
            # enough to be a real item.
            head = re.sub(r'^\s*(trir|tri)\s*[\d.]*\s*', '', head, flags=re.I)
            head = _IWTE_ITEM_START.sub('', head).strip()
            items, _ = _iwte_field_items(
                lines, idx, anchor_x - 10, stop_re=_IWTE_NEXT_FIELD,
                item_re=_IWTE_ITEM_START, strip_ho=False)
            if len(head) >= 10 and not _IWTE_NEXT_FIELD.match(head):
                items.insert(0, head)
            return [it for it in items
                    if it and it.upper() not in ('N/A', 'NA', 'NONE')]
    except Exception as e:
        print(f"  [iwte lookahead error] {e}")
    return None


def extract_iwte_concerns_from_section(doc):
    """
    The report's own "Areas of Concern" section, used when the Executive
    Summary field is a cross-reference to it. Section divider slides carry
    only the title and a page number; the content slides list numbered items
    whose first line states the issue ("1. Problem : Electric Pole obstruct
    Entrance") followed by a long, often Thai "Action :" narrative that is
    too much for a dashboard line. Take the statement, leave the narrative.
    Returns a list, empty when the section says "None".
    """
    items = []
    for page in doc.pages:
        text = page.extract_text() or ''
        if not _IWTE_CONCERN_TITLE.search(text):
            continue
        for raw in text.split('\n')[1:]:
            line = raw.strip()
            if not line or _PAGE_NUM.match(line):
                continue
            if line.upper() in ('NONE', 'N/A', 'NA'):
                continue
            if _IWTE_NUM_BULLET.match(line):
                items.append(_IWTE_NUM_BULLET.sub('', line).strip())
    return [it for it in items if len(it) > 5]


def extract_iwte_concerns_from_exec_summary(pdf_path):
    """
    Read the "Area of Concern" field of the iWTE Executive Summary slide.
    The column is located from the heading's own x-position (contractors
    place it anywhere from 0.42 to 0.5 of the page width) and continuation
    lines are allowed a little to its left, since wrapped text is often
    indented outside the heading's own start.
    Returns a list of concerns, or None when the field could not be found at
    all - the caller uses that to distinguish "no concerns this week" from
    "this reader did not work", and only falls back to the generic scan for
    the latter.
    """
    try:
        with pdfplumber.open(pdf_path) as doc:
            for page in doc.pages[:8]:
                text = page.extract_text() or ''
                if 'executive summary' not in text.lower():
                    continue
                # Pass 1 locates the heading; pass 2 re-reads the page
                # keeping only words from the heading's column onward, so
                # wrapped continuations survive without the left column's
                # text being spliced in front of them.
                idx = next((i for i, ln in enumerate(_page_lines(page))
                            if _IWTE_AREA_OF_CONCERN.search(ln['text'])), None)
                if idx is None:
                    continue
                anchor_x = _page_lines(page)[idx]['x0']
                lines = _page_lines(page, min_x0=anchor_x - 10)
                idx = next((i for i, ln in enumerate(lines)
                            if _IWTE_AREA_OF_CONCERN.search(ln['text'])), None)
                if idx is None:
                    continue

                # The heading line can carry the whole field inline
                # ("Area of Concern: Refer to page 15").
                head = _IWTE_AREA_OF_CONCERN.sub('', lines[idx]['text'], count=1)
                head = _IWTE_HO_SUPPORT.sub('', head).strip(' :')
                items, cross_ref = _iwte_field_items(
                    lines, idx, anchor_x - 10)
                if head:
                    if _IWTE_CROSS_REF.match(head):
                        cross_ref = True
                    elif head.upper() not in ('N/A', 'NA', 'NONE'):
                        items.insert(0, _IWTE_NUM_BULLET.sub('', head).strip())

                if not items and cross_ref:
                    return extract_iwte_concerns_from_section(doc)
                return [it for it in items
                        if it and it.upper() not in ('N/A', 'NA', 'NONE')]
    except Exception as e:
        print(f"  [iwte concerns error] {e}")
    return None


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


# The 'Plan / Actual :' label and its values can be split across lines, and a
# y-axis gridline label (e.g. "60%") sometimes lands between them in the text
# layer — "Plan / Actual :\n60% 24.02% / 28.79% (4.76%)". Allow a short gap
# (non-greedy, capped) so the stray text is skipped and the real
# "plan% / actual%" pair is captured.
#
# The gap has to be generous: on pages that also carry a discipline table
# (3.2.8 Construction in WR-054) a whole table row lands between the label and
# its values — "Plan / Actual : Mech 2.02 0.00 0.00 0.00" then "18.81% /
# 23.14%" on the next line. A 20-char window missed it and the project lost
# Construction entirely. Only the "X% / Y%" slash pair can satisfy the tail,
# and table cells carry no percent signs, so a wider window costs no precision.
_GMTP_PLAN_ACTUAL_CALLOUT = re.compile(
    r'plan\s*/\s*actual\s*:?[\s\S]{0,120}?'
    r'(\d{1,3}(?:\.\d{1,2})?)\s*%\s*/\s*(\d{1,3}(?:\.\d{1,2})?)\s*%', re.I)


# Every "3.2 S-Curve" page carries a numbered heading ("3.2.6 Engineering")
# next to one clean "Plan / Actual :" callout. Reading those headings is far
# more robust than parsing the 3.1.1 Overall Progress table, whose text layer
# some exports scramble to one character per line (unrecoverable by any
# regex — which is why GMTP reported discs=0 every week since W30), and it
# recovers the per-area breakdown (LNG Tank / BOP & Utility / Marine) that
# section 3.1.1 has but nothing was surfacing to the dashboard at all.
_GMTP_SCURVE_HEADING = re.compile(r'^\s*3\.2\.(\d+)\s+(\S.*?)\s*$', re.M)
_GMTP_DISC_TITLES = {
    'engineering':   'Engineering',
    'procurement':   'Procurement',
    'construction':  'Construction',
    'commissioning': 'Commissioning',
}


def _gmtp_section_kind(title):
    """
    Classify a 3.2.x section heading: 'overall', a discipline name, or None
    (= a project area/package such as "LNG Tank Area").
    Matched on the heading with any parenthetical and punctuation stripped,
    and only as a whole word — a future area named e.g. "Marine Construction
    Area" must not be mistaken for the project-wide Construction discipline.
    """
    key = re.sub(r'\([^)]*\)', ' ', title).strip().lower()
    key = re.sub(r'[^a-z ]+', ' ', key).strip()
    if 'overall' in key:
        return 'overall'
    return _GMTP_DISC_TITLES.get(key)


# The area slides letter-space their small summary table, so words and cell
# boundaries are both lost to the text layer. Characters keep their true
# x-positions though, and the gap between two fields is far wider than the gap
# inside a letter-spaced word - so the fields can be recovered geometrically.
_GMTP_MATRIX_HEADER = re.compile(r'A\s*r\s*e\s*a\s+D\s*i\s*s\s*c', re.I)
_GMTP_DISC_WORD = re.compile(
    r'(engineering|procurement|construction|commissioning|overall)', re.I)
# Var must equal Act - Plan; anything further off than this is not a row we
# have understood, and is dropped rather than reported.
_GMTP_MATRIX_TOLERANCE = 0.06


def _split_char_run(chars, slack=1.2):
    """
    Group characters into fields by horizontal gap. Letter-spacing inside a
    word leaves a near-zero (often slightly negative) gap; a column boundary
    leaves tens of points. Splitting at anything wider than the row's median
    gap plus `slack` separates the two reliably.
    """
    chars = sorted(chars, key=lambda c: c['x0'])
    if not chars:
        return []
    gaps = [chars[i + 1]['x0'] - chars[i]['x1'] for i in range(len(chars) - 1)]
    median = sorted(gaps)[len(gaps) // 2] if gaps else 0.0
    groups, current = [], [chars[0]]
    for i in range(len(chars) - 1):
        if gaps[i] > median + slack:
            groups.append(current)
            current = []
        current.append(chars[i + 1])
    groups.append(current)
    return [''.join(c['text'] for c in g).strip() for g in groups]


def _y_bands(chars, tolerance=4):
    """Group characters into visual rows by their top coordinate."""
    bands = []
    for c in sorted(chars, key=lambda c: c['top']):
        for band in bands:
            if abs(band['top'] - c['top']) <= tolerance:
                band['chars'].append(c)
                break
        else:
            bands.append({'top': c['top'], 'chars': [c]})
    return sorted(bands, key=lambda b: b['top'])


def _gmtp_matrix_rows(page):
    """
    Read the "Area | Discipline | W.F | Plan | Act. | Var." table on one
    section-3.2 slide. Returns {discipline: {'plan', 'actual', 'weight'}},
    including an 'Overall' entry when the table carries one. Rows whose Var.
    column does not agree with Act - Plan are discarded.
    """
    header = next((ln for ln in _page_lines(page)
                   if _GMTP_MATRIX_HEADER.search(ln['text'])), None)
    if header is None:
        return {}

    col_x = header['x0'] - 20
    body = [c for c in page.chars
            if c['x0'] >= col_x and c['top'] > header['top'] + 2]

    out = {}
    for band in _y_bands(body):
        fields = _split_char_run(band['chars'])
        label = ' '.join(f for f in fields if not _is_number(f))
        m = _GMTP_DISC_WORD.search(label)
        if not m:
            continue
        nums = [_parse_pct(f) for f in fields if _is_number(f)]
        nums = [n for n in nums if n is not None]
        if len(nums) < 4:
            continue
        weight, plan, actual, var = nums[:4]
        if abs(abs(actual - plan) - var) > _GMTP_MATRIX_TOLERANCE:
            print(f"  [gmtp matrix] {label.strip()}: dropped, "
                  f"Var {var} does not match {actual} - {plan}")
            continue
        out[m.group(1).capitalize()] = {'plan': plan, 'actual': actual,
                                       'weight': weight}
    return out


def _gmtp_scurve_sections(pdf):
    """
    Walk the section-3.2 S-Curve pages and read each one's "Plan / Actual :"
    callout, keyed by its heading.
    Returns (overall_plan, overall_actual, disciplines, areas) where
    disciplines/areas are {name: {'plan': float, 'actual': float}}.
    """
    overall_plan = overall_actual = None
    discs, areas = {}, {}
    try:
        with pdfplumber.open(pdf) as doc:
            for page in doc.pages:
                text = page.extract_text() or ''
                if '3.2 s-curve' not in text.lower():
                    continue
                mh = _GMTP_SCURVE_HEADING.search(text)
                mv = _GMTP_PLAN_ACTUAL_CALLOUT.search(text)
                if not mh or not mv:
                    continue
                plan, actual = _parse_pct(mv.group(1)), _parse_pct(mv.group(2))
                if plan is None or actual is None:
                    continue
                title = mh.group(2).strip()
                kind  = _gmtp_section_kind(title)
                if kind == 'overall':
                    overall_plan, overall_actual = plan, actual
                elif kind:
                    discs[kind] = {'plan': plan, 'actual': actual}
                else:
                    area = {'plan': plan, 'actual': actual}
                    # The area slides also carry their own EPCC split with
                    # weight factors; the dashboard renders a scope's nested
                    # disciplines on drill-down, so keep them.
                    matrix = _gmtp_matrix_rows(page)
                    matrix.pop('Overall', None)   # already have it above
                    if matrix:
                        area['disciplines'] = {
                            name: {'plan': v['plan'], 'actual': v['actual'],
                                   'wf': v['weight']}
                            for name, v in matrix.items()}
                    areas[title] = area
    except Exception as e:
        print(f"  [gmtp s-curve error] {e}")
    return overall_plan, overall_actual, discs, areas


def _gmtp_overall_scurve(pdf):
    """
    The '3.2 S-Curve' / '3.2.1 Overall Progress' page prints a clean
    'Plan / Actual : XX.XX% / YY.YY%' callout for the whole project. Prefer
    reading that directly over the '3.1.1 Overall Progress' per-package
    table (_parse_gmtp_overall_row): some report exports scramble that
    table's text layer to one character per line, which no regex can
    recover. The per-area S-curve pages (3.2.2, 3.2.3, ...) use the same
    callout format but also list a discipline breakdown
    (Engineering/Procurement/...) — skip those so this only matches the
    project-wide page.
    """
    try:
        with pdfplumber.open(pdf) as doc:
            for page in doc.pages:
                text = page.extract_text() or ''
                tl = text.lower()
                if 's-curve' not in tl or 'overall progress' not in tl:
                    continue
                if 'engineering' in tl:
                    continue
                m = _GMTP_PLAN_ACTUAL_CALLOUT.search(text)
                if m:
                    return _parse_pct(m.group(1)), _parse_pct(m.group(2))
    except Exception as e:
        print(f"  [gmtp overall error] {e}")
    return None, None


def extract_progress_gmtp(pdf):
    """
    GMTP/LNG report: find the Overall Progress table (text-based).
    Looks for page with discipline breakdown and plan/actual headers.
    Returns { plan, actual, disciplines: {disc: {plan, actual}} }.
    """
    # Section 3.2's per-section callouts first — they carry the whole picture
    # (project overall + all four disciplines + every project area) in the one
    # format that has never failed to parse.
    s_plan, s_actual, s_discs, s_areas = _gmtp_scurve_sections(pdf)
    if s_plan is not None and (s_discs or s_areas):
        result = {'plan': s_plan, 'actual': s_actual, 'disciplines': s_discs}
        if s_areas:
            # Scopes are the project's real areas (LNG Tank / BOP & Utility /
            # Marine). The project-wide EPCC breakdown rides along in
            # 'disciplines' above and the dashboard renders it as its own
            # table, so there is no synthetic 'overall' scope here - that
            # only duplicated the discipline table's Overall row while
            # reading like a fourth area.
            result['scopes'] = dict(s_areas)
        return result

    scurve_plan, scurve_actual = _gmtp_overall_scurve(pdf)

    with pdfplumber.open(pdf) as doc:
        for page in doc.pages:
            text = page.extract_text() or ''
            tl = text.lower()

            if not GMTP_PAGE_INCLUDE.search(tl):
                continue
            if not any(d.lower() in tl for d in
                       ['engineering', 'procurement', 'construction']):
                continue
            # Section 3.2's per-area S-curve pages (BOP & Utility Area, Marine
            # Area, ...) also mention every discipline, but their breakdown is
            # scoped to that one area, not the whole project — mistaking it
            # for the 3.1.1 Overall Progress table's project-wide breakdown
            # silently substitutes one area's numbers for the project's (see
            # _table_rows_to_disc, which additionally misreads that page's
            # noisy weekly time-series cells on top of that scope mismatch).
            # Only the non-S-curve 3.1.1 page is the right shape for the
            # methods below.
            if 's-curve' in tl:
                continue

            # ── Try 'Overall Progress' per-package breakdown row first ───────
            plan, actual, discs = _parse_gmtp_overall_row(text)
            if plan is not None and actual is not None:
                p, a = (scurve_plan, scurve_actual) if scurve_plan is not None else (plan, actual)
                return {'plan': p, 'actual': a, 'disciplines': discs}

            # ── Try table extraction ─────────────────────────────────────────
            tables = page.extract_tables()
            for tbl in tables:
                if not tbl:
                    continue
                discs, overall = _table_rows_to_disc(tbl)
                if len(discs) >= 2:
                    plan, actual = overall
                    p, a = (scurve_plan, scurve_actual) if scurve_plan is not None else (plan, actual)
                    disciplines = {d: {'plan': p2, 'actual': a2}
                                   for d, (p2, a2) in discs.items()}
                    return {'plan': p, 'actual': a,
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
                p, a = (scurve_plan, scurve_actual) if scurve_plan is not None else (overall_plan, overall_actual)
                return {'plan': p, 'actual': a, 'disciplines': discs}

    if scurve_plan is not None:
        return {'plan': scurve_plan, 'actual': scurve_actual, 'disciplines': {}}
    return {'plan': None, 'actual': None, 'disciplines': {}}


# ── GMTP concerns & next-week activities ──────────────────────────────────────
#
# The generic whole-PDF bullet scan in extract_from_pdf is a bad fit for this
# report: it matched the SHE NCR table and the "This Week Key Milestone" list
# as "concerns", and the safety/TBT event log as "activities", while the two
# sections that actually say what the project is worried about and what it
# will build next week were never read at all. Both are cleanly structured, so
# read them directly.
#
#   Section 10 "KEY ISSUES AND CONCERNS" — one item per '○' heading, followed
#   by free-text description lines and '-' detail bullets, closed by a
#   "<no.> <owner> <status>" row (e.g. "1 Eng. Info.").
#
#   Sections 6.3 / 6.4 / 6.5 "Main Activities in Next Week (<Area>)" — '•'
#   bullets under numbered sub-headings ("6.4.2 Piperack – B, C, D, E").

# Pak Beng shares this report's section layout but not its exact wording, and
# loosening these patterns to cover both is what broke GMTP: allowing numbered
# concern items pulled unrelated numbered lines off its Key Issues pages, and
# making the activities heading's area suffix optional let that reader match
# GMTP's own "5.7 Main Activities in Next Week" under Procurement. Each
# project passes its own pair; neither is widened for the other.
_GMTP_CONCERN_PAGE    = re.compile(r'key\s+issues\s+and\s+concerns', re.I)
_GMTP_CONCERN_ITEM    = re.compile(r'^[○●◦o]\s+(\S.*)$')
_PAKBENG_CONCERN_PAGE = re.compile(r'this\s+week\s+issues\s+and\s+concerns', re.I)
_PAKBENG_CONCERN_ITEM = re.compile(r'^\d+\s*\.\s*(\S.*)$')
# Pak Beng repeats each concern in Chinese; the English above it is enough.
_CJK_LINE = re.compile(r'[\u3000-\u9fff\uff00-\uffef]')
_GMTP_CONCERN_CLOSE  = re.compile(r'^\d+\s+\w+\.?\s+\w+\.?\s*$')
_GMTP_CONCERN_NOISE  = re.compile(
    r'^(status|no\.|<\s*open|\d+\s*$|key\s+issues)', re.I)

# GMTP's construction sections suffix the heading with their area ("... in
# Next Week (BOP Area)"). Requiring the suffix is what keeps this off the
# report's other "Main Activities in Next Week" sections, notably 5.7 under
# Procurement. Pak Beng's single unsuffixed section has its own pattern below.
_GMTP_NEXT_WEEK_PAGE = re.compile(
    r'main\s+activities\s+in\s+next\s+week\s*\(([^)]*)\)', re.I)
_PAKBENG_NEXT_WEEK_PAGE = re.compile(
    r'main\s+activities\s+in\s+next\s+week', re.I)
_GMTP_SUBHEADING     = re.compile(r'^\s*6\.\d+(?:\.\d+)*\s+(\S.*?)\s*$')
_GMTP_BULLET         = re.compile(r'^\s*[•▪·]\s*(\S.*)$')


def extract_gmtp_concerns_checked(pdf, max_items=20, page_re=None,
                                  item_re=None):
    """
    Section 10 "KEY ISSUES AND CONCERNS" → one line per item, rendered as
    "<heading> — <first description line>" so the dashboard shows what the
    issue is, not just its title. Returns [] if the section is absent.

    page_re/item_re let Pak Beng reuse this: its section is titled "THIS WEEK
    ISSUES AND CONCERNS" and numbers its items rather than bulleting them.
    """
    page_re = page_re or _GMTP_CONCERN_PAGE
    item_re = item_re or _GMTP_CONCERN_ITEM
    items = []
    # Whether the section was located at all. A report that states
    # "None." (Pak Beng 066) yields an empty list, and the caller has to
    # be able to tell that from "this reader found nothing to read" -
    # otherwise the generic scan's table wreckage keeps winning against
    # a report that genuinely has no concerns this week.
    found_section = False
    try:
        with pdfplumber.open(pdf) as doc:
            for page in doc.pages:
                text = page.extract_text() or ''
                if not page_re.search(text):
                    continue
                found_section = True
                # An item reads:
                #   ○ <heading>
                #   <description, wrapped over 1-2 lines>      ← preferred
                #   - <supporting detail bullet>               ← fallback
                #   ...
                #   <no.> <owner> <status>
                # Wrapped continuation lines resume AFTER the bullets too, so
                # collecting plain lines indiscriminately splices sentence
                # fragments together ("and gondola for LNG tanks. to be
                # possible."). Take only the plain lines that precede the
                # first bullet, and fall back to the first bullet for items
                # whose description is entirely bulleted.
                heading, lead, first_bullet, seen_bullet = None, [], None, False

                def _flush(heading, lead, first_bullet):
                    if not heading:
                        return None
                    body = ' '.join(lead).strip() or (first_bullet or '')
                    return f"{heading} — {body}" if body else heading

                for raw in text.splitlines():
                    line = raw.strip()
                    # Pak Beng's section title ("8.THIS WEEK ISSUES AND
                    # CONCERNS") is itself a numbered line, so skip it before
                    # the item test or it becomes the first concern.
                    if page_re.search(line):
                        continue
                    m = item_re.match(line)
                    if m:
                        done = _flush(heading, lead, first_bullet)
                        if done:
                            items.append(done)
                        heading, lead, first_bullet, seen_bullet =                             m.group(1).strip(), [], None, False
                        continue
                    if heading is None:
                        continue
                    if (not line or _GMTP_CONCERN_NOISE.match(line)
                            or _GMTP_CONCERN_CLOSE.match(line)):
                        continue
                    if _CJK_LINE.search(line):
                        continue
                    # Layout artefacts: a line made only of dashes or other
                    # punctuation carries no text ("- -" sits between Pak
                    # Beng's concern title and its description).
                    if not re.search(r'[A-Za-z]{3}', line):
                        continue
                    if line.startswith(('-', '–', '•', '‑')):
                        seen_bullet = True
                        if first_bullet is None:
                            first_bullet = line.lstrip('-–• ').strip()
                        continue
                    if not seen_bullet and sum(len(d) for d in lead) < 220:
                        lead.append(line)
                done = _flush(heading, lead, first_bullet)
                if done:
                    items.append(done)
    except Exception as e:
        print(f"  [gmtp concerns error] {e}")

    out, seen = [], set()
    for it in items:
        if it not in seen and len(it) > 10:
            seen.add(it)
            out.append(it)
    return out[:max_items], found_section


def extract_gmtp_concerns(pdf, max_items=20, page_re=None, item_re=None):
    """Concern list only; see extract_gmtp_concerns_checked for the flag."""
    return extract_gmtp_concerns_checked(pdf, max_items, page_re, item_re)[0]


def extract_gmtp_next_week_activities(pdf, max_items=20):
    """
    Sections 6.3-6.5 "Main Activities in Next Week (<Area>)" → one line per
    '•' bullet, prefixed with its area and sub-heading
    ("[BOP Area] Piling Work: Pile Driving at Process Area"). Returns [] if
    no next-week section is present.
    """
    items = []
    try:
        with pdfplumber.open(pdf) as doc:
            for page in doc.pages:
                text = page.extract_text() or ''
                mp = _GMTP_NEXT_WEEK_PAGE.search(text)
                if not mp:
                    continue
                area = (mp.group(1) or '').strip()
                sub = None
                for raw in text.splitlines():
                    line = raw.strip()
                    mb = _GMTP_BULLET.match(line)
                    if mb:
                        body = mb.group(1).strip()
                        label = (f"[{area}] " if area else '') \
                            + (f"{sub}: " if sub else '')
                        items.append(label + body)
                        continue
                    mh = _GMTP_SUBHEADING.match(line)
                    if mh and not _GMTP_NEXT_WEEK_PAGE.search(line):
                        sub = mh.group(1).strip()
    except Exception as e:
        print(f"  [gmtp activities error] {e}")

    out, seen = [], set()
    for it in items:
        if it not in seen and len(it) > 10:
            seen.add(it)
            out.append(it)
    return out[:max_items]


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
            # An Exec Summary cell does not always carry an "Overall progress"
            # line. NWT3's GUE cell dropped it in Week 47 and the reader fell
            # back to that scope's Construction figure — a genuinely different,
            # lower number (61.10% against plan 100%, where Overall was 69.37%
            # against plan 90%). Published as the project's overall it looked
            # like progress had gone backwards week on week. The scope's own
            # section page still states Overall, so prefer that over the proxy.
            if exec_entry.pop('overall_is_proxy', False) and tw_entry                     and tw_entry.get('plan') is not None:
                exec_entry['plan']   = tw_entry['plan']
                exec_entry['actual'] = tw_entry['actual']
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
                    overall_is_proxy = False
                    if plan_v is None and construction_v is not None:
                        plan_v, actual_v = construction_v
                        overall_is_proxy = True
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
                        if overall_is_proxy:
                            # Caller decides whether a better Overall exists.
                            entry['overall_is_proxy'] = True
                        if scope == 'GUE' and discs:
                            entry['disciplines'] = discs
                        scopes[scope] = entry

    return scopes


# Cells in this table are bulleted with whichever glyph the contractor's
# template uses. An unrecognised marker does not fail loudly - it just folds
# every bullet in the cell into one run-on item - so the set is deliberately
# broad.
_EXEC_CELL_BULLET = re.compile(r'^[\u2022\u25aa\u25a1\u2751\u274f\u27a2\u25cb\u25e6*-]\s*')
# A cell can carry the slide's page number in its corner ("N/A 4"), so strip a
# trailing bare number before deciding whether the cell is empty.
_EXEC_CELL_EMPTY = ('N/A', 'NA', 'N.A', '-', 'NONE')


def _exec_summary_row_items(pdf, row_key, header_keys,
                            require_site_progress=False, page_limit=10):
    """
    Read one row of the Solar/Wind Executive Summary table and return
    (items, found_row). Each item is prefixed with its scope column header,
    e.g. "[PVfarm(GUE)] Installation Pump water tank".

    row_key is matched against the row's text with spaces removed: these
    cells lose their spacing in extraction ("Concernneedmanagement attention",
    "Keyhighlight activitythisweek").

    found_row distinguishes "the row is there and says nothing this week" from
    "the row could not be located", which callers need in order to decide
    whether falling back to the generic whole-PDF scan is right. Keeping stale
    text from elsewhere in the report over a genuine "no concerns" is not.
    """
    items_out = []
    found_row = False
    with pdfplumber.open(pdf) as doc:
        for page in doc.pages[:page_limit]:
            text = (page.extract_text() or '').lower()
            has_exec = 'executive summary' in text
            has_site = 'site progress' in text
            if require_site_progress:
                if not (has_exec and has_site):
                    continue
            elif not (has_exec or has_site):
                continue

            for tbl in page.extract_tables():
                if not tbl:
                    continue
                header_row = target_row = None
                for row in tbl:
                    if not _is_multicol_row(row):
                        continue
                    joined = ' '.join(str(c or '') for c in row).lower()
                    if header_row is None and any(k in joined for k in header_keys):
                        header_row = row
                    if row_key in joined.replace(' ', ''):
                        target_row = row

                if header_row is None or target_row is None:
                    continue
                found_row = True

                for ci, cell in enumerate(target_row):
                    header = (str(header_row[ci])
                              if ci < len(header_row) and header_row[ci] else None)
                    if not header or not cell or header.strip().lower() == 'topic':
                        continue
                    header_label = ' '.join(header.split())
                    # One bullet's text wraps over several lines in these
                    # narrow columns: only a bullet marker starts a new item,
                    # everything else continues the one above.
                    items, current = [], None
                    for line in str(cell).split('\n'):
                        line = line.strip()
                        if not line:
                            continue
                        if _EXEC_CELL_BULLET.match(line):
                            if current:
                                items.append(current)
                            current = _EXEC_CELL_BULLET.sub('', line)
                        elif current is not None:
                            current += ' ' + line
                        else:
                            current = line
                    if current:
                        items.append(current)
                    for it in items:
                        it_check = re.sub(r'\s*\d+\s*$', '', it).strip()
                        if not it_check or it_check.upper() in _EXEC_CELL_EMPTY:
                            continue
                        items_out.append(f'[{header_label}] {it}')
    return items_out, found_row


_SOLAR_SCOPE_HEADERS = ['gue', 'pvfarm', 'pv farm', 'siemens']
_WIND_SCOPE_HEADERS  = ['pcz', 'siemens', 'goldwind', 'gold wind', 'tsa']


def extract_solar_concerns_from_exec_summary(pdf):
    """The 'Concern need management attention' row of a Solar report."""
    items, _ = _exec_summary_row_items(pdf, 'concern', _SOLAR_SCOPE_HEADERS)
    return items


def extract_solar_highlights_from_exec_summary(pdf):
    """The 'Key highlight activity this week' row of a Solar report."""
    items, _ = _exec_summary_row_items(pdf, 'keyhighlight', _SOLAR_SCOPE_HEADERS)
    return items


# ── Wind Executive Summary ────────────────────────────────────────────────────

def _parse_wind_exec_cell(cell_text):
    """
    Try each known Wind Executive Summary cell phrasing in turn (see
    WIND_EXEC_* patterns in patterns.py for what each project's report
    actually looks like) and return (plan, actual), preferring an explicit
    Overall line over a Construction line as scope-level proxy — same
    convention as the Solar Exec Summary parser. Returns (None, None) if no
    pattern matched (e.g. a plain "N/A" cell, or Wayu's ambiguous TSA column
    with two different "plan" figures, which is deliberately left unparsed).
    """
    overall = construction = None

    for m in WIND_EXEC_PLAN_THEN_ACTUAL.finditer(cell_text):
        label = m.group(1).lower()
        p, a = _parse_pct(m.group(2)), _parse_pct(m.group(3))
        if p is None or a is None:
            continue
        if 'construction' in label:
            construction = (p, a)
        else:
            overall = (p, a)

    if overall is None:
        m = WIND_EXEC_ACTUAL_THEN_PLAN.search(cell_text)
        if m:
            a, p = _parse_pct(m.group(1)), _parse_pct(m.group(2))
            if a is not None and p is not None:
                overall = (p, a)

    if overall is None:
        for m in WIND_EXEC_ACTUAL_NEWLINE_PLAN.finditer(cell_text):
            label = m.group(1).lower()
            a, p = _parse_pct(m.group(2)), _parse_pct(m.group(3))
            if a is None or p is None:
                continue
            if 'construction' in label:
                construction = (p, a)
            else:
                overall = (p, a)

    if overall is None:
        for m in WIND_EXEC_DELTA.finditer(cell_text):
            label = m.group(1).lower()
            a = _parse_pct(m.group(2))
            direction = m.group(3).lower()
            delta = _parse_pct(m.group(4))
            if a is None or delta is None:
                continue
            p = round(a + delta, 2) if direction == 'delayed' else round(a - delta, 2)
            if 'construction' in label:
                if construction is None:
                    construction = (p, a)
            elif overall is None:
                overall = (p, a)

    if overall is None and construction is None:
        m = WIND_EXEC_INLINE.search(cell_text)
        if m:
            a, p = _parse_pct(m.group(1)), _parse_pct(m.group(2))
            if a is not None and p is not None:
                overall = (p, a)

    if overall is None and construction is None:
        m = WIND_EXEC_PROGRESS_DELTA.search(cell_text)
        if m:
            a = _parse_pct(m.group(1))
            try:
                delta = float(m.group(2))  # signed — may be negative, _parse_pct rejects those
            except ValueError:
                delta = None
            if a is not None and delta is not None:
                overall = (round(a - delta, 2), a)

    if overall is not None:
        return overall
    if construction is not None:
        return construction
    return None, None


def _extract_wind_from_exec_summary(pdf):
    """
    Wind Executive Summary page: one column per scope (CBOP, TSA, Substation,
    115kV T/L, and sometimes a project-specific extra column like AL1's
    "Pre-NTP" or Alpha 2's "Terminal Substation"). This is the authoritative
    source going forward — the "Overall Progress S-Curve" page-scan below
    misses scopes entirely for some projects.
    Returns { scope_name: {plan, actual} }.
    """
    scopes = {}
    with pdfplumber.open(pdf) as doc:
        for page in doc.pages[:8]:
            text = page.extract_text() or ''
            if 'executive summary' not in text.lower() or 'site progress' not in text.lower():
                continue
            for tbl in page.extract_tables():
                if not tbl:
                    continue
                header_row = None
                site_progress_row = None
                for row in tbl:
                    if not _is_multicol_row(row):
                        continue
                    joined = ' '.join(str(c or '') for c in row).lower()
                    if header_row is None and any(
                            k in joined for k in ['pcz', 'siemens', 'goldwind', 'gold wind', 'tsa']):
                        header_row = row
                    if 'site progress' in joined:
                        site_progress_row = row

                if header_row is None or site_progress_row is None:
                    continue

                col_scope = {}
                for ci, cell in enumerate(header_row):
                    if not cell:
                        continue
                    cl = str(cell).lower()
                    for key, scope_name in WIND_EXEC_SCOPE_COL_NAMES.items():
                        if key in cl:
                            col_scope[ci] = scope_name
                            break

                for ci, cell in enumerate(site_progress_row):
                    scope_name = col_scope.get(ci)
                    if not scope_name or not cell:
                        continue
                    plan_v, actual_v = _parse_wind_exec_cell(str(cell))
                    if plan_v is not None and actual_v is not None:
                        scopes[scope_name] = {'plan': plan_v, 'actual': actual_v}
            if scopes:
                break
    return scopes


def extract_wind_concerns_from_exec_summary(pdf):
    """
    The 'Concern need management attention' row of a Wind report. Returns None
    when the row itself could not be found, so the caller can fall back to the
    generic scan; [] means the row is there and empty this week.
    """
    items, found = _exec_summary_row_items(
        pdf, 'concern', _WIND_SCOPE_HEADERS, require_site_progress=True,
        page_limit=8)
    return items if found else None


def extract_wind_highlights_from_exec_summary(pdf):
    """
    The 'Key highlight activity this week' row of a Wind report. Same
    None-versus-empty contract as the concerns reader above.
    """
    items, found = _exec_summary_row_items(
        pdf, 'keyhighlight', _WIND_SCOPE_HEADERS, require_site_progress=True,
        page_limit=8)
    return items if found else None


# ── Wind extractor ─────────────────────────────────────────────────────────────

def extract_progress_wind(pdf, prj_id=None):
    """
    Wind report: find 'Overall Progress S-Curve (SCOPE)' pages.
    Each page has a table: Description | Plan% | Actual% | Ahead/Delay
    Rows: Engineering, Procurement, Construction, Overall
    Returns { plan, actual, scopes }.

    The Executive Summary's 'Site Progress' row (_extract_wind_from_exec_summary)
    is merged in afterward as the authoritative scope-level plan/actual —
    same convention as Solar: it fills scopes this scan missed entirely, and
    overrides plan/actual where both agree on the scope, while any
    discipline breakdown this scan found (Engineering/Procurement/...) is
    kept since the Exec Summary table doesn't have that level of detail.
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

    exec_scopes = _extract_wind_from_exec_summary(pdf)
    for scope_name, exec_entry in exec_scopes.items():
        tw_entry = scopes.get(scope_name)
        if tw_entry and tw_entry.get('disciplines'):
            exec_entry['disciplines'] = tw_entry['disciplines']
        scopes[scope_name] = exec_entry

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


def extract_progress_ttt(pdf):
    """
    TTT (Chang) weekly report, '3.2 S-Curve (Overall)' page.

    The W.F. Accumulation table beside the curve carries the whole picture:
    one row per EPCC discipline (weight factor, plan, actual, variance) and an
    'OverallProgress' row. Falls back to the 'Plan/Actual: X% / Y%' label above
    the curve if the table rows come out of pdfplumber too mangled to match.
    Returns { plan, actual, disciplines: {disc: {plan, actual, wf}} }.
    """
    empty = {'plan': None, 'actual': None, 'disciplines': {}}
    with pdfplumber.open(pdf) as doc:
        page_text = None
        for page in doc.pages:
            text = page.extract_text() or ''
            if TTT_SCURVE_PAGE.search(text):
                page_text = text
                break
    if page_text is None:
        return empty

    disciplines = {}
    for name, wf, plan_s, actual_s in TTT_AREA_ROW.findall(page_text):
        canon = TTT_DISC_CANON.get(name.lower())
        if not canon or canon in disciplines:
            continue
        plan_v, actual_v, wf_v = (_parse_pct(plan_s), _parse_pct(actual_s),
                                  _parse_pct(wf))
        if plan_v is None or actual_v is None:
            continue
        disciplines[canon] = {'plan': plan_v, 'actual': actual_v, 'wf': wf_v}

    plan = actual = None
    m = TTT_OVERALL_ROW.search(page_text)
    if m:
        plan, actual = _parse_pct(m.group(1)), _parse_pct(m.group(2))
    if plan is None or actual is None:
        m = TTT_PLAN_ACTUAL.search(page_text)
        if m:
            plan, actual = _parse_pct(m.group(1)), _parse_pct(m.group(2))
    if plan is None or actual is None:
        return {'plan': None, 'actual': None, 'disciplines': disciplines}
    return {'plan': plan, 'actual': actual, 'disciplines': disciplines}


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


# The chart page's only real (non-image) text is its title; the callout
# boxes ("Accumulative Planned/Achived Progress up to <week> =X.XX%") are
# baked into the page's image content, so OCR misreads "Achieved" as
# "Achived" fairly consistently — match both spellings.
_PAKLAY_CHART_PAGE_TITLE = re.compile(r'construction\s+progress\s+curve', re.I)
# As of W32/2026 the EPC dropped the "Accumulative ... up to <week>" wording
# and now labels the boxes plainly ("Planned Progress = 5.34%"), which the
# old regexes required and so silently returned nothing for two weeks running.
# Everything that isn't load-bearing is optional now: the "Accumulative"
# prefix, the "up to <week>" clause, and the spaces OCR drops between words
# ("PlannedProgress = 4.95 %" is a real W33 read) or adds before the '%'.
_PAKLAY_OCR_PLANNED  = re.compile(
    r'(?:accumulative\s*)?planned\s*progress\s*(?:up\s*to[^=]{0,25})?=\s*'
    r'(\d{1,3}(?:\.\d{1,2})?)\s*%', re.I)
_PAKLAY_OCR_ACHIEVED = re.compile(
    r'(?:accumulative\s*)?achi(?:e)?ved\s*progress\s*(?:up\s*to[^=]{0,25})?=\s*'
    r'(\d{1,3}(?:\.\d{1,2})?)\s*%', re.I)


def _ocr_paklay_scurve(search_dir):
    """
    Pak Lay's EPC weekly report has a '2.4 Construction progress curve' page
    that is entirely image content (a pasted chart) apart from its title —
    the Cumulative Plan/Achieved % callout boxes are baked into that image,
    not real PDF text, so OCR is the only way to read them. Searches every
    PDF under search_dir (the file isn't always named the same week to
    week) for a page whose real (non-OCR) text matches the title, since
    that's the only reliable, non-OCR way to locate the right page.
    When more than one 'Achieved' callout is present (a stale prior-week box
    alongside the current one), the LAST match wins — PowerPoint lists them
    in chronological order, so the last one is the most recent.
    Returns {'plan': float|None, 'actual': float|None} — never raises;
    any failure (Tesseract missing, page not found, OCR miss) yields Nones
    so a bad OCR read can never break the weekly run.
    """
    try:
        import pytesseract
        tess_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        if os.path.exists(tess_path):
            pytesseract.pytesseract.tesseract_cmd = tess_path

        for pdf_path in glob.glob(os.path.join(search_dir, '**', '*.pdf'), recursive=True):
            try:
                with pdfplumber.open(pdf_path) as doc:
                    target_page = None
                    for page in doc.pages:
                        t = page.extract_text() or ''
                        if _PAKLAY_CHART_PAGE_TITLE.search(t):
                            target_page = page
                            break
                    if target_page is None:
                        continue

                    # 300 DPI regularly loses the "Achieved" callout box entirely
                    # (Tesseract returns zero matches for it while still reading
                    # the "Planned" box fine) — 450 DPI reads both reliably.
                    #
                    # The "Achieved" callouts sit on a saturated ORANGE fill;
                    # Tesseract binarizes those RGB boxes inconsistently and often
                    # drops them entirely (while the blue "Planned" box reads fine),
                    # e.g. W30/2026 read Planned=4.71% but no Achieved at all.
                    # Converting to grayscale first (luminance → dark text on mid-
                    # grey) lets Tesseract's Otsu threshold recover the orange boxes;
                    # grayscale reads a strict superset of RGB here, so it's primary
                    # and RGB stays only as a fallback if grayscale finds nothing.
                    base_img = target_page.to_image(resolution=450).original
                    ocr_text = pytesseract.image_to_string(base_img.convert('L'))
                    plan_matches = _PAKLAY_OCR_PLANNED.findall(ocr_text)
                    achieved_matches = _PAKLAY_OCR_ACHIEVED.findall(ocr_text)
                    if not plan_matches and not achieved_matches:
                        ocr_text = pytesseract.image_to_string(base_img.convert('RGB'))
                        plan_matches = _PAKLAY_OCR_PLANNED.findall(ocr_text)
                        achieved_matches = _PAKLAY_OCR_ACHIEVED.findall(ocr_text)
                    plan_v   = _parse_pct(plan_matches[-1]) if plan_matches else None
                    actual_v = _parse_pct(achieved_matches[-1]) if achieved_matches else None
                    if plan_v is not None or actual_v is not None:
                        return {'plan': plan_v, 'actual': actual_v}
            except Exception as e:
                print(f"  [paklay OCR error] {os.path.basename(pdf_path)}: {e}")
    except Exception as e:
        print(f"  [paklay OCR error] {e}")

    return {'plan': None, 'actual': None}


# ── Pak Beng next-week activities ─────────────────────────────────────────────
#
# "6.2 Main Activities in Next Week" is a table, not a bulleted list:
#
#   No.  Construction Location   Work Content        Unit  Design   Remaining  Next-Week  Next-Week
#                                                          Quantity Quantity   Planned    Planned %
#   1.1  Right-bank Abutment     Earth-rock Excav.   m3    369000   53401.00   0          0.00%
#                               EL.364~379 Slope Sup Item  1        0.28       0.03       3.00%
#
# Most rows plan nothing for the coming week, so only those with a non-zero
# planned figure are reported, tagged with the location heading above them.

_PAKBENG_WORK_ROW = re.compile(
    r'^\s*(?:\d+(?:\.\d+)*\s+)?(\S.*?)\s+'
    r'(m\u00b3|m\u00b2|m3|m2|m|Item|Pcs|No|set|t)\s+'
    r'([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s*%\s*$', re.I)


def extract_pakbeng_activities(pdf, max_items=20):
    """
    Rows of "6.2 Main Activities in Next Week" that plan work for the coming
    week, with the quantity planned.
    Returns [] when the section is absent.
    """
    items = []
    try:
        with pdfplumber.open(pdf) as doc:
            for page in doc.pages:
                text = page.extract_text() or ''
                if not _PAKBENG_NEXT_WEEK_PAGE.search(text):
                    continue
                for raw in text.splitlines():
                    line = raw.strip()
                    m = _PAKBENG_WORK_ROW.match(line)
                    if m:
                        desc, unit, planned, pct = (m.group(1).strip(), m.group(2),
                                                    m.group(5), m.group(6))
                        try:
                            if float(pct.replace(',', '')) <= 0:
                                continue
                        except ValueError:
                            continue
                        # No location tag. This table wraps its location names
                        # across lines ("Downstream Approach" on one, "1.2
                        # ChannelAbove EL.349m" on the next), so the heading
                        # above a row cannot be tracked reliably - every row
                        # after the first location ended up carrying that one,
                        # putting the cofferdam's drainage channel under the
                        # right-bank abutment. A wrong location reads as fact;
                        # no location does not.
                        items.append(
                            f'{desc} \u2014 {planned} {unit} ({pct}%)')
    except Exception as e:
        print(f"  [pakbeng activities error] {e}")

    out, seen = [], set()
    for it in items:
        if it not in seen and len(it) > 10:
            seen.add(it)
            out.append(it)
    return out[:max_items]


# ── Pak Lay content ───────────────────────────────────────────────────────────
#
# Pak Lay's EPC uses its own report format. Section 2 is the next week's work
# plan, and its 2.3 Construction tables list the work item by item with
# quantities; 2.5 "Weekly Look Ahead Plan" in the same section is a percentage
# table, which is what the generic scan had been serving as activities
# ("Design Works 0.404% 17.14% 0.804% 17.946% ..."). Section 3.1 is the
# issues list.

_PAKLAY_NEXT_WEEK_SECTION = re.compile(r"next\s+week'?s?\s+work\s+plan", re.I)
_PAKLAY_CONSTRUCTION_PAGE = re.compile(r'^\s*2\.3\s+construction', re.I | re.M)
_PAKLAY_LOOKAHEAD_PAGE    = re.compile(r'^\s*2\.\d+\s+weekly\s+look\s+ahead',
                                       re.I | re.M)
_PAKLAY_ISSUES_PAGE       = re.compile(
    r'^\s*3(?:\.\d+)?\s+list\s+of\s+issues\s+to\s+be\s+coordinated',
    re.I | re.M)
# "1.2 Site hardening(#2~#5) m2 5213 600 180" - a numbered work row whose
# description is followed by a unit and its quantities.
_PAKLAY_WORK_ROW = re.compile(
    r'^\s*(\d+(?:\.\d+)+)\s+(\S.*?)\s+'
    r'(m3|m2|m\u00b3|m\u00b2|m|Item|Pcs|No|set|t)\s+'
    r'([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s*$', re.I)
# "1 EPC Hub Camp" - the group heading the rows below belong to.
_PAKLAY_WORK_GROUP = re.compile(r'^\s*(\d+)\s+(\D[^\d]{3,})\s*$')


def extract_paklay_activities(pdf, max_items=20):
    """
    Section 2.3's construction work rows, prefixed with their group
    ("[Left Bank Aggregate & Batching System] Conveyor belt concrete
    pouring - 293 m3"). Rows planning nothing this period are skipped, and
    2.5's percentage table is excluded outright.
    Returns [] when the section is absent.
    """
    items = []
    try:
        with pdfplumber.open(pdf) as doc:
            for page in doc.pages:
                text = page.extract_text() or ''
                if not _PAKLAY_CONSTRUCTION_PAGE.search(text):
                    continue
                if _PAKLAY_LOOKAHEAD_PAGE.search(text):
                    continue
                group = None
                for raw in text.splitlines():
                    line = raw.strip()
                    m = _PAKLAY_WORK_ROW.match(line)
                    if m:
                        desc, unit, planned = (m.group(2).strip(), m.group(3),
                                               m.group(5))
                        try:
                            if float(planned.replace(',', '')) <= 0:
                                continue
                        except ValueError:
                            pass
                        label = f'[{group}] ' if group else ''
                        items.append(f'{label}{desc} \u2014 {planned} {unit}')
                        continue
                    g = _PAKLAY_WORK_GROUP.match(line)
                    if g:
                        group = g.group(2).strip()
    except Exception as e:
        print(f"  [paklay activities error] {e}")

    out, seen = [], set()
    for it in items:
        if it not in seen and len(it) > 10:
            seen.add(it)
            out.append(it)
    return out[:max_items]


def extract_paklay_concerns(pdf, max_items=20):
    """
    Section 3.1 "List of Issues to be Coordinated". Numbered items, or the
    single word "None" - in which case there are genuinely no issues to
    coordinate this week and the honest answer is an empty list, not the
    twenty rows of table headers the generic scan was producing.
    Returns [] when the section is absent or says None.
    """
    items = []
    try:
        with pdfplumber.open(pdf) as doc:
            for page in doc.pages:
                text = page.extract_text() or ''
                if not _PAKLAY_ISSUES_PAGE.search(text):
                    continue
                current = None
                for raw in text.splitlines()[1:]:
                    line = raw.strip()
                    if not line or _PAGE_NUM.match(line):
                        continue
                    if line.upper() in ('NONE', 'N/A', 'NA'):
                        continue
                    if _IWTE_NUM_BULLET.match(line):
                        if current:
                            items.append(current)
                        current = _IWTE_NUM_BULLET.sub('', line).strip()
                    elif current is not None:
                        current += ' ' + line
                    else:
                        current = line
                if current:
                    items.append(current)
    except Exception as e:
        print(f"  [paklay concerns error] {e}")
    return [it for it in items if len(it) > 10][:max_items]


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
            return extract_progress_wind(pdf_path, prj_id=prj_id)
        elif extract_type == 'hydro_pakbeng':
            return extract_progress_hydro_pakbeng(pdf_path)
        elif extract_type == 'ttt':
            return extract_progress_ttt(pdf_path)
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

    # iWTE: Executive Summary's 'Area of Concern' field is the authoritative
    # source going forward. Not every iWTE report lays this field out on the
    # Executive Summary page (some instead have a separate, differently
    # structured 'Areas of Concern' section later in the document) — when
    # this comes back empty, keep whatever the generic scan above found
    # rather than discarding it.
    if AUTO_EXTRACT_PROJECTS.get(prj_id) == 'iwte':
        try:
            exec_concerns = extract_iwte_concerns_from_exec_summary(pdf_path)
            # [] means the report says there are none this week; None means
            # the field could not be located. Only the latter justifies the
            # generic whole-PDF scan, which on five of these reports had been
            # serving 20 rows of the milestone schedule table as "concerns".
            if exec_concerns is not None:
                concerns = exec_concerns
        except Exception as e:
            print(f"  [iwte concerns error] {e}")
        try:
            lookahead = extract_iwte_lookahead(pdf_path)
            if lookahead is not None:
                activities = lookahead
        except Exception as e:
            print(f"  [iwte lookahead error] {e}")

    # Wind: Executive Summary's 'Concern need management attention' row is
    # the authoritative source going forward. Unlike iWTE, the row is always
    # laid out on the same page in every Wind report, so when it's found and
    # genuinely empty ([], as opposed to None = "row not found at all") that
    # means no concerns this week — use it as-is rather than falling back to
    # whatever the generic whole-PDF scan picked up (which isn't scoped to
    # the Executive Summary and can surface stale/unrelated text).
    if AUTO_EXTRACT_PROJECTS.get(prj_id) == 'wind':
        try:
            exec_concerns = extract_wind_concerns_from_exec_summary(pdf_path)
            if exec_concerns is not None:
                concerns = exec_concerns
        except Exception as e:
            print(f"  [wind concerns error] {e}")

    # Solar and Wind both state the period's work per scope in their Executive
    # Summary's "Key highlight activity this week" row. The generic scan found
    # nothing at all for five of these projects and hit its item cap on table
    # text for three more, so prefer the row wherever it is present.
    if AUTO_EXTRACT_PROJECTS.get(prj_id) in ('solar', 'wind'):
        try:
            reader = (extract_wind_highlights_from_exec_summary
                      if AUTO_EXTRACT_PROJECTS.get(prj_id) == 'wind'
                      else extract_solar_highlights_from_exec_summary)
            highlights = reader(pdf_path)
            if highlights:
                activities = highlights
        except Exception as e:
            print(f"  [highlights error] {e}")

    # GMTP: section 10 "Key Issues and Concerns" and sections 6.3-6.5 "Main
    # Activities in Next Week" are the authoritative sources; the generic scan
    # above reads the SHE/milestone tables instead. Only replace what the
    # dedicated readers actually found, so a report that drops a section keeps
    # whatever the generic scan managed to pick up.
    if AUTO_EXTRACT_PROJECTS.get(prj_id) == 'gmtp':
        try:
            gmtp_concerns, gmtp_found = extract_gmtp_concerns_checked(pdf_path)
            if gmtp_found:
                concerns = gmtp_concerns
            gmtp_activities = extract_gmtp_next_week_activities(pdf_path)
            if gmtp_activities:
                activities = gmtp_activities
        except Exception as e:
            print(f"  [gmtp content error] {e}")

    # Pak Beng shares GMTP's report template, so the same section readers
    # apply once their patterns allow its spelling.
    if AUTO_EXTRACT_PROJECTS.get(prj_id) == 'hydro_pakbeng':
        try:
            pb_concerns, pb_found = extract_gmtp_concerns_checked(
                pdf_path, page_re=_PAKBENG_CONCERN_PAGE,
                item_re=_PAKBENG_CONCERN_ITEM)
            if pb_found:
                concerns = pb_concerns
            pb_activities = extract_pakbeng_activities(pdf_path)
            if pb_activities:
                activities = pb_activities
        except Exception as e:
            print(f"  [pakbeng content error] {e}")

    # Pak Lay uses its own format; both of its sections are authoritative,
    # including when the issues list says "None".
    if AUTO_EXTRACT_PROJECTS.get(prj_id) == 'hydro_paklay':
        try:
            concerns = extract_paklay_concerns(pdf_path)
            pl_activities = extract_paklay_activities(pdf_path)
            if pl_activities:
                activities = pl_activities
        except Exception as e:
            print(f"  [paklay content error] {e}")

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
