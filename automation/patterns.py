# ── patterns.py ── PDF Reading Patterns per Project Type ──────────────────
#
# Documents HOW to extract Plan% / Actual% from each project type.
# Updated as new patterns are learned from PDF reports.
#
# LEGEND:
#   S-CURVE IMAGE  = values are in a chart image → cannot auto-extract
#   TEXT TABLE     = values are in text/table → can auto-extract with regex
#   MANUAL         = must be entered by hand
# ──────────────────────────────────────────────────────────────────────────

import re

# ── iWTE (PRJ-013 to PRJ-024) ─────────────────────────────────────────────
# Source: Weekly Progress Table (text, NOT the monthly S-curve page)
# Page pattern: contains "Weekly" or "Week" AND plan/actual rows
# Table format example (GCE Week 24):
#   Discipline      Plan%    Actual%
#   Engineering     100.00   100.00
#   Procurement      95.00    90.00
#   Construction     17.91     9.31
#   Commissioning     0.00     0.00
#   Overall          40.47    36.41
#
# IMPORTANT: Use WEEKLY table only, skip Monthly table (different page)
# Keyword to identify weekly page: "weekly" + "plan" + "actual" + "%"
# Keyword to SKIP: "monthly" or "cumulative monthly"

IWTE_DISC_ORDER = ['Engineering', 'Procurement', 'Construction', 'Commissioning']
IWTE_DISC_WEIGHTS = {'Engineering': 0.10, 'Procurement': 0.40,
                     'Construction': 0.40, 'Commissioning': 0.10}

IWTE_PAGE_INCLUDE = re.compile(r'weekly construction progress|weekly.*progress.*epc|week\s+\d', re.I)
IWTE_PAGE_EXCLUDE = re.compile(r'monthly progress status|cumulative\s+monthly', re.I)
# Table format: columns 0=discipline, 1=weight, 2=prev_plan, 3=prev_actual, 4=prev_diff,
#               5=curr_plan, 6=curr_actual, 7=curr_diff, 8=next_plan
IWTE_CURR_PLAN_COL   = 5
IWTE_CURR_ACTUAL_COL = 6

# Matches a % value like 40.47 or 100.00 or 0.00
PCT = r'(\d{1,3}(?:\.\d{1,2})?)'

# Row patterns in iWTE weekly table
IWTE_ROWS = {
    'Engineering':    re.compile(r'engineer\w*\s+' + PCT + r'\s+' + PCT, re.I),
    'Procurement':    re.compile(r'procure\w*\s+' + PCT + r'\s+' + PCT, re.I),
    'Construction':   re.compile(r'construct\w*\s+' + PCT + r'\s+' + PCT, re.I),
    'Commissioning':  re.compile(r'commission\w*\s+' + PCT + r'\s+' + PCT, re.I),
    'Overall':        re.compile(r'overall\s+' + PCT + r'\s+' + PCT, re.I),
}

# ── GMTP / LNG (PRJ-001) ──────────────────────────────────────────────────
# Source: Section 3 "Overall Progress" table — text-based discipline table
# Page pattern: contains discipline names + plan/actual headers
# Format example:
#   Discipline       Plan%    Actual%   Variance
#   Engineering      64.01    72.60     +8.59
#   Procurement      24.27    26.73     +2.46
#   Construction     15.04    18.41     +3.37
#   Commissioning     0.00     0.00      0.00
#   Overall          20.89    23.94     +3.05

GMTP_PAGE_INCLUDE = re.compile(r'overall.*progress|progress.*summary|discipline', re.I)
GMTP_ROWS = {
    'Engineering':   re.compile(r'eng\w*\s+' + PCT + r'\s+' + PCT, re.I),
    'Procurement':   re.compile(r'proc\w*\s+' + PCT + r'\s+' + PCT, re.I),
    'Construction':  re.compile(r'construct\w*\s+' + PCT + r'\s+' + PCT, re.I),
    'Commissioning': re.compile(r'commission\w*\s+' + PCT + r'\s+' + PCT, re.I),
    'Overall':       re.compile(r'overall\s+' + PCT + r'\s+' + PCT, re.I),
}

# ── Solar GSO2026 (PRJ-002, PRJ-003, PRJ-004) ─────────────────────────────
# Source: Overall S-Curve → IMAGE (cannot auto-extract plan/actual %)
# MANUAL: read from S-curve image, enter manually
# Discipline breakdown: GUE has 4 disciplines (Eng/Pro/Con/Com)
#   BUT Eng and Pro are usually 100% → only Con and Com matter
# Siemens: Civil + EE + T&C (EE often 100%)
# T/L: PEA-FEC or PEA-TEC → overall only, no sub-disciplines
# → Overall: S-CURVE IMAGE

# ── Solar SOSB2026 LNE/SSE (PRJ-005, PRJ-006) ────────────────────────────
# Source: GUE S-Curve (overall) → IMAGE
# GUE: Eng 100%, Pro 100% → only read Con and Com S-curves
# Siemens: Civil S-curve + T&C S-curve → IMAGE
# 115kV T/L: completed 100%
# → Overall: S-CURVE IMAGE

# ── Solar SDCE (PRJ-007) ──────────────────────────────────────────────────
# Source: GPD Overall S-Curve → IMAGE
# GPD: Eng/Pro/Con&Com combined
# Siemens Sub: Civil + EE + T&C → IMAGE
# T/L: PEA-TEC → overall only → IMAGE
# → Overall: S-CURVE IMAGE

# ── Wind AL1 (PRJ-009) ────────────────────────────────────────────────────
# CBOP (PCZ): S-Curve "Overall Progress S-Curve (PCZ)" → IMAGE
# TSA (Goldwind GW): Recovery Plan S-Curve → IMAGE
# Substation: 0% (not started yet as of W24)
# 115kV T/L (SCT): S-Curve → IMAGE
# → Overall: S-CURVE IMAGE (4 scopes)

# ── Wind AL2 (PRJ-010) ────────────────────────────────────────────────────
# CBOP (PCZ): S-Curve → IMAGE
# TSA (Goldwind GW): same report as AL1 TSA → IMAGE
# Substation (Siemens): S-Curve → IMAGE
# 115kV T/L (ENCOM): S-Curve → IMAGE
# → Overall: S-CURVE IMAGE (4 scopes)

# ── Wind ECE (PRJ-011) ────────────────────────────────────────────────────
# CBOP: Rev.B Acc. Schedule Revised S-Curve → IMAGE
# TSA (Goldwind GW): Recovery Plan S-Curve → IMAGE
# Substation: S-Curve → IMAGE
# 115kV T/L (RSS): bi-weekly report, S-Curve → IMAGE
# → Overall: S-CURVE IMAGE (4 scopes)

# ── Hydro Pak Lay (PRJ-025) ───────────────────────────────────────────────
# Source: PLHPP Overall S-Curve → IMAGE
# Overall only, no discipline breakdown
# → Overall: S-CURVE IMAGE

# ── Hydro Pak Beng (PRJ-026) ──────────────────────────────────────────────
# Source: POWERCHINA Overall S-Curve → IMAGE
# Overall only, no discipline breakdown
# → Overall: S-CURVE IMAGE

# ── TTT (PRJ-027) ─────────────────────────────────────────────────────────
# No data yet → MANUAL when reports available

# ── Solar (GSO2026: PRJ-002,003,004 / SOSB2026: PRJ-005,006,007) ─────────
# Source: "This week activities" pages per contractor section
# Page identification:
#   GUE scope:       page title contains "Construction Status : GUE"
#   Siemens scope:   page title contains "Siemens" (substation)
#   115kV T/L scope: page title contains "115kV" or "Transmission Line"
#   Comm/Fiber:      page title contains "Fiber" or "ITL" or "Communication"
#
# Pattern in text (on "This week activities" page, NOT "Previous week"):
#   "Overall progress stands at A.AA% against the planned progress of B.BB%"
#   "Engineering progress stands at A.AA% against the planned progress of B.BB%"
#   "Procurement progress stands at A.AA% against the planned progress of B.BB%"
#   "Construction progress stands at A.AA% against the planned progress of B.BB%"
#
# Scope mapping (same across all Solar projects):
#   GUE       → disciplines: Engineering, Procurement, Construction
#   Siemens   → single overall (Construction only usually)
#   115kV T/L → single overall
#   Comm/Fiber→ single overall (optional, may not always have report)

SOLAR_THIS_WEEK = re.compile(r'this week activities', re.I)
SOLAR_PREV_WEEK = re.compile(r'previous week activities', re.I)
# NWT3-style: "Overall progress stands at A.AA% against the planned progress of B.BB%"
# group(1)=discipline, group(2)=actual, group(3)=plan
SOLAR_PROGRESS  = re.compile(
    r'(overall|engineering|procurement|construction)\s+(?:project\s+)?progress\s+stands\s+at\s+'
    r'(\d{1,3}(?:\.\d{1,2})?)\s*%\s+against\s+the\s+planned\s+progress\s+of\s+'
    r'(\d{1,3}(?:\.\d{1,2})?)\s*%',
    re.I)

# PTN/STN-style: "Overall project progress 58.26% compared to plan at 82.19%"
# group(1)=discipline, group(2)=actual, group(3)=plan
SOLAR_PROGRESS_ALT = re.compile(
    r'(overall|engineering|procurement|construction|civil|electrical)\s+'
    r'(?:project\s+|pre-civil\s+|construction\s+|work\s+)?progress\s+'
    r'(\d{1,3}(?:\.\d{1,2})?)\s*%\s+compared\s+to\s+plan\s+at\s+'
    r'(\d{1,3}(?:\.\d{1,2})?)\s*%',
    re.I)

# SSE/LNE/SDCE-style (executive summary Site Progress cell).
# Handles all 3 sub-variants:
#   "/Plan at"  (SSE):  "Construction progress 97.49% /Plan at 99.91%"
#   "vs plan"   (LNE):  "Overall progress is 93.97% vs plan 96.66%"
#   "against plan" (SDCE): "Overall progress 81.39% against plan 83.02%"
# group(1)=discipline, group(2)=actual, group(3)=plan
SOLAR_PROGRESS_SUMMARY = re.compile(
    r'(overall|construction|commissioning|civil|electrical)\s*'
    r'progress\s*(?:is\s+)?'           # "progress" then optional "is"
    r'(\d{1,3}(?:\.\d{1,2})?)\s*%\s*'
    r'(?:/\s*Plan\s*at|vs\s+plan|against\s+plan)\s*'
    r'(\d{1,3}(?:\.\d{1,2})?)\s*%',
    re.I)

# Scope detection from page text (title line)
SOLAR_SCOPE_GUE   = re.compile(r'construction status\s*[:\-]\s*gue|status.*\bgue\b', re.I)
SOLAR_SCOPE_SIE   = re.compile(r'siemens|substation.*siemens|siemens.*substation', re.I)
SOLAR_SCOPE_TL    = re.compile(r'115\s*kv|transmission line|t/l', re.I)
SOLAR_SCOPE_COMM  = re.compile(r'fiber|communication|itl|กบศ', re.I)

# ── Wind (PRJ-009 AL1, PRJ-010 AL2, PRJ-011 ECE) ──────────────────────────
# Source: "Overall Progress S-Curve (SCOPE)" pages — text table
# Page title format: "Overall Progress S-Curve (PCZ)" or "(GW)" or "(SCT)" etc.
# Table format (same for all scopes):
#   Description  Plan     Actual   Ahead/Delay
#   Engineering  100%     84.12%   15.88%
#   Procurement  57.20%   47.00%   10.20%
#   Construction 20.07%   13.95%   6.13%
#   Overall      23.05%   16.62%   6.43%
#
# Scope mapping:
#   AL1: CBOP=PCZ, TSA=GW, Substation=SIEMENS, T/L=SCT
#   AL2: CBOP=PCZ, TSA=GW, Substation=SIEMENS, T/L=ENCOM
#   ECE: CBOP=PCZ, TSA=GW, Substation=SIEMENS, T/L=RSS
#
# IMPORTANT: Use "Overall Progress S-Curve" page, NOT "Construction Progress S-Curve"
# The "Construction Progress" page has only the Construction row → skip it

WIND_SCURVE_PAGE = re.compile(r'overall progress s-?curve\s*\(([^)]+)\)', re.I)
WIND_DISC_ROW    = re.compile(
    r'^(engineering|procurement|construction|commissioning|overall)\s+'
    r'(\d{1,3}(?:\.\d{1,2})?)\s*%\s+'
    r'(\d{1,3}(?:\.\d{1,2})?)\s*%',
    re.I | re.M)

# Scope name normalization for Wind
WIND_SCOPE_MAP = {
    'pcz':    'CBOP',
    'gw':     'TSA',
    'goldwind': 'TSA',
    'siemens': 'Substation',
    'sct':    '115kV T/L',
    'encom':  '115kV T/L',
    'rss':    '115kV T/L',
}

# ── Hydro Pak Beng (PRJ-026) ───────────────────────────────────────────────
# Source: Monthly progress report, page with S-Curve summary
# Pattern 1 (text line): "Plan/Actual：5.65%/5.24%（0.41%）"
# Pattern 2 (table row): "Cum. Progress -Planned ... 5.65%"
#                         "Cum. Progress -Actual  ... 5.24%"
# → Use the LAST value in each row (current month's cumulative)

PAKBENG_PLAN_ACTUAL = re.compile(
    r'plan/actual\s*[：:]\s*(\d{1,3}(?:\.\d{1,2})?)%\s*/\s*(\d{1,3}(?:\.\d{1,2})?)%',
    re.I)
PAKBENG_CUM_PLANNED = re.compile(
    r'cum\.?\s*progress\s*[-–]\s*planned[^\n]*?(\d{1,3}\.\d{2})%\s*$', re.I | re.M)
PAKBENG_CUM_ACTUAL  = re.compile(
    r'cum\.?\s*progress\s*[-–]\s*actual[^\n]*?(\d{1,3}\.\d{2})%\s*$', re.I | re.M)

# ── Hydro Pak Lay (PRJ-025) ────────────────────────────────────────────────
# Source: EPC Weekly Report
# Note: Report focuses on issue tracking / design submissions, NOT a clear
# overall progress % page in text form. S-curve may be image-based.
# If a "Physical Progress" or "Overall Progress" text line appears, capture it.
# Pattern: "Physical Progress: Plan X.XX% / Actual Y.YY%"
#          or "overall progress ... X.XX% ... Y.YY%"

PAKLAY_PROGRESS = re.compile(
    r'physical\s+progress.*?(\d{1,3}(?:\.\d{1,2})?)%.*?(\d{1,3}(?:\.\d{1,2})?)%',
    re.I)

# ── Summary: which projects CAN be auto-extracted ──────────────────────────
AUTO_EXTRACT_PROJECTS = {
    # project_id: extraction_type
    'PRJ-001': 'gmtp',    # GMTP/LNG — text discipline table
    'PRJ-002': 'solar',   # GOE2-NWT3
    'PRJ-003': 'solar',   # GSPG-PTN
    'PRJ-004': 'solar',   # GSPG-STN
    'PRJ-005': 'solar',   # LNE
    'PRJ-006': 'solar',   # SSE
    'PRJ-007': 'solar',   # SDCE
    'PRJ-009': 'wind',    # AL1
    'PRJ-010': 'wind',    # AL2
    'PRJ-011': 'wind',    # ECE
    'PRJ-013': 'iwte',    # GCE
    'PRJ-014': 'iwte',    # GSE
    'PRJ-015': 'iwte',    # KKE
    'PRJ-016': 'iwte',    # MKP
    'PRJ-017': 'iwte',    # MPE
    'PRJ-018': 'iwte',    # PFP
    'PRJ-019': 'iwte',    # PKP
    'PRJ-020': 'iwte',    # PSD
    'PRJ-021': 'iwte',    # PWW1
    'PRJ-022': 'iwte',    # PWW2
    'PRJ-023': 'iwte',    # TPP
    'PRJ-024': 'iwte',    # TSE
    'PRJ-025': 'hydro_paklay',   # Pak Lay
    'PRJ-026': 'hydro_pakbeng',  # Pak Beng
}
