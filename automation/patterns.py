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

IWTE_PAGE_INCLUDE = re.compile(r'weekly|week\s+\d', re.I)
IWTE_PAGE_EXCLUDE = re.compile(r'monthly|cumulative\s+monthly', re.I)

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

# ── Summary: which projects CAN be auto-extracted ──────────────────────────
AUTO_EXTRACT_PROJECTS = {
    # project_id: extraction_type
    'PRJ-001': 'gmtp',    # GMTP/LNG — text table
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
    # All Solar, Wind, Hydro → S-CURVE IMAGE → cannot auto-extract
}
