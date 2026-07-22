# ── config.py ── Gulf Dashboard Automation Configuration ──────────────────

BASE_REPORT_PATH = r"C:\Users\danaya.th\Gulf\Engineering - Engineering Documents\00 Project Reports\2026"
DASHBOARD_HTML   = r"C:\Users\danaya.th\OneDrive - Gulf\Documents\GitHub\gulf-dashboard\index.html"
EXCEL_DIR        = r"C:\Users\danaya.th\OneDrive - Gulf\Documents\GitHub\gulf-dashboard"
GIT_REPO         = r"C:\Users\danaya.th\OneDrive - Gulf\Documents\GitHub\gulf-dashboard"
DASHBOARD_URL    = "https://danameaw.github.io/gulf-dashboard/"

# Folder inside week directory → project IDs in that folder
FOLDER_PROJECTS = {
    "GMTP":     ["PRJ-001"],
    "GSO2026":  ["PRJ-002", "PRJ-003", "PRJ-004"],
    "iWTE":     ["PRJ-013","PRJ-014","PRJ-015","PRJ-016","PRJ-017",
                 "PRJ-018","PRJ-019","PRJ-020","PRJ-021","PRJ-022",
                 "PRJ-023","PRJ-024"],
    "Pak Beng": ["PRJ-026"],
    "Pak Lay":  ["PRJ-025"],
    "SOSB2026": ["PRJ-005","PRJ-006","PRJ-007"],
    "TTT":      ["PRJ-027"],
    "Wind2027": ["PRJ-008","PRJ-009","PRJ-010","PRJ-011","PRJ-012"],
}

# Keywords to match PDF filename → project ID
PROJECT_KEYWORDS = {
    "PRJ-001": ["GMTP", "LNG"],
    "PRJ-002": ["NWT3", "GOE2-NWT3"],
    "PRJ-003": ["PTN",  "GSPG-PTN"],
    "PRJ-004": ["STN",  "GSPG-STN"],
    "PRJ-005": ["LNE",  "Lopburi"],
    "PRJ-006": ["SSE",  "Sara"],
    "PRJ-007": ["SDCE", "PBR8", "Chai Nat"],
    "PRJ-008": ["AC8", "GULF AC8"],
    "PRJ-009": ["ALPHA 1", "AL1", "Alpha1", "Alpha 1"],
    "PRJ-010": ["ALPHA 2", "AL2", "Alpha2", "Alpha 2"],
    "PRJ-011": ["ECE"],
    "PRJ-012": ["WAYU", "Wayu", "GULF_Wayu", "GULF Wayu"],
    "PRJ-013": ["GCE"],
    "PRJ-014": ["GSE"],
    "PRJ-015": ["KKE",  "Khon Kaen"],
    "PRJ-016": ["MKP",  "Map Kha"],
    "PRJ-017": ["MPE",  "Map Pha"],
    "PRJ-018": ["PFP",  "Pranburi"],
    "PRJ-019": ["PKP",  "Pak Kret"],
    "PRJ-020": ["PSD",  "Panat"],
    "PRJ-021": ["PWW1"],
    "PRJ-022": ["PWW2"],
    "PRJ-023": ["TPP",  "Trang"],
    "PRJ-024": ["TSE",  "Thung Song"],
    "PRJ-025": ["Pak Lay",  "PakLay",  "PLHPP"],
    "PRJ-026": ["Pak Beng", "PakBeng", "POWERCHINA"],
    "PRJ-027": ["TTT", "CHANG"],
}

# Projects with no extraction pattern yet — report has arrived (so not
# "missing") but data is filled in by hand instead of parsed from the PDF.
MANUAL_OVERRIDES = {
    "PRJ-027": {"concerns": [], "activities": ["To be later"], "plan": None, "actual": None},
}

PROJECT_NAMES = {
    "PRJ-001": "LNG Terminal (GMTP)",
    "PRJ-002": "GOE2-NWT3",
    "PRJ-003": "GSPG-PTN",
    "PRJ-004": "GSPG-STN",
    "PRJ-005": "LNE",
    "PRJ-006": "SSE",
    "PRJ-007": "SDCE",
    "PRJ-008": "GULF AC8",
    "PRJ-009": "AL1 (Alpha 1)",
    "PRJ-010": "AL2 (Alpha 2)",
    "PRJ-011": "ECE",
    "PRJ-012": "GULF_Wayu",
    "PRJ-013": "GCE",
    "PRJ-014": "GSE",
    "PRJ-015": "KKE",
    "PRJ-016": "MKP",
    "PRJ-017": "MPE",
    "PRJ-018": "PFP",
    "PRJ-019": "PKP",
    "PRJ-020": "PSD",
    "PRJ-021": "PWW1",
    "PRJ-022": "PWW2",
    "PRJ-023": "TPP",
    "PRJ-024": "TSE",
    "PRJ-025": "Pak Lay",
    "PRJ-026": "Pak Beng",
    "PRJ-027": "TTT",
}
