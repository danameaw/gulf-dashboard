# ── run.py ── Gulf Dashboard Weekly Automation ────────────────────────────
# Schedule: every Wednesday via Windows Task Scheduler
# Usage: python run.py  (or python run.py --week 25 --year 2026)
import sys, os, re, json, glob, shutil, subprocess, argparse
from datetime import datetime
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

sys.path.insert(0, os.path.dirname(__file__))
from config import (BASE_REPORT_PATH, DASHBOARD_HTML, EXCEL_DIR,
                    GIT_REPO, DASHBOARD_URL,
                    FOLDER_PROJECTS, PROJECT_KEYWORDS, PROJECT_NAMES)
from extract import extract_from_pdf


# ── 1. Find latest week folder ─────────────────────────────────────────────
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
    return latest, week, yr


# ── 2. Find PDF for a project ──────────────────────────────────────────────
def find_pdf(folder_path, prj_id):
    keywords = PROJECT_KEYWORDS.get(prj_id, [])
    pdfs = glob.glob(os.path.join(folder_path, '*.pdf'))
    pdfs += glob.glob(os.path.join(folder_path, '**', '*.pdf'), recursive=True)
    for pdf in pdfs:
        fname = os.path.basename(pdf).upper()
        if any(kw.upper() in fname for kw in keywords):
            return pdf
    return None


# ── 3. Build JS seed snippet ───────────────────────────────────────────────
def js_escape(s):
    return s.replace('\\', '\\\\').replace("'", "\\'").replace('\n', ' ')

def build_seed(prj_id, week, year, data):
    concerns   = ',\n    '.join(f"'{js_escape(c)}'" for c in data['concerns'])
    activities = ',\n    '.join(f"'{js_escape(a)}'" for a in data['activities'])
    return (
        f"// ── {PROJECT_NAMES.get(prj_id, prj_id)} ({prj_id}) Week {week} — auto-extracted\n"
        f"seedIfEmpty('{prj_id}_W{week}_{year}', {{\n"
        f"  plan: null, actual: null,\n"
        f"  concerns: [{concerns}],\n"
        f"  activities: [{activities}],\n"
        f"}});\n"
    )


# ── 4. Update index.html ───────────────────────────────────────────────────
SEED_MARKER   = '// ── SEED DATA (pre-loaded from PDF reports) ──'
MISSING_MARKER_START = '// ── AUTO: MISSING REPORTS ──'
MISSING_MARKER_END   = '// ── END MISSING REPORTS ──'

def update_html(week, year, seeds_js, missing_ids):
    with open(DASHBOARD_HTML, encoding='utf-8') as f:
        html = f.read()

    # Inject seed data after SEED_MARKER
    insert_after = SEED_MARKER + '\nfunction seedIfEmpty(key, data) {\n  if (!progressData[key]) { progressData[key] = data; }\n}'
    if insert_after in html:
        inject_point = html.index(insert_after) + len(insert_after)
        new_seeds = '\n\n// ── Week ' + str(week) + '/' + str(year) + ' — AUTO-GENERATED ──\n'
        new_seeds += '\n'.join(seeds_js)
        html = html[:inject_point] + new_seeds + html[inject_point:]
    else:
        print("  [warn] SEED_MARKER not found, appending seeds before </script>")
        html = html.replace('</script>', '\n'.join(seeds_js) + '\n</script>', 1)

    # Update missing reports variable
    missing_js = (
        f"{MISSING_MARKER_START}\n"
        f"const MISSING_REPORTS = {json.dumps(missing_ids)};\n"
        f"{MISSING_MARKER_END}"
    )
    if MISSING_MARKER_START in html:
        start = html.index(MISSING_MARKER_START)
        end   = html.index(MISSING_MARKER_END) + len(MISSING_MARKER_END)
        html  = html[:start] + missing_js + html[end:]
    else:
        # Insert before first <script> tag
        html = html.replace('<script>', missing_js + '\n<script>', 1)

    with open(DASHBOARD_HTML, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"  ✓ index.html updated")


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


# ── 7. Windows Notification ────────────────────────────────────────────────
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
    else:
        week_folder, week, year = find_week_folder()

    print(f"  Week: {week}/{year}")
    print(f"  Folder: {week_folder}")
    print()

    results    = {}   # prj_id → {found, data}
    seeds_js   = []
    missing    = []

    # Process each project group
    for folder_name, prj_ids in FOLDER_PROJECTS.items():
        group_path = os.path.join(week_folder, folder_name)
        folder_exists = os.path.isdir(group_path)

        for prj_id in prj_ids:
            name = PROJECT_NAMES.get(prj_id, prj_id)
            print(f"  [{prj_id}] {name}", end=' ... ')

            if not folder_exists:
                print(f"FOLDER NOT FOUND ({folder_name})")
                results[prj_id] = {'found': False, 'data': {'concerns': [], 'activities': []}}
                missing.append(prj_id)
                continue

            pdf = find_pdf(group_path, prj_id)
            if not pdf:
                print("PDF NOT FOUND")
                results[prj_id] = {'found': False, 'data': {'concerns': [], 'activities': []}}
                missing.append(prj_id)
                continue

            print(f"OK → {os.path.basename(pdf)}")
            data = extract_from_pdf(pdf)
            print(f"       concerns={len(data['concerns'])}, activities={len(data['activities'])}")
            results[prj_id] = {'found': True, 'data': data}
            seeds_js.append(build_seed(prj_id, week, year, data))

    print()
    print(f"  Found:   {sum(1 for r in results.values() if r['found'])} projects")
    print(f"  Missing: {len(missing)} projects: {missing}")

    if args.dry_run:
        print("\n  [dry-run] Skipping file updates.")
        return

    # Update files
    print("\n  Updating dashboard...")
    update_html(week, year, seeds_js, missing)
    update_excel(week, year, results)
    git_push(week, year)
    notify(week, year,
           sum(1 for r in results.values() if r['found']),
           len(missing))

    print("\n  Done! Dashboard URL:", DASHBOARD_URL)
    print("=" * 60)


if __name__ == '__main__':
    main()
