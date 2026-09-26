"""
Radiometer ABL Blood Gas — Operator Error Report Generator  v1
Run via start_windows.bat or start_mac.command

Security hardening built in from the start (PSPF / ISM / IS18 / APP):
  - Active network guard (netguard.py): outbound connections blocked + logged
  - tempfile.mkstemp() for all uploads — randomised, deleted in finally (no data remanence)
  - /open_folder confined to OUTPUT_ROOT via realpath + commonpath (path traversal)
  - Origin/Referer CSRF check on all POST routes
  - Audit log: timestamp, OS user, action, period, scope
  - Explicit 127.0.0.1 loopback bind (never exposed on network)
  - output folder built by _output_dir_for from digits/letters only (no path traversal via crafted month_name)
"""
import os, re, sys, threading, subprocess, platform, json, io, tempfile, getpass, logging

# Make this app's own folder importable regardless of how Python was started
# (the embeddable runtime does not add the script directory to sys.path).
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import netguard
netguard.install()   # loopback-only networking from this point on (see netguard.py)

from flask import Flask, render_template, request, jsonify, abort
import openpyxl
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
APP_COMPONENT = 'abl'

# ═══════════════════════════════════════════════════════════════════════════════
#  PC ENROLMENT — the bundle runs only on PCs enrolled via the portal (:5750).
#  pq_enrolment.py lives at the bundle root; without it, or without a valid
#  enrolment for this machine, every request is answered with a lock page.
# ═══════════════════════════════════════════════════════════════════════════════
def _load_enrolment():
    import importlib.util
    _here = os.path.dirname(os.path.abspath(__file__))
    for cand in (os.path.join(_here, '..'), os.path.join(_here, '..', 'PQ_Portable')):
        path = os.path.join(cand, 'pq_enrolment.py')
        if os.path.exists(path):
            try:
                spec = importlib.util.spec_from_file_location('pq_enrolment', path)
                mod  = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(mod)
                return mod
            except Exception as exc:
                print(f'[ENROL] failed to load {path}: {exc}')
                return None
    return None

_ENROL = _load_enrolment()
if _ENROL is not None:
    try:
        _ENROL.record_launch(APP_COMPONENT)
    except Exception:
        pass

_LOCK_PAGE = """<!DOCTYPE html><html><head><meta charset="utf-8"><title>PC not enrolled</title>
<style>body{font-family:-apple-system,Segoe UI,Arial,sans-serif;background:#F2F2F7;color:#17272E;display:flex;
align-items:center;justify-content:center;height:100vh;margin:0}.card{background:#fff;border-radius:16px;
padding:36px 40px;max-width:560px;box-shadow:0 10px 34px rgba(1,38,50,.25)}h1{color:#01485C;font-size:22px;margin:0 0 12px}
p{line-height:1.5;margin:8px 0}code{background:#E5F2F0;padding:2px 6px;border-radius:6px}a{color:#01485C}</style></head>
<body><div class="card"><h1>This PC is not enrolled</h1><p>{msg}</p>
<p>Machine ID <code>{mid}</code> &middot; {host}</p>
<p>Open the selection page at <a href="http://localhost:5750/">localhost:5750</a> and enrol this PC with the
PoC team&rsquo;s enrolment password. Nothing can be generated until then.</p></div></body></html>"""

@app.before_request
def _enforce_enrolment():
    if _ENROL is None:
        return ('<h2 style="font-family:sans-serif;color:#B4540A">Enrolment module missing &mdash; '
                'this copy of the bundle is incomplete and cannot run.</h2>'), 403
    st = _ENROL.status()
    if not st['allowed']:
        body = (_LOCK_PAGE.replace('{msg}', st['message']).replace('{mid}', st['machine_id'])
                .replace('{host}', st['hostname']))
        return body, 403, {'Content-Type': 'text/html; charset=utf-8'}


progress = {"total": 0, "done": 0, "current": "", "errors": [], "complete": False, "output": "", "reports": 0}

# ── Colour palette (matches iSTAT app for PQ visual consistency) ──────────────
C_DARK_BLUE  = RGBColor(0x1F, 0x49, 0x7D)
C_GREEN      = RGBColor(0x9B, 0xBB, 0x59)
C_BLUE       = RGBColor(0x30, 0x54, 0x96)
C_WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
C_RED_TEXT   = RGBColor(0xFF, 0x00, 0x00)
C_GREY_TEXT  = RGBColor(0x60, 0x60, 0x60)
C_GREEN_PASS = RGBColor(0x00, 0x70, 0x00)

FILL_GREY_HDR = 'AEAAAA'
FILL_RED_HDR  = 'FF0000'
FILL_BLUE_HDR = 'D9E1F2'
FILL_ALT_ROW  = 'D9D9D9'
FILL_WHITE    = 'FFFFFF'

# ── Fixed Radiometer ABL pre-analytical error codes ───────────────────────────
ABL_ERRORS = {
    '328': 'No leading air segment in inlet\'s liquid sensor within time frame',
    '331': 'No Sample Detected',
    '521': 'Inhomogenous sample',
    '593': 'Insufficient sample',
    '664': 'Sample Problem',
    '722': 'Sample Error',
    '791': 'Insufficient or Inhomogenous sample',
}
ERROR_CODES = ['328', '331', '521', '593', '664', '722', '791']

# Fixed colour per error code (charts + legend) — never reassigned by rank
ERROR_COLORS = {
    '328': '#4A3AA7',   # violet
    '331': '#2A78D6',   # blue
    '521': '#1BAF7A',   # teal
    '593': '#EDA100',   # amber
    '664': '#E34948',   # red
    '722': '#E87BA4',   # pink
    '791': '#EB6834',   # orange
}

# ── Flagging thresholds ───────────────────────────────────────────────────────
MIN_TESTS = 10    # flagging requires MORE than this many tests in the month (11+)
FLAG_PCT  = 10.0  # Pct. error above this flags the operator

# ── PICU operator list ────────────────────────────────────────────────────────
# Staff on this list appear under ICU in the export but are PICU employees.
# ICU reports separate them into their own section.  Update the list by
# replacing picu_operator_list.xlsx — it is re-read on every generation run.

def _name_key(name):
    """Order-insensitive name key: 'Smythe, Ashleigh' == 'Ashleigh Smythe'."""
    tokens = [t for t in re.split(r'[,\s]+', str(name).strip().upper())
              if t and t != '.']
    return frozenset(tokens)

def _load_picu_operators():
    """Read the PICU staff list; returns {name_key: display_name}."""
    ops = {}
    try:
        import openpyxl as _oxl
        wb = _oxl.load_workbook(PICU_LIST_PATH, data_only=True)
        ws = wb[wb.sheetnames[0]]
        for row in ws.iter_rows(values_only=True):
            for cell in row:
                if cell is None:
                    continue
                text = str(cell).strip()
                if not text or 'access' in text.lower():
                    continue
                if len(text.split()) >= 2:      # looks like a name
                    ops[_name_key(text)] = text
                break                            # first non-empty cell per row
    except Exception:
        pass
    return ops

# ── Paths ─────────────────────────────────────────────────────────────────────
APP_DIR     = os.path.dirname(os.path.abspath(__file__))
REPORTS_ROOT = os.path.realpath(os.path.join(APP_DIR, '..', 'Reports'))   # shared by both apps
OUTPUT_ROOT  = os.path.join(REPORTS_ROOT, 'ABL')

def _output_dir_for(month_name, month_num, year):
    """Reports\\ABL\\<YYYY-MM Month> under the bundle root — sorts by date in
    Explorer. Inputs are sanitised so a crafted period can't escape OUTPUT_ROOT."""
    try:
        mm = int(re.sub(r'[^0-9]', '', str(month_num)) or 0)
        yy = int(re.sub(r'[^0-9]', '', str(year)) or 0)
    except ValueError:
        mm, yy = 0, 0
    mon = re.sub(r'[^A-Za-z]', '', str(month_name).split(' ')[0])[:12] or 'Month'
    label = f'{yy:04d}-{mm:02d} {mon}'
    return os.path.normpath(os.path.join(OUTPUT_ROOT, label))

HISTORY_PATH    = os.path.join(APP_DIR, 'abl_analyzer_history.json')
OP_HIST_PATH    = os.path.join(APP_DIR, 'abl_operator_history.json')
AUDIT_LOG_PATH  = os.path.join(APP_DIR, 'abl_audit.log')
PICU_LIST_PATH  = os.path.join(APP_DIR, 'picu_operator_list.xlsx')
TEMPLATE_PATH   = os.path.join(APP_DIR, 'template.docx')   # PQ letterhead (logo header/footer)

PORT = 5758   # Separate from iSTAT app (5757)

# ── Audit log ─────────────────────────────────────────────────────────────────
_audit = logging.getLogger('abl.audit')
_audit.setLevel(logging.INFO)
if not _audit.handlers:
    _h = logging.FileHandler(AUDIT_LOG_PATH, encoding='utf-8')
    _h.setFormatter(logging.Formatter('%(asctime)s\t%(message)s'))
    _audit.addHandler(_h)

def _audit_event(action, **fields):
    try:
        user = getpass.getuser()
    except Exception:
        user = 'unknown'
    detail = '\t'.join(f'{k}={v}' for k, v in fields.items())
    _audit.info('user=%s\taction=%s\t%s', user, action, detail)

def _same_origin(req):
    """Reject cross-origin requests (CSRF defence)."""
    origin = req.headers.get('Origin') or req.headers.get('Referer') or ''
    if not origin:
        return True  # non-browser callers (no Origin header) are fine on loopback
    return (origin.startswith(f'http://localhost:{PORT}') or
            origin.startswith(f'http://127.0.0.1:{PORT}'))


# ═══════════════════════════════════════════════════════════════════════════════
#  DOCX HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def _shd(cell, fill):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    old = tcPr.find(qn('w:shd'))
    if old is not None:
        tcPr.remove(old)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  fill)
    tcPr.append(shd)

def _cell_valign(cell, val='center'):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    vA = OxmlElement('w:vAlign')
    vA.set(qn('w:val'), val)
    tcPr.append(vA)

def _cell_width(cell, dxa):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    w = OxmlElement('w:tcW')
    w.set(qn('w:w'),    str(dxa))
    w.set(qn('w:type'), 'dxa')
    existing = tcPr.find(qn('w:tcW'))
    if existing is not None:
        tcPr.remove(existing)
    tcPr.insert(0, w)

def _table_width(table, dxa):
    tblPr = table._tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        table._tbl.insert(0, tblPr)
    w = OxmlElement('w:tblW')
    w.set(qn('w:w'),    str(dxa))
    w.set(qn('w:type'), 'dxa')
    existing = tblPr.find(qn('w:tblW'))
    if existing is not None:
        tblPr.remove(existing)
    tblPr.insert(0, w)

def _add_borders(table):
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    bdr = OxmlElement('w:tblBorders')
    for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'),   'single')
        b.set(qn('w:sz'),    '4')
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), 'auto')
        bdr.append(b)
    existing = tblPr.find(qn('w:tblBorders'))
    if existing is not None:
        tblPr.remove(existing)
    tblPr.append(bdr)

def _set_col_widths(table, widths_dxa):
    _table_width(table, sum(widths_dxa))
    tblGrid = OxmlElement('w:tblGrid')
    for w in widths_dxa:
        gc = OxmlElement('w:gridCol')
        gc.set(qn('w:w'), str(w))
        tblGrid.append(gc)
    existing = table._tbl.find(qn('w:tblGrid'))
    if existing is not None:
        table._tbl.remove(existing)
    table._tbl.insert(1, tblGrid)
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            if i < len(widths_dxa):
                _cell_width(cell, widths_dxa[i])

def _run(para, text, bold=False, italic=False, size=None, color=None, underline=False):
    r = para.add_run(text)
    r.bold   = bold
    r.italic = italic
    if size:
        r.font.size = Pt(size)
    if color:
        r.font.color.rgb = color
    if underline:
        r.font.underline = True
    return r

def _heading(doc, text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.size      = Pt(12)
    r.font.color.rgb = C_BLUE
    r.font.underline = True
    r.bold           = True
    return p

def _hdr_cell(cell, text, fill, text_color=None):
    _shd(cell, fill)
    _cell_valign(cell, 'center')
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.clear()
    r = p.add_run(text)
    r.bold           = True
    r.font.size      = Pt(9)
    r.font.color.rgb = text_color or C_WHITE

def _data_cell(cell, text, fill=FILL_WHITE, color=None, bold=False,
               align=WD_ALIGN_PARAGRAPH.LEFT, size=9):
    _shd(cell, fill)
    p = cell.paragraphs[0]
    p.alignment = align
    p.clear()
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.bold      = bold
    if color:
        r.font.color.rgb = color
    return r

def get_perf_label(rate):
    if rate < 4:   return 'Excellent',       RGBColor(0x53, 0x81, 0x35)
    if rate <= 6:  return 'Acceptable',      RGBColor(0x2F, 0x54, 0x96)
    if rate <= 10: return 'Monitor',         RGBColor(0xFF, 0x80, 0x00)
    return             'Needs attention',    RGBColor(0xFF, 0x00, 0x00)


# ═══════════════════════════════════════════════════════════════════════════════
#  EXCEL PARSING
# ═══════════════════════════════════════════════════════════════════════════════

def _find_header_row(ws, max_scan=60):
    """
    Locate the data table header by finding a row that contains at least 2
    of the known column-name markers. The top of the ABL export contains
    error-code definition grids, so a single-word match is not reliable.
    Falls back to any row containing 'hospital'.
    """
    markers = {'hospital', 'operator', 'department', 'analyzer', 'analyser'}
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True)):
        row_lc = {str(c).strip().lower() for c in row if c}
        if len(row_lc & markers) >= 2:
            return i + 1
    # Single-word fallback
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True)):
        if any(c and str(c).strip().lower() == 'hospital' for c in row):
            return i + 1
    return None

def _norm_header(h):
    """Collapse line breaks / repeated whitespace: 'Error\\n328' -> 'error 328'."""
    return re.sub(r'\s+', ' ', str(h)).strip().lower()

def load_data(excel_path):
    wb = openpyxl.load_workbook(excel_path, data_only=True)

    # The workbook has three tabs: 'Operator error ABL', 'errors greater than 10%',
    # 'operators for follow up'.  All reporting is driven from the first —
    # find it by name, fall back to the first sheet.
    ws = None
    for name in wb.sheetnames:
        if 'operator error' in name.strip().lower():
            ws = wb[name]
            break
    if ws is None:
        ws = wb[wb.sheetnames[0]]

    hdr_row = _find_header_row(ws)
    if hdr_row is None:
        raise ValueError(
            "Could not find the data table.  Make sure the spreadsheet contains "
            "columns named Hospital, Operator, Department, and Analyzer."
        )

    rows = list(ws.iter_rows(min_row=hdr_row, values_only=True))
    raw_headers = [str(h).strip() if h else f'_blank{i}' for i, h in enumerate(rows[0])]

    data_rows = [r for r in rows[1:] if any(v is not None for v in r)]
    if not data_rows:
        raise ValueError("No data rows found below the header row.")

    df = pd.DataFrame([dict(zip(raw_headers, r)) for r in data_rows])

    # ── Map raw headers → canonical names ─────────────────────────────────
    # Headers in the ABL export wrap over two lines ('Error\n328'), so all
    # matching is done on whitespace-normalised names.
    col_map = {}
    for col in df.columns:
        nc = _norm_header(col)
        if nc == 'hospital':
            col_map[col] = 'Hospital'
        elif nc == 'operator':
            col_map[col] = 'Operator'
        elif nc in ('department', 'dept', 'department name'):
            col_map[col] = 'Department'
        elif nc in ('analyzer', 'analyser', 'analyzer name', 'analyser name', 'instrument'):
            col_map[col] = 'Analyzer'
        elif 'tests by operator' in nc or nc in ('tests by op', 'operator tests', 'tests'):
            col_map[col] = 'Tests'
        elif ('pct' in nc and 'error' in nc) or nc in ('% error', 'error %', 'error rate'):
            col_map[col] = 'PctError'
        elif 'total tests on' in nc or 'total tests' in nc:
            col_map[col] = 'TotalTests'
        else:
            m = re.match(r'^error\s*(\d{3})$', nc)
            if m and m.group(1) in ERROR_CODES:
                col_map[col] = f'E{m.group(1)}'

    df = df.rename(columns=col_map)

    # The real export's Hospital header cell is a broken formula ('#NAME?').
    # Hospital is always the first column — fall back to it by position.
    if 'Hospital' not in df.columns and len(df.columns):
        df = df.rename(columns={df.columns[0]: 'Hospital'})

    if not any(f'E{c}' in df.columns for c in ERROR_CODES):
        raise ValueError(
            "None of the Error 328–791 columns were recognised — "
            "check the 'Operator error ABL' tab headers."
        )

    # ── Forward-fill grouped text columns ──────────────────────────────────
    # The export writes Hospital once at the top and leaves Operator/Department
    # blank on continuation rows (an operator's 2nd analyzer).  Strip strings,
    # blank→NaN, then fill down.
    for col in ('Hospital', 'Operator', 'Department', 'Analyzer'):
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df.loc[df[col].isin(['', 'None', 'nan']), col] = pd.NA
    for col in ('Hospital', 'Operator', 'Department'):
        if col in df.columns:
            df[col] = df[col].ffill()

    # ── Ensure all error columns exist and are numeric ─────────────────────
    for code in ERROR_CODES:
        ec = f'E{code}'
        if ec not in df.columns:
            df[ec] = 0
        df[ec] = pd.to_numeric(df[ec], errors='coerce').fillna(0).astype(int)

    for col in ('Tests', 'TotalTests'):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
        else:
            df[col] = 0

    if 'PctError' in df.columns:
        raw_pct  = df['PctError'].astype(str).str.strip()
        # Text like '3.51%' → already percent units.  Plain numerics that are
        # all <= 1 are Excel percent-formatted fractions (0.0351 == 3.51%).
        had_sign = raw_pct.str.endswith('%').any()
        df['PctError'] = pd.to_numeric(raw_pct.str.rstrip('%'),
                                       errors='coerce').fillna(0.0)
        if (not had_sign and df['PctError'].max() <= 1.0
                and df['PctError'].max() > 0):
            df['PctError'] = df['PctError'] * 100
        df['PctError'] = df['PctError'].round(2)
    else:
        df['PctError'] = 0.0

    df['TotalErrors'] = df[[f'E{c}' for c in ERROR_CODES]].sum(axis=1)

    # Operator error rate: use the sheet's Pct. error column; fall back to
    # computing it from the raw error counts if the column is empty.
    df['ErrorRate'] = df.apply(
        lambda r: r['PctError'] if r['PctError'] > 0
        else (round(r['TotalErrors'] / r['Tests'] * 100, 2) if r['Tests'] > 0 else 0.0),
        axis=1
    )

    # Flag rule: operators with MIN_TESTS tests or fewer are excluded
    # (low volume); above that, a Pct. error over FLAG_PCT is flagged.
    df['Flagged'] = (df['Tests'] > MIN_TESTS) & (df['ErrorRate'] > FLAG_PCT)

    # Drop blank rows
    for col in ('Operator', 'Hospital', 'Analyzer'):
        if col in df.columns:
            df = df[df[col].notna() & (df[col].astype(str).str.strip() != '')]

    # Department is a grouping level — normalise blanks so every analyzer
    # lands under a department header in the selection tree.
    if 'Department' not in df.columns:
        df['Department'] = 'Unassigned'
    else:
        df['Department'] = df['Department'].astype(str).str.strip()
        df.loc[df['Department'].isin(['', 'None', 'nan']), 'Department'] = 'Unassigned'

    return df


# ═══════════════════════════════════════════════════════════════════════════════
#  HISTORY
# ═══════════════════════════════════════════════════════════════════════════════

def _load_json(path):
    if os.path.exists(path):
        try:
            with open(path, encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def _save_json(path, data):
    try:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass

def _dept_key(hospital, department):
    """Report unit key: Hospital | Department."""
    return f"{str(hospital).strip()}|{str(department).strip()}"

def _anlz_key(hospital, department, analyzer):
    """Trend-history key for one analyzer within a department report."""
    return (f"{str(hospital).strip()}|{str(department).strip()}"
            f"|{str(analyzer).strip()}")

def _update_analyzer_history(history, akey, year, month_num, error_rate,
                             total_tests=0, codes=None):
    entries = [e for e in history.get(akey, [])
               if not (e['year'] == int(year) and e['month'] == int(month_num))]
    entries.append({'year': int(year), 'month': int(month_num),
                    'error_rate': round(float(error_rate), 2),
                    'total_tests': int(total_tests),
                    'codes': {str(c): int(n) for c, n in (codes or {}).items() if n}})
    entries.sort(key=lambda e: (e['year'], e['month']))
    history[akey] = entries[-24:]
    return history

def _month_idx(year, month):
    """Absolute month number, for lookback-window arithmetic."""
    return int(year) * 12 + int(month) - 1

def _update_op_history(op_hist, akey, operator, year, month_num, error_rate,
                       total_errors, flagged, codes=None):
    # Only flagged operators (>MIN_TESTS tests, >FLAG_PCT% error) are recorded,
    # so every entry represents a flagged month.
    if not flagged:
        return op_hist
    key     = f"{akey}|{str(operator).strip().upper()}"
    entries = [e for e in op_hist.get(key, [])
               if not (e['year'] == int(year) and e['month'] == int(month_num))]
    entries.append({
        'year':         int(year),
        'month':        int(month_num),
        'error_rate':   round(float(error_rate), 2),
        'total_errors': int(total_errors),
        'operator':     str(operator).strip(),
        'codes':        {str(c): int(n) for c, n in (codes or {}).items() if n},
    })
    entries.sort(key=lambda e: (e['year'], e['month']))
    op_hist[key] = entries[-24:]
    return op_hist


# ═══════════════════════════════════════════════════════════════════════════════
#  TREND CHART
# ═══════════════════════════════════════════════════════════════════════════════

def _plot_error_trend(entries, label):
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        import matplotlib.patches as mpatches
        import calendar

        if not entries:
            return None

        labels = [f"{calendar.month_abbr[e['month']]}\n{str(e['year'])[2:]}" for e in entries]
        values = [e['error_rate'] for e in entries]

        fig, ax = plt.subplots(figsize=(9, 3.5))
        fig.patch.set_facecolor('white')
        ax.set_facecolor('white')

        y_ceil = max(max(values) * 1.35, 14)
        ax.axhspan(0,  4,      alpha=0.07, color='#538135', zorder=0)
        ax.axhspan(4,  6,      alpha=0.07, color='#2F5496', zorder=0)
        ax.axhspan(6,  10,     alpha=0.07, color='#FF8000', zorder=0)
        ax.axhspan(10, y_ceil, alpha=0.07, color='#FF0000', zorder=0)

        for thresh, col in [(4, '#538135'), (6, '#2F5496'), (10, '#FF0000')]:
            ax.axhline(thresh, color=col, linewidth=0.8, linestyle='--', alpha=0.55)

        xs = list(range(len(labels)))
        ax.plot(xs, values, color='#305496', linewidth=2.2, marker='o',
                markersize=6, markerfacecolor='white',
                markeredgecolor='#305496', markeredgewidth=2, zorder=3)

        for xi, yi in zip(xs, values):
            _, lc = get_perf_label(yi)
            rgb = (lc[0] / 255, lc[1] / 255, lc[2] / 255)
            ax.annotate(f'{yi:.1f}%', (xi, yi),
                        textcoords='offset points', xytext=(0, 9),
                        ha='center', fontsize=7.5, color=rgb, fontweight='bold')

        ax.set_xticks(xs)
        ax.set_xticklabels(labels, fontsize=8)
        ax.set_ylabel('Error Rate %', fontsize=9)
        ax.set_ylim(0, y_ceil)
        ax.set_title(f'Operator Error Rate Trend  —  {label}',
                     fontsize=10, color='#1F497D', fontweight='bold', pad=10)
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f'{v:.0f}%'))
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.tick_params(axis='both', labelsize=8)

        legend = [
            mpatches.Patch(facecolor='#538135', alpha=0.4, label='Excellent  <4%'),
            mpatches.Patch(facecolor='#2F5496', alpha=0.4, label='Acceptable  4–6%'),
            mpatches.Patch(facecolor='#FF8000', alpha=0.4, label='Monitor  6–10%'),
            mpatches.Patch(facecolor='#FF0000', alpha=0.4, label='Needs attention  >10%'),
        ]
        ax.legend(handles=legend, loc='upper right', fontsize=7,
                  framealpha=0.85, ncol=2)

        plt.tight_layout()
        buf = io.BytesIO()
        plt.savefig(buf, format='png', dpi=150, bbox_inches='tight',
                    facecolor='white', edgecolor='none')
        plt.close(fig)
        buf.seek(0)
        return buf

    except ImportError:
        return None
    except Exception:
        return None


def _plot_volume_trend(entries, label):
    """Two-panel chart: tests-per-month line on top, and beneath it stacked
    monthly bars of each error type as % of that month's tests — one colour
    per error code, legend under the chart."""
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        import matplotlib.patches as mpatches
        import calendar

        pts = [e for e in entries if int(e.get('total_tests', 0) or 0) > 0]
        if not pts:
            return None

        labels = [f"{calendar.month_abbr[e['month']]}\n{str(e['year'])[2:]}" for e in pts]
        values = [int(e['total_tests']) for e in pts]
        xs     = list(range(len(labels)))

        # Per-month error-type % of tests, in fixed code order
        codes_present = [c for c in ERROR_CODES
                         if any((e.get('codes') or {}).get(c) for e in pts)]
        pct = {c: [round((e.get('codes') or {}).get(c, 0) / e['total_tests'] * 100, 3)
                   for e in pts] for c in codes_present}

        fig, ax = plt.subplots(figsize=(9, 3.9))
        fig.patch.set_facecolor('white')
        ax.set_facecolor('white')
        ax2 = ax.twinx()

        # ── Error-mix bars on the RIGHT axis (% of that month's tests) ────
        month_tot = []
        for i, e in enumerate(pts):
            codes  = {c: n for c, n in (e.get('codes') or {}).items() if n}
            m_errs = sum(codes.values())
            month_tot.append(m_errs / e['total_tests'] * 100 if m_errs else 0.0)
        r_max = max(max(month_tot) * 2.2, 1.0)   # keep bars in the lower zone
        ax2.set_ylim(0, r_max)

        # Dark text on the light segment colours, white on the dark ones
        _DARK_TEXT = {'593': '#4A3000', '722': '#5A1030'}

        for i, e in enumerate(pts):
            codes  = {c: n for c, n in (e.get('codes') or {}).items() if n}
            m_errs = sum(codes.values())
            if not m_errs:
                continue
            bottom = 0.0
            for c in ERROR_CODES:
                n = codes.get(c, 0)
                if not n:
                    continue
                seg = n / e['total_tests'] * 100
                ax2.bar([i], [seg], bottom=[bottom], width=0.5,
                        color=ERROR_COLORS.get(c, '#888780'), zorder=2,
                        edgecolor='white', linewidth=0.5)
                # Percentage inside the segment (skipped if too thin to read)
                if seg / r_max >= 0.055:
                    ax2.annotate(f'{seg:.2f}%', (i, bottom + seg / 2),
                                 ha='center', va='center', fontsize=6.2,
                                 fontweight='bold', zorder=3,
                                 color=_DARK_TEXT.get(c, 'white'))
                bottom += seg
            # Total for the month above the stack
            ax2.annotate(f'{month_tot[i]:.1f}%', (i, bottom),
                         textcoords='offset points', xytext=(0, 4),
                         ha='center', fontsize=7, color='#444444',
                         fontweight='bold', zorder=3)

        ax2.set_ylabel('Errors, % of tests', fontsize=9, color='#7A7A7A')
        ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f'{v:g}%'))
        ax2.tick_params(axis='y', labelsize=8, colors='#7A7A7A')
        ax2.spines['top'].set_visible(False)

        # ── Tests-per-month line on the LEFT axis, drawn above the bars ───
        ax.set_ylim(0, max(values) * 1.30)
        ax.plot(xs, values, color='#538135', linewidth=2.2, marker='o',
                markersize=6, markerfacecolor='white',
                markeredgecolor='#538135', markeredgewidth=2, zorder=4)
        for xi, yi in zip(xs, values):
            ax.annotate(f'{yi:,}', (xi, yi),
                        textcoords='offset points', xytext=(0, 9),
                        ha='center', fontsize=7.5, color='#538135',
                        fontweight='bold')
        ax.set_zorder(ax2.get_zorder() + 1)   # line layer above the bars
        ax.patch.set_visible(False)

        ax.set_ylabel('Tests run', fontsize=9, color='#538135')
        ax.set_title(f'Tests Run & Error Mix  —  {label}',
                     fontsize=10, color='#1F497D', fontweight='bold', pad=10)
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f'{v:,.0f}'))
        ax.set_xticks(xs)
        ax.set_xticklabels(labels, fontsize=8)
        ax.spines['top'].set_visible(False)
        ax.tick_params(axis='y', labelsize=8, colors='#538135')
        ax.tick_params(axis='x', labelsize=8)
        ax.grid(axis='y', color='#D9D9D9', linewidth=0.6, alpha=0.6, zorder=0)

        # ── Legend under the chart: one entry per error code present ──────
        if codes_present:
            handles = [mpatches.Patch(
                           facecolor=ERROR_COLORS.get(c, '#888780'),
                           label=f'{c}  {ABL_ERRORS.get(c, "")[:26]}')
                       for c in codes_present]
            fig.legend(handles=handles, loc='lower center',
                       ncol=min(3, len(handles)), fontsize=7,
                       frameon=False, bbox_to_anchor=(0.5, -0.01))
            fig.subplots_adjust(bottom=0.26)

        buf = io.BytesIO()
        plt.savefig(buf, format='png', dpi=150, bbox_inches='tight',
                    facecolor='white', edgecolor='none')
        plt.close(fig)
        buf.seek(0)
        return buf

    except ImportError:
        return None
    except Exception:
        return None


# ═══════════════════════════════════════════════════════════════════════════════
#  REPORT GENERATION
# ═══════════════════════════════════════════════════════════════════════════════

def _render_op_rows(op_tbl, op_grp):
    """Fill an operator table (built by _operator_table) with rows + TOTAL."""
    for ri, (_, row) in enumerate(op_grp.iterrows()):
        fill       = FILL_BLUE_HDR if ri % 2 == 1 else FILL_WHITE
        op_name    = str(row.get('Operator', '')).strip()
        er         = float(row.get('ErrorRate', 0))
        tests      = int(row.get('Tests', 0))
        is_flagged = bool(row.get('Flagged', False))
        low_vol    = 0 < tests <= MIN_TESTS
        _, lc      = get_perf_label(er)
        if is_flagged:
            hi_col = C_RED_TEXT
        elif low_vol:
            hi_col = C_GREY_TEXT      # low volume — not assessed
            lc     = C_GREY_TEXT
        else:
            hi_col = C_BLUE

        disp_name = op_name
        if is_flagged:
            disp_name += '  ⚠'
        elif low_vol:
            disp_name += '  *'
        dr = op_tbl.add_row().cells
        _data_cell(dr[0], disp_name, fill, hi_col, bold=is_flagged, size=8)
        _data_cell(dr[1], str(tests), fill, C_BLUE, size=8,
                   align=WD_ALIGN_PARAGRAPH.CENTER)

        for j, code in enumerate(ERROR_CODES):
            cnt   = int(row.get(f'E{code}', 0))
            ecol  = C_RED_TEXT if cnt > 0 else C_GREY_TEXT
            _data_cell(dr[2 + j], str(cnt) if cnt else '—', fill, ecol, size=8,
                       align=WD_ALIGN_PARAGRAPH.CENTER)

        _data_cell(dr[9], str(int(row.get('TotalErrors', 0))), fill, hi_col, bold=True, size=8,
                   align=WD_ALIGN_PARAGRAPH.CENTER)

        _shd(dr[10], fill)
        p5 = dr[10].paragraphs[0]; p5.clear()
        p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r5 = p5.add_run(f'{er:.1f}%')
        r5.font.size = Pt(8); r5.bold = True; r5.font.color.rgb = lc


def generate_report(hospital, department, df_dept, report_month,
                    report_month_num, report_year, history=None, op_history=None,
                    picu_ops=None):
    """One report per hospital+department.  Every analyzer belonging to the
    department appears inside the report with its own error details."""
    if df_dept.empty:
        return None

    dkey = _dept_key(hospital, department)

    # ── PICU cross-reference (ICU reports only) ────────────────────────────
    # Analyzer/department totals always include every row; the split only
    # affects which operator table a person is listed in.
    is_picu = None
    picu_names = []
    if picu_ops and department.strip().upper() == 'ICU' and 'Operator' in df_dept.columns:
        is_picu = df_dept['Operator'].apply(lambda o: _name_key(o) in picu_ops)
        if is_picu.any():
            picu_names = sorted(df_dept[is_picu]['Operator'].astype(str).str.strip().unique())
        else:
            is_picu = None
    df_picu = df_dept[is_picu] if is_picu is not None else None
    df_main = df_dept[~is_picu] if is_picu is not None else df_dept

    # ── Per-analyzer stats ─────────────────────────────────────────────────
    # 'Total tests on analyzer' repeats on every row for that analyzer → MAX;
    # analyzer error % = sum of the seven error columns ÷ that total.
    analyzer_stats = []
    for anlz, grp in df_dept.groupby('Analyzer'):
        an_total = int(grp['TotalTests'].max()) if 'TotalTests' in grp.columns else 0
        if an_total <= 0:
            an_total = int(grp['Tests'].sum())   # fallback if column missing
        an_errors = int(grp['TotalErrors'].sum())
        an_rate   = round(an_errors / an_total * 100, 2) if an_total > 0 else 0.0
        analyzer_stats.append({'analyzer': str(anlz).strip(), 'total': an_total,
                               'errors': an_errors, 'rate': an_rate,
                               'rows': grp})
    analyzer_stats.sort(key=lambda a: a['analyzer'])

    # ── Department-level metrics (across all its analyzers) ────────────────
    total_tests  = sum(a['total']  for a in analyzer_stats)
    total_errors = sum(a['errors'] for a in analyzer_stats)
    dept_rate    = round(total_errors / total_tests * 100, 2) if total_tests > 0 else 0.0
    dept_label, dept_lc = get_perf_label(dept_rate)

    error_totals = {code: int(df_dept[f'E{code}'].sum()) for code in ERROR_CODES}
    top_code     = max(error_totals, key=error_totals.get) if any(error_totals.values()) else None

    n_operators = (df_dept[df_dept['TotalErrors'] > 0]['Operator'].nunique()
                   if 'Operator' in df_dept.columns else 0)

    # Operators flagged THIS month: Tests > MIN_TESTS and Pct. error > FLAG_PCT
    cur_flagged = (sorted(df_dept[df_dept['Flagged']]['Operator'].astype(str).str.strip().unique())
                   if 'Flagged' in df_dept.columns and 'Operator' in df_dept.columns else [])
    low_volume  = (df_dept[(df_dept['Tests'] > 0) & (df_dept['Tests'] <= MIN_TESTS)]
                   ['Operator'].nunique()
                   if 'Operator' in df_dept.columns else 0)

    # 12-month lookback: operators flagged (>FLAG_PCT% with >MIN_TESTS tests)
    # in two or more months within the window go to the review table.
    LOOKBACK_MONTHS = 12
    rep_idx    = _month_idx(report_year, report_month_num)
    review_ops = []          # (operator name, [flagged-month entries in window])
    if op_history:
        prefix = dkey + '|'
        for key, entries in op_history.items():
            if not key.startswith(prefix):
                continue
            recent = [e for e in entries
                      if 0 <= rep_idx - _month_idx(e['year'], e['month'])
                      < LOOKBACK_MONTHS]
            if len(recent) >= 2:
                review_ops.append((recent[-1]['operator'], recent))
        review_ops.sort()
    flagged_count = len(review_ops)

    # ── Build document from the PQ template (keeps logo header/footer) ─────
    if os.path.exists(TEMPLATE_PATH):
        doc  = Document(TEMPLATE_PATH)
        body = doc.element.body
        sect = body.find(qn('w:sectPr'))
        for child in list(body):
            if child != sect:
                body.remove(child)
        # Force page background to white (overrides any theme default)
        root = doc.element
        old_bg = root.find(qn('w:background'))
        if old_bg is not None:
            root.remove(old_bg)
        bg = OxmlElement('w:background')
        bg.set(qn('w:color'), 'FFFFFF')
        bg.set(qn('w:themeColor'), 'background1')
        root.insert(0, bg)
    else:
        doc = Document()
        for section in doc.sections:
            section.top_margin    = Inches(1.0)
            section.bottom_margin = Inches(1.0)
            section.left_margin   = Inches(1.0)
            section.right_margin  = Inches(1.0)

    # ── Title block (indented to clear the logo on the left) ───────────────
    LOGO_INDENT = Inches(2127 / 1440)

    def _titled(style):
        try:
            return doc.add_paragraph(style=style)
        except Exception:
            return doc.add_paragraph()

    p = _titled('Title')
    p.paragraph_format.left_indent = LOGO_INDENT
    _run(p, 'Pathology Queensland', size=12, color=C_DARK_BLUE)

    p = _titled('Title')
    p.paragraph_format.left_indent = LOGO_INDENT
    _run(p, 'Radiometer ABL Summary Report', size=20, color=C_DARK_BLUE)

    p = _titled('Subtitle')
    p.paragraph_format.first_line_indent = LOGO_INDENT
    r = _run(p, f'{hospital} · {department} – {report_month}', size=18, color=C_GREEN)
    try:
        r.style = doc.styles['DocSubTitle']   # footer STYLEREF picks this up
    except Exception:
        pass

    doc.add_paragraph()

    # ── Monthly Snapshot ───────────────────────────────────────────────────
    _heading(doc, 'Monthly Snapshot')

    snap = doc.add_table(rows=0, cols=3)
    _add_borders(snap)
    _set_col_widths(snap, [2500, 4200, 2300])

    hdr = snap.add_row().cells
    for c, lbl in zip(hdr, ['Indicator', 'Value', 'Status']):
        _hdr_cell(c, lbl, FILL_GREY_HDR)

    def _snap_row(label, value_str, status_str, status_color=None, fill=FILL_WHITE):
        r = snap.add_row().cells
        _data_cell(r[0], label,      fill, C_BLUE,              bold=True, size=9)
        _data_cell(r[1], value_str,  fill, C_BLUE,              size=9)
        _data_cell(r[2], status_str, fill, status_color or C_BLUE, bold=True, size=9)

    _snap_row('Department Error Rate',
              f'{dept_rate:.2f}%  ({total_errors} errors / {total_tests} tests '
              f'across {len(analyzer_stats)} analyzer{"s" if len(analyzer_stats) != 1 else ""})',
              dept_label, dept_lc, fill=FILL_ALT_ROW)

    _snap_row('Analyzers',
              ',  '.join(a['analyzer'] for a in analyzer_stats), '')

    if top_code:
        top_desc = ABL_ERRORS.get(top_code, '')[:50]
        _snap_row('Top Error Code',
                  f'Error {top_code}: {top_desc}  ({error_totals[top_code]} occurrences)',
                  'Most frequent', C_BLUE, fill=FILL_ALT_ROW)
    else:
        _snap_row('Top Error Code', 'No errors recorded', '✓ Clear', C_GREEN_PASS,
                  fill=FILL_ALT_ROW)

    op_detail = f'{n_operators} operator{"s" if n_operators != 1 else ""} this period'
    if low_volume:
        op_detail += f'  ({low_volume} with {MIN_TESTS} tests or fewer, not assessed)'
    _snap_row('Operators with Errors', op_detail, '')

    if picu_names:
        _snap_row('PICU Operators',
                  f'{len(picu_names)} PICU staff recorded errors on ICU '
                  'analysers — listed separately below',
                  'See PICU section', C_BLUE)

    # Intervention rows appear ONLY when operators are flagged in the
    # current month — a clean month shows no intervention call-outs.
    if cur_flagged:
        _flag_disp = [n + ('  (PICU)' if n in picu_names else '')
                      for n in cur_flagged[:6]]
        _snap_row('Operators Flagged This Month',
                  f'{len(cur_flagged)} operator{"s" if len(cur_flagged) != 1 else ""} '
                  f'over {FLAG_PCT:.0f}% error (more than {MIN_TESTS} tests): '
                  + ', '.join(_flag_disp)
                  + ('…' if len(cur_flagged) > 6 else ''),
                  '⚠ Review required', C_RED_TEXT, fill=FILL_ALT_ROW)

        if flagged_count:
            _snap_row('Recurring Alerts',
                      f'{flagged_count} operator{"s" if flagged_count != 1 else ""} '
                      f'flagged in 2+ months within the last {LOOKBACK_MONTHS} months',
                      '⚠ Review required', C_RED_TEXT)

    doc.add_paragraph()

    # ── Historical Performance ─────────────────────────────────────────────
    _heading(doc, 'Historical Performance')

    plotted = False
    if history:
        for an in analyzer_stats:
            entries = history.get(_anlz_key(hospital, department, an['analyzer']))
            if not entries:
                continue
            chart = _plot_error_trend(entries, f"{an['analyzer']} — {department}")
            if chart:
                p = doc.add_paragraph()
                p.add_run().add_picture(chart, width=Inches(6.2))
                plotted = True
            vol = _plot_volume_trend(entries, f"{an['analyzer']} — {department}")
            if vol:
                p = doc.add_paragraph()
                p.add_run().add_picture(vol, width=Inches(6.2))
                plotted = True
    if not plotted:
        p = doc.add_paragraph()
        _run(p, 'Trend data will appear here as monthly reports are generated.',
             size=9, color=C_GREY_TEXT)

    doc.add_paragraph()

    # ── Analyzer Performance ───────────────────────────────────────────────
    _heading(doc, 'Analyzer Performance')

    an_tbl = doc.add_table(rows=1, cols=4)
    _add_borders(an_tbl)
    _set_col_widths(an_tbl, [3200, 2100, 1800, 2000])
    for c, lbl in zip(an_tbl.rows[0].cells,
                       ['Analyzer', 'Total Tests on Analyzer', 'Total Errors', 'Error %']):
        _hdr_cell(c, lbl, FILL_GREY_HDR)

    for ri, an in enumerate(analyzer_stats):
        fill = FILL_BLUE_HDR if ri % 2 == 1 else FILL_WHITE
        albl, alc = get_perf_label(an['rate'])
        dr = an_tbl.add_row().cells
        _data_cell(dr[0], an['analyzer'],    fill, C_BLUE, size=9)
        _data_cell(dr[1], str(an['total']),  fill, C_BLUE, size=9,
                   align=WD_ALIGN_PARAGRAPH.CENTER)
        _data_cell(dr[2], str(an['errors']), fill, C_BLUE, size=9,
                   align=WD_ALIGN_PARAGRAPH.CENTER)
        _shd(dr[3], fill)
        p3 = dr[3].paragraphs[0]; p3.clear()
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r3 = p3.add_run(f'{an["rate"]:.2f}%  {albl}')
        r3.font.size = Pt(9); r3.bold = True; r3.font.color.rgb = alc

    p = doc.add_paragraph()
    _run(p, 'Analyzer error % = sum of error columns ÷ the "Total tests on '
            'analyzer" value for that analyzer.',
         size=8, color=C_GREY_TEXT)
    doc.add_paragraph()

    # ── Operator Error Details — one section per analyzer ──────────────────
    _heading(doc, 'Operator Error Details by Analyzer')

    def _operator_table(df_rows):
        """Render the operator table for one analyzer's rows."""
        # Cols: Operator | Tests | E328..E791 (×7) | Total Errors | Error %
        COL_W = [2900, 800] + [480] * 7 + [820, 820]
        op_tbl = doc.add_table(rows=1, cols=len(COL_W))
        _add_borders(op_tbl)
        _set_col_widths(op_tbl, COL_W)

        hdr_lbls = ['Operator', 'Tests'] + [f'E{c}' for c in ERROR_CODES] + ['Total\nErrors', 'Error %']
        for i, (c, lbl) in enumerate(zip(op_tbl.rows[0].cells, hdr_lbls)):
            fill = FILL_RED_HDR if 2 <= i <= 8 else FILL_GREY_HDR
            _hdr_cell(c, lbl, fill)

        agg_dict = {'Tests': ('Tests', 'sum'), 'TotalErrors': ('TotalErrors', 'sum'),
                    'PctError': ('PctError', 'max')}
        agg_dict.update({f'E{c}': (f'E{c}', 'sum') for c in ERROR_CODES})
        op_grp = df_rows.groupby('Operator').agg(**agg_dict).reset_index()
        op_grp['ErrorRate'] = op_grp.apply(
            lambda r: r['PctError'] if r['PctError'] > 0
            else (round(r['TotalErrors'] / r['Tests'] * 100, 2) if r['Tests'] > 0 else 0.0),
            axis=1
        )
        op_grp['Flagged'] = (op_grp['Tests'] > MIN_TESTS) & (op_grp['ErrorRate'] > FLAG_PCT)
        # Flagged operators first, then by error rate
        op_grp = op_grp.sort_values(['Flagged', 'ErrorRate'], ascending=[False, False])
        return op_tbl, op_grp

    for an_i, an in enumerate(analyzer_stats):
        albl, alc = get_perf_label(an['rate'])
        ph = doc.add_paragraph()
        _run(ph, f'{an["analyzer"]}', bold=True, size=11, underline=True, color=C_BLUE)
        _run(ph, f'   {an["total"]} tests on analyzer  ·  {an["errors"]} errors  ·  ',
             size=9, color=C_GREY_TEXT)
        _run(ph, f'{an["rate"]:.2f}% {albl}', bold=True, size=9, color=alc)

        # PICU staff are listed in their own section below, not here —
        # analyzer totals above still include their errors.
        an_rows = (df_main[df_main['Analyzer'].astype(str).str.strip()
                           == an['analyzer']]
                   if df_picu is not None else an['rows'])
        if an_rows.empty and df_picu is not None:
            p = doc.add_paragraph()
            _run(p, 'All operators recorded on this analyser are PICU staff — '
                    'see the PICU section below.',
                 size=9, color=C_GREY_TEXT)
        else:
            op_tbl, op_grp = _operator_table(an_rows)
            _render_op_rows(op_tbl, op_grp)
        doc.add_paragraph()

    # ── PICU Operators — separate list on the ICU report ───────────────────
    if df_picu is not None and not df_picu.empty:
        _heading(doc, 'PICU Operators — ICU Analysers')
        p = doc.add_paragraph()
        _run(p,
             'The following staff appear under ICU in the analyser export but '
             'are PICU employees (per the PICU operator list).  Their errors '
             'are included in the analyser totals above and are listed '
             'separately here:',
             size=9, color=C_GREY_TEXT)

        picu_tbl, picu_grp = _operator_table(df_picu)
        _render_op_rows(picu_tbl, picu_grp)
        doc.add_paragraph()

    p = doc.add_paragraph()
    _run(p, 'Operator error rate:  ', bold=True, size=9)
    _run(p, 'Excellent', size=9, color=RGBColor(0x53, 0x81, 0x35)); _run(p, ' <4%   ', size=9)
    _run(p, 'Acceptable', size=9, color=RGBColor(0x2F, 0x54, 0x96)); _run(p, ' 4–6%   ', size=9)
    _run(p, 'Monitor', size=9, color=RGBColor(0xFF, 0x80, 0x00)); _run(p, ' 6–10%   ', size=9)
    _run(p, 'Needs attention', size=9, color=RGBColor(0xFF, 0x00, 0x00)); _run(p, ' >10%', size=9)

    p = doc.add_paragraph()
    _run(p, '⚠ ', size=9, color=C_RED_TEXT)
    _run(p, f'flagged: over {FLAG_PCT:.0f}% error with more than {MIN_TESTS} tests.    ', size=8, color=C_GREY_TEXT)
    _run(p, '* ', size=9)
    _run(p, f'low volume: {MIN_TESTS} tests or fewer — not assessed.', size=8, color=C_GREY_TEXT)

    doc.add_paragraph()

    # ── Error Type Summary (department-wide) ───────────────────────────────
    _heading(doc, 'Error Type Summary')

    et = doc.add_table(rows=1, cols=4)
    _add_borders(et)
    _set_col_widths(et, [900, 5400, 900, 900])
    for c, lbl in zip(et.rows[0].cells, ['Code', 'Description', 'Count', '% of Errors']):
        _hdr_cell(c, lbl, FILL_BLUE_HDR, C_BLUE)

    sorted_errs = sorted(error_totals.items(), key=lambda x: x[1], reverse=True)
    for ri, (code, cnt) in enumerate(sorted_errs):
        fill = FILL_ALT_ROW if ri % 2 == 0 else FILL_WHITE
        pct  = round(cnt / total_errors * 100, 1) if total_errors else 0.0
        col  = C_RED_TEXT if cnt > 0 else C_GREY_TEXT
        dr   = et.add_row().cells
        _data_cell(dr[0], code,                    fill, col, bold=(cnt > 0), size=9,
                   align=WD_ALIGN_PARAGRAPH.CENTER)
        _data_cell(dr[1], ABL_ERRORS.get(code,''), fill, col, size=9)
        _data_cell(dr[2], str(cnt),                fill, col, size=9,
                   align=WD_ALIGN_PARAGRAPH.CENTER)
        _data_cell(dr[3], f'{pct:.1f}%',           fill, col, size=9,
                   align=WD_ALIGN_PARAGRAPH.CENTER)

    tr = et.add_row().cells
    _data_cell(tr[0], 'Total', FILL_ALT_ROW, C_BLUE, bold=True, size=9)
    _data_cell(tr[1], '', FILL_ALT_ROW)
    _data_cell(tr[2], str(total_errors), FILL_ALT_ROW, C_BLUE, bold=True, size=9,
               align=WD_ALIGN_PARAGRAPH.CENTER)
    _data_cell(tr[3], '100%', FILL_ALT_ROW, C_BLUE, bold=True, size=9,
               align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_paragraph()

    # ── Operators Requiring Review — 12-month lookback ─────────────────────
    if review_ops:
        import calendar as _cal
        _heading(doc, 'Operators Requiring Review')
        p = doc.add_paragraph()
        _run(p,
             f'The following operators have exceeded {FLAG_PCT:.0f}% error '
             f'(with more than {MIN_TESTS} tests) in two or more months within '
             f'the last {LOOKBACK_MONTHS} months and may benefit from '
             'additional training or review:',
             size=9, color=C_GREY_TEXT)

        fu = doc.add_table(rows=1, cols=6)
        _add_borders(fu)
        _set_col_widths(fu, [2100, 1250, 700, 800, 900, 3610])
        for c, lbl in zip(fu.rows[0].cells,
                           ['Operator', 'Month', 'Year', 'Errors', 'Error %',
                            'Error Types']):
            _hdr_cell(c, lbl, FILL_BLUE_HDR, C_BLUE)

        for op_name, entries in review_ops:
            for ri, e in enumerate(entries):
                is_cur = (e['year'] == int(report_year)
                          and e['month'] == int(report_month_num))
                fill  = FILL_WHITE if ri % 2 == 0 else FILL_BLUE_HDR
                color = C_RED_TEXT if is_cur else C_BLUE
                codes = e.get('codes') or {}
                codes_txt = ',  '.join(
                    f'{c} ×{n}' if n > 1 else str(c)
                    for c, n in sorted(codes.items())) or '—'
                dr    = fu.add_row().cells
                _data_cell(dr[0], op_name if ri == 0 else '', fill, color, size=9)
                _data_cell(dr[1], _cal.month_name[e['month']],  fill, color, size=9)
                _data_cell(dr[2], str(e['year']),               fill, color, size=9,
                           align=WD_ALIGN_PARAGRAPH.CENTER)
                _data_cell(dr[3], str(e['total_errors']),       fill, color, size=9,
                           align=WD_ALIGN_PARAGRAPH.CENTER)
                _data_cell(dr[4], f"{e['error_rate']:.1f}%",   fill, color, size=9,
                           align=WD_ALIGN_PARAGRAPH.CENTER)
                _data_cell(dr[5], codes_txt,                    fill, color, size=9)

        p = doc.add_paragraph()
        _run(p, 'Error type codes are described in the Error Type Summary '
                'table above.  Current report month shown in red.',
             size=8, color=C_GREY_TEXT)

    doc.add_paragraph()
    p = doc.add_paragraph()
    _run(p,
         'Radiometer ABL Report Generator  ·  Pathology Queensland  ·  All data processed locally',
         size=8, color=C_GREY_TEXT)

    return doc


def safe_fn(name):
    return re.sub(r'[^\w\s\-]', '', str(name)).strip().replace(' ', '_')


def _patch_white_background(docx_path):
    """Ensure Word renders a white page background (displayBackgroundShape)."""
    import zipfile, shutil as _sh
    tmp = docx_path + '.tmp'
    with zipfile.ZipFile(docx_path, 'r') as zin, \
         zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'word/settings.xml':
                xml = data.decode('utf-8')
                if 'displayBackgroundShape' not in xml:
                    xml = xml.replace('</w:settings>',
                                      '<w:displayBackgroundShape/></w:settings>')
                data = xml.encode('utf-8')
            zout.writestr(item, data)
    _sh.move(tmp, docx_path)


# ═══════════════════════════════════════════════════════════════════════════════
#  BACKGROUND WORKER
# ═══════════════════════════════════════════════════════════════════════════════

def run_generation(excel_path, output_dir, report_month, report_month_num, report_year, selected=None):
    global progress
    progress = {"total": 0, "done": 0, "current": "Loading data…",
                "errors": [], "complete": False, "output": output_dir, "reports": 0}
    try:
        df         = load_data(excel_path)
        history    = _load_json(HISTORY_PATH)
        op_history = _load_json(OP_HIST_PATH)

        for req in ('Hospital', 'Analyzer'):
            if req not in df.columns:
                raise ValueError(f"No '{req}' column found — check your spreadsheet format.")

        # ── (hospital, department) pairs — the report unit ──────────────────
        units = sorted(
            {(str(h).strip(), str(d).strip())
             for h, d in zip(df['Hospital'], df['Department'])
             if str(h).strip() and str(d).strip()
             and str(h).strip().lower() not in ('none', 'nan')
             and str(d).strip().lower() not in ('none', 'nan')}
        )

        def _unit_rows(hosp, dept):
            return df[(df['Hospital'].astype(str).str.strip() == hosp) &
                      (df['Department'].astype(str).str.strip() == dept)]

        # Update operator history — only flagged operators are recorded.
        # Department scope: an operator is flagged if any of their analyzer
        # rows meets the rule (> MIN_TESTS tests and > FLAG_PCT % error).
        if 'Operator' in df.columns:
            for (hosp, dept, op), grp in df.groupby(['Hospital', 'Department', 'Operator']):
                flag_rows = grp[grp['Flagged']]
                flagged   = not flag_rows.empty
                rate = float(flag_rows['ErrorRate'].max()) if flagged \
                    else float(grp['ErrorRate'].max())
                codes = {c: int(grp[f'E{c}'].sum()) for c in ERROR_CODES
                         if int(grp[f'E{c}'].sum()) > 0}
                op_history = _update_op_history(op_history, _dept_key(hosp, dept),
                                                str(op).strip(),
                                                report_year, report_month_num,
                                                rate, int(grp['TotalErrors'].sum()),
                                                flagged, codes)

        # Update per-analyzer trend history (errors ÷ 'Total tests on analyzer')
        for (hosp, dept, anlz), grp in df.groupby(['Hospital', 'Department', 'Analyzer']):
            tot = int(grp['TotalTests'].max()) if 'TotalTests' in grp.columns else 0
            if tot <= 0:
                tot = int(grp['Tests'].sum())
            errs  = int(grp['TotalErrors'].sum())
            rate  = round(errs / tot * 100, 2) if tot > 0 else 0.0
            codes = {c: int(grp[f'E{c}'].sum()) for c in ERROR_CODES
                     if int(grp[f'E{c}'].sum()) > 0}
            history = _update_analyzer_history(
                history, _anlz_key(hosp, dept, anlz),
                report_year, report_month_num, rate, tot, codes)

        # selected ids arrive as "Hospital|Department"
        if selected:
            sel = set(selected)
            units = [u for u in units if _dept_key(*u) in sel]

        progress['total'] = len(units)
        os.makedirs(output_dir, exist_ok=True)

        picu_ops = _load_picu_operators()

        for hosp, dept in units:
            progress['current'] = f'{hosp} — {dept}'
            try:
                df_d = _unit_rows(hosp, dept).copy()
                doc = generate_report(hosp, dept, df_d,
                                      report_month, report_month_num,
                                      report_year, history=history, op_history=op_history,
                                      picu_ops=picu_ops)
                if doc:
                    hdir = os.path.join(output_dir, safe_fn(hosp))
                    os.makedirs(hdir, exist_ok=True)
                    fname    = f"ABL_{safe_fn(dept)}_{report_month.replace(' ', '')}.docx"
                    out_path = os.path.join(hdir, fname)
                    doc.save(out_path)
                    _patch_white_background(out_path)
                    progress['reports'] += 1
            except Exception:
                import traceback
                progress['errors'].append(f"{hosp} — {dept}: "
                                          f"{traceback.format_exc(limit=2)}")
            progress['done'] += 1

        _save_json(HISTORY_PATH,  history)
        _save_json(OP_HIST_PATH,  op_history)

    except Exception:
        import traceback
        progress['errors'].append(f"Fatal: {traceback.format_exc(limit=3)}")
    finally:
        progress['complete'] = True
        progress['current']  = 'Done'
    try:
        os.remove(excel_path)
    except Exception:
        pass


# ═══════════════════════════════════════════════════════════════════════════════
#  FLASK ROUTES
# ═══════════════════════════════════════════════════════════════════════════════

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/preview', methods=['POST'])
def preview():
    if not _same_origin(request):
        abort(403)
    if 'file' not in request.files:
        return jsonify({'error': 'No file'}), 400
    f = request.files['file']
    fd, tmp = tempfile.mkstemp(suffix='.xlsx', prefix='abl_preview_')
    os.close(fd)
    try:
        f.save(tmp)
        df = load_data(tmp)
        for req in ('Hospital', 'Analyzer'):
            if req not in df.columns:
                return jsonify({'error': f"No '{req}' column found in the spreadsheet."}), 400
        # One selectable unit per Hospital+Department, listing its analyzers
        units = {}
        for _, r in df.iterrows():
            h, d = str(r['Hospital']).strip(), str(r['Department']).strip()
            a    = str(r['Analyzer']).strip()
            if not h or not d or h.lower() in ('none', 'nan') or d.lower() in ('none', 'nan'):
                continue
            units.setdefault((h, d), set())
            if a and a.lower() not in ('none', 'nan'):
                units[(h, d)].add(a)
        result = [{'id': _dept_key(h, d), 'hhs': h, 'name': d,
                   'analyzers': sorted(alz)}
                  for (h, d), alz in sorted(units.items())]
        return jsonify({'hospitals': result})
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass

@app.route('/generate', methods=['POST'])
def generate():
    global progress
    if not _same_origin(request):
        abort(403)
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    f = request.files['file']
    if not f.filename.endswith('.xlsx'):
        return jsonify({'error': 'Please upload an .xlsx file'}), 400

    month_name   = request.form.get('month_name', 'April 2025')
    month_num    = request.form.get('month_num',  '04')
    year         = request.form.get('year',       '2025')
    sel_json     = request.form.get('selected_hospitals', '')
    selected     = json.loads(sel_json) if sel_json else None

    fd, upload_path = tempfile.mkstemp(suffix='.xlsx', prefix='abl_upload_')
    os.close(fd)
    f.save(upload_path)

    _audit_event('generate', period=month_name,
                 analyzers=(len(selected) if selected else 'all'))

    output_dir = _output_dir_for(month_name, month_num, year)
    os.makedirs(output_dir, exist_ok=True)
    progress   = {"total": 0, "done": 0, "current": "Starting…",
                  "errors": [], "complete": False, "output": output_dir, "reports": 0}
    t = threading.Thread(
        target=run_generation,
        args=(upload_path, output_dir, month_name, month_num, year, selected),
        daemon=True
    )
    t.start()
    return jsonify({'status': 'started', 'output': output_dir})

@app.route('/progress')
def get_progress():
    return jsonify(progress)

@app.route('/open_folder')
def open_folder():
    folder = request.args.get('path', '')
    if not folder:
        return jsonify({'error': 'Folder not found'}), 404
    real = os.path.realpath(folder)
    if os.path.commonpath([real, OUTPUT_ROOT]) != OUTPUT_ROOT:
        abort(403)
    if not os.path.isdir(real):
        return jsonify({'error': 'Folder not found'}), 404
    try:
        if platform.system() == 'Windows':
            os.startfile(real)
        elif platform.system() == 'Darwin':
            subprocess.Popen(['open', real])
        else:
            subprocess.Popen(['xdg-open', real])
        return jsonify({'status': 'opened'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    import webbrowser
    print(f"\n  Radiometer ABL Report Generator  v1")
    print(f"  Running at http://localhost:{PORT}\n")
    if not os.environ.get('PQ_PORTAL'):   # the portal opens the browser itself
        threading.Timer(1.2, lambda: webbrowser.open(f'http://localhost:{PORT}')).start()
    app.run(host='127.0.0.1', port=PORT, debug=False)
