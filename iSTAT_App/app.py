# -*- coding: utf-8 -*-
"""
i-STAT Report Generator — Local Web App  v4
Run via start_windows.bat or start_mac.command

v4 changes:
  - Reports generated from template.docx, preserving logos, header/footer, page numbers
  - Exact colour scheme and table styling from the supplied example report
  - Red text for staff with ≥3 total errors in the error details table
  - Italic "i" in report title, green subtitle, underlined blue section headings

Security hardening (PSPF / ISM / IS18 / APP):
  - Active network guard (netguard.py): outbound connections blocked + logged
  - tempfile.mkstemp() for all uploads — randomised, deleted in finally (no data remanence)
  - /open_folder confined to OUTPUT_ROOT via realpath + commonpath (path traversal)
  - Origin/Referer CSRF check on all POST routes
  - Audit log: timestamp, OS user, action, period, scope
  - Explicit 127.0.0.1 loopback bind (never exposed on network)
"""
import os, re, sys, threading, subprocess, platform, shutil, json, io, tempfile, getpass, logging
from datetime import datetime

# Make this app's own folder importable regardless of how Python was started
# (the embeddable runtime does not add the script directory to sys.path).
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import netguard
netguard.install()   # loopback-only networking from this point on (see netguard.py)

from flask import Flask, render_template, request, jsonify, abort
import openpyxl
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
APP_COMPONENT = 'istat'

# ═══════════════════════════════════════════════════════════════════════════════
#  LICENCE GATE — the bundle runs only under a licence file issued by the PoC
#  lead's offline keygen (pq_licence.py at the bundle root holds the public
#  key). Without the module, or without a valid licence covering this machine,
#  every request is answered with a lock page.
# ═══════════════════════════════════════════════════════════════════════════════
def _load_licence():
    import importlib.util
    _here = os.path.dirname(os.path.abspath(__file__))
    for cand in (os.path.join(_here, '..'), os.path.join(_here, '..', 'PQ_Portable')):
        path = os.path.join(cand, 'pq_licence.py')
        if os.path.exists(path):
            try:
                spec = importlib.util.spec_from_file_location('pq_licence', path)
                mod  = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(mod)
                return mod
            except Exception as exc:
                print(f'[LICENCE] failed to load {path}: {exc}')
                return None
    return None

_LIC = _load_licence()
if _LIC is not None:
    try:
        _LIC.record_launch(APP_COMPONENT)
    except Exception:
        pass

_LOCK_PAGE = """<!DOCTYPE html><html><head><meta charset="utf-8"><title>PC not licensed</title>
<style>body{font-family:-apple-system,Segoe UI,Arial,sans-serif;background:#F2F2F7;color:#17272E;display:flex;
align-items:center;justify-content:center;height:100vh;margin:0}.card{background:#fff;border-radius:16px;
padding:36px 40px;max-width:560px;box-shadow:0 10px 34px rgba(1,38,50,.25)}h1{color:#01485C;font-size:22px;margin:0 0 12px}
p{line-height:1.5;margin:8px 0}code{background:#E5F2F0;padding:2px 6px;border-radius:6px}a{color:#01485C}</style></head>
<body><div class="card"><h1>This PC is not licensed</h1><p>{msg}</p>
<p>Machine ID <code>{mid}</code> &middot; {host}</p>
<p>Open the selection page at <a href="http://localhost:5750/">localhost:5750</a> for the licence status and to
install a licence file from the PoC lead. Nothing can be generated until then.</p></div></body></html>"""

@app.before_request
def _enforce_licence():
    if _LIC is None:
        return ('<h2 style="font-family:sans-serif;color:#B4540A">Licence module missing &mdash; '
                'this copy of the bundle is incomplete and cannot run.</h2>'), 403
    st = _LIC.status()
    if not st['allowed']:
        body = (_LOCK_PAGE.replace('{msg}', st['message']).replace('{mid}', st['machine_id'])
                .replace('{host}', st['hostname']))
        return body, 403, {'Content-Type': 'text/html; charset=utf-8'}


progress = {"total": 0, "done": 0, "current": "", "errors": [], "complete": False, "output": "", "reports": 0}

# ── Exact colours from the template ────────────────────────────────────────────
C_DARK_BLUE  = RGBColor(0x1F, 0x49, 0x7D)   # title text
C_GREEN      = RGBColor(0x9B, 0xBB, 0x59)   # subtitle
C_BLUE       = RGBColor(0x30, 0x54, 0x96)   # section headings / table data text
C_WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
C_RED_TEXT   = RGBColor(0xFF, 0x00, 0x00)   # high-error staff highlight
C_GREY_TEXT  = RGBColor(0x60, 0x60, 0x60)
C_GREEN_PASS = RGBColor(0x00, 0x70, 0x00)
C_RED_FAIL   = RGBColor(0x80, 0x00, 0x00)

FILL_GREY_HDR   = 'AEAAAA'   # performance table header (label cols)
FILL_RED_HDR    = 'FF0000'   # performance table header (metric cols)
FILL_BLUE_HDR   = 'D9E1F2'   # error table header / alternating rows (light blue)
FILL_ALT_ROW    = 'D9D9D9'   # alternating body rows (light grey)
FILL_WHITE      = 'FFFFFF'

# ── Template path ───────────────────────────────────────────────────────────────
APP_DIR       = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH  = os.path.join(APP_DIR, 'template.docx')
HISTORY_PATH         = os.path.join(APP_DIR, 'error_rate_history.json')
STAFF_HISTORY_PATH      = os.path.join(APP_DIR, 'staff_error_history.json')
ERROR_TYPE_HISTORY_PATH = os.path.join(APP_DIR, 'error_type_history.json')
ISTAT_CODES_PATH        = os.path.join(APP_DIR, 'istat_error_codes.json')

# Reports are only ever written under the parent of the app folder. /open_folder
# is restricted to this root so a crafted path cannot open arbitrary locations.
REPORTS_ROOT = os.path.realpath(os.path.join(APP_DIR, '..', 'Reports'))   # shared by both apps
OUTPUT_ROOT  = os.path.join(REPORTS_ROOT, 'iSTAT')

def _output_dir_for(month_name, month_num, year):
    """Reports\\iSTAT\\<YYYY-MM Month> under the bundle root — sorts by date in
    Explorer. Inputs are sanitised so a crafted period can't escape OUTPUT_ROOT."""
    try:
        mm = int(re.sub(r'[^0-9]', '', str(month_num)) or 0)
        yy = int(re.sub(r'[^0-9]', '', str(year)) or 0)
    except ValueError:
        mm, yy = 0, 0
    mon = re.sub(r'[^A-Za-z]', '', str(month_name).split(' ')[0])[:12] or 'Month'
    label = f'{yy:04d}-{mm:02d} {mon}'
    return os.path.normpath(os.path.join(OUTPUT_ROOT, label))

# ── Audit log (who generated what, when) — required for PSPF/ISM event logging ──
AUDIT_LOG_PATH = os.path.join(APP_DIR, 'audit.log')
_audit = logging.getLogger('istat.audit')
_audit.setLevel(logging.INFO)
if not _audit.handlers:
    _h = logging.FileHandler(AUDIT_LOG_PATH, encoding='utf-8')
    _h.setFormatter(logging.Formatter('%(asctime)s\t%(message)s'))
    _audit.addHandler(_h)

def _audit_event(action, **fields):
    """Append a tab-delimited audit record: timestamp, OS user, action, details."""
    try:
        user = getpass.getuser()
    except Exception:
        user = 'unknown'
    detail = '\t'.join(f'{k}={v}' for k, v in fields.items())
    _audit.info('user=%s\taction=%s\t%s', user, action, detail)

def _same_origin(req):
    """Reject cross-origin POSTs (CSRF defence) — only allow our own localhost UI."""
    origin = req.headers.get('Origin') or req.headers.get('Referer') or ''
    if not origin:
        return True  # non-browser / same-process callers send no Origin
    return origin.startswith(f'http://localhost:{PORT}') or \
           origin.startswith(f'http://127.0.0.1:{PORT}')

PORT = 5757

# ═══════════════════════════════════════════════════════════════════════════════
#  DOCX HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def _shd(cell, fill):
    """Set cell background shading (fill = 6-char hex string)."""
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    old  = tcPr.find(qn('w:shd'))
    if old is not None:
        tcPr.remove(old)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  fill)
    tcPr.append(shd)

def _cell_valign(cell, val='center'):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    vA   = OxmlElement('w:vAlign')
    vA.set(qn('w:val'), val)
    tcPr.append(vA)

def _cell_width(cell, dxa):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    w    = OxmlElement('w:tcW')
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
    tbl  = table._tbl
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

def _cell_padding(table, top=90, bottom=90, left=110, right=110):
    """Breathing room inside every cell of a table (values in twips)."""
    tblPr = table._tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        table._tbl.insert(0, tblPr)
    old = tblPr.find(qn('w:tblCellMar'))
    if old is not None:
        tblPr.remove(old)
    mar = OxmlElement('w:tblCellMar')
    for side, val in (('top', top), ('left', left), ('bottom', bottom), ('right', right)):
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:w'), str(val))
        el.set(qn('w:type'), 'dxa')
        mar.append(el)
    tblPr.append(mar)

def _set_col_widths(table, widths_dxa):
    """Set column and cell widths from a list of DXA values."""
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
    # Apply to each cell in every row
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
    """Add a blue underlined section heading matching the template Heading3 style."""
    p = doc.add_paragraph(style='Heading3')
    # Clear any default text
    for r in p.runs:
        r.text = ''
    r = p.add_run(text)
    r.font.size      = Pt(12)
    r.font.color.rgb = C_BLUE
    r.font.underline = True
    r.bold           = True
    return p

def _hdr_cell(cell, text, fill, text_color=None):
    """Style a header cell: fill + white centred bold text."""
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
    r.font.size      = Pt(size)
    r.bold           = bold
    if color:
        r.font.color.rgb = color
    return r

def get_perf_label(rate):
    if rate < 4:   return 'Champions',      RGBColor(0x53, 0x81, 0x35)
    if rate <= 6:  return 'Excellent',       RGBColor(0x2F, 0x54, 0x96)
    if rate <= 10: return 'Acceptable',      RGBColor(0xFF, 0x80, 0x00)
    return             'Needs attention',    RGBColor(0xFF, 0x00, 0x00)


# ═══════════════════════════════════════════════════════════════════════════════
#  ROBUST EXCEL LOADING
# ═══════════════════════════════════════════════════════════════════════════════

_USE_KEYS = ['cart', 'cartridge']
_SIM_KEYS = ['sim', 'simulator']
_ERR_KEYS     = ['event', 'operator', 'userevents']
_PERUSER_KEYS = ['permonthperuser', 'per month per user', 'errorpermonth', 'peruser']

_COL_ALIASES = {
    'Hospital Name': ['hospital name', 'hospital', 'site', 'facility'],
    'HHS':           ['hhs'],
    'Department':    ['department', 'dept'],
    'Location':      ['location', 'ward', 'area'],
    'Device Name':   ['device name', 'device', 'instrument'],
    'DeviceID':      ['deviceid', 'device id', 'serial'],
    'Year':          ['year'],
    'Month':         ['month'],
    'Total Carts':   ['total carts', 'total cartridges', 'cartridges'],
    'Total Res':     ['total res', 'total results', 'results'],
    'Result':        ['result', 'results'],
    'Surname':       ['surname', 'last name', 'family name'],
    'First Name':    ['first name', 'given name', 'forename'],
    'User ID':       ['user id', 'userid', 'staff id'],
    'Sub Code':      ['sub code', 'subcode'],
    'Error Code':    ['error code', 'code'],
    'TEXT':          ['text', 'description', 'error description'],
    'PATS':          ['pats'],
}

def _find_sheet(wb, kws, ex=None):
    ex = ex or []
    for name in wb.sheetnames:
        low = name.lower()
        if any(e in low for e in ex):
            continue
        if any(k in low for k in kws):
            return wb[name]
    return wb[wb.sheetnames[0]]

def _norm_hdrs(raw):
    result = []
    for i, h in enumerate(raw):
        s = str(h).strip() if h else ''
        m = s if s else f'_blank{i}'
        sl = s.lower()
        # Exact alias match first — otherwise 'Device ID' falls into
        # 'Device Name' because 'device' substring-matches earlier.
        hit = None
        for canon, aliases in _COL_ALIASES.items():
            if sl in aliases:
                hit = canon
                break
        if hit is None:
            for canon, aliases in _COL_ALIASES.items():
                if any(a in sl for a in aliases):
                    hit = canon
                    break
        result.append(hit or m)
    return result

def _find_hdr_row(ws, max_scan=15):
    known = set(a for al in _COL_ALIASES.values() for a in al)
    best, bs = 0, -1
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True)):
        sc = sum(1 for c in row if c and str(c).strip().lower() in known)
        if sc > bs:
            bs, best = sc, i + 1
    return best

def _load_df(ws):
    hr   = _find_hdr_row(ws)
    rows = list(ws.iter_rows(min_row=hr, values_only=True))
    if not rows:
        return pd.DataFrame()
    nh = _norm_hdrs(rows[0])
    return pd.DataFrame([dict(zip(nh, r)) for r in rows[1:] if any(v is not None for v in r)])

def _extract_did(device_name):
    s = str(device_name or '')
    if ';' in s:
        tail = s.rsplit(';', 1)[-1].strip()
        if tail.isdigit():
            return tail
    return None

def _norm_uid(v):
    """Normalise a User ID for cross-sheet matching ('000043' == '43')."""
    s = str(v or '').strip()
    return s.lstrip('0') or s

# ── Device rules ──────────────────────────────────────────────────────────────
# Serial numbers NEVER to appear in any generated report, even when the
# supplied export contains data for them (retired / spare instruments).
EXCLUDED_DEVICE_IDS = {'456190', '308066', '308068', '306666',
                       '453165', '319576', '341888', '345220',
                       # retired 2026-09:
                       '336713', '359416', '386056', '343167',
                       # older models retired 2026-09:
                       '306665', '318248', '323289', '336715', '434958', '434963',
                       # Ayr 2026-09: 357192 retired (456192 reinstated)
                       '357192',
                       # Atherton 2026-09: phantom unit, 0 cartridges in 12 months
                       '341361'}

# Every i-STAT must have this many ceramic cleaning cartridges per month.
CERAMIC_REQUIRED = 3

# ── Per-device report titles ─────────────────────────────────────────────────
# A per-device report is titled (and its file named) after the WARD the
# analyser sits in, taken from the export's Location column — e.g. serial
# 307545 at "Townsville O/T" becomes "Townsville OT – August 2026" and
# "i-STAT_Townsville_OT_307545_August2026.docx".
# Put a serial here to force exact wording instead:
DEVICE_TITLE_OVERRIDES = {
    # '307545': 'Townsville OT',
}
# Ward labels to reword (applied to the finished label, so the wording
# survives an analyser being swapped for a new serial).
WARD_TITLE_RENAMES = {
    'Townsville MAT': 'Townsville Maternity',
    'Townsville MET CT2': 'Townsville MET CT 2',
}
# Site abbreviations the middleware uses at the start of a Location.
LOCATION_PREFIX_EXPANSIONS = {'TN ': 'Townsville ', 'Towns ': 'Townsville ',
                              'RD ': 'Redcliffe '}

def _device_title_label(short, device, location):
    """Ward-based title for one analyser's report ('Townsville OT')."""
    did = _extract_did(str(device)) or ''
    if did in DEVICE_TITLE_OVERRIDES:
        return DEVICE_TITLE_OVERRIDES[did]
    loc = str(location or '').strip()
    if not loc or loc.lower() == 'nan':
        return short                                   # no ward known
    loc = re.sub(r'(?<=\w)/(?=\w)', '', loc)             # 'O/T' -> 'OT'
    loc = re.sub(r'\s+', ' ', loc)
    for abbr, full in LOCATION_PREFIX_EXPANSIONS.items():
        if loc.startswith(abbr):
            loc = full + loc[len(abbr):]
            break
    if short.lower() not in loc.lower():               # 'Chest Clinic' -> 'Townsville Chest Clinic'
        loc = f'{short} {loc}'
    return WARD_TITLE_RENAMES.get(loc, loc)

# Reports for these hospitals (matched on the site name after the HHS
# prefix, case-insensitive) are filed under a different HHS folder.
OUTPUT_FOLDER_OVERRIDES = {'bowen': 'Townsv'}   # Bowen sits under Townsville

# Hospitals that get ONE REPORT PER DEVICE (substring match on the hospital
# name, case-insensitive). Mackay is special-cased in run_generation: only the
# main Mackay hospital splits; Proserpine, Bowen etc. stay combined.
PER_DEVICE_HOSPITALS = (
    'townsville', 'mt isa', 'mount isa', 'redcliffe',
    # group one, added 2026-09-27 (full lower-case hospital names = exact match):
    'meno_royal brisbane wh', 'ca_cairns', 'dado_toowoomba', 'wemo_ipswich',
    'childqld_lcch', 'meno_caboolture', 'meno_prince charles', 'goco_gold coast',
    'meso_logan', 'suco_scuh', 'meso_princess alexandra', 'cqld_rockhampton',
    'wiba_bundaberg', 'meso_redland', 'wiba_hervey bay', 'suco_gympie',
    'wiba_maryborough', 'toca_thursday island', 'cqld_emerald',
    'goco_robina',   # 5 wards (ED, ED 2, RASS, Radiology, Lab) — added 2026-09-27
)

# Devices whose data must generate under a different hospital, e.g.
#   '123456': {'Hospital Name': 'HHS_Site', 'Department': 'Site', 'Location': 'Site'}
# (343167 was relocated to Magnetic Island until it was retired in 2026-09.)
DEVICE_RELOCATIONS = {}

def _norm_did(v):
    """Normalise a device serial ('343167.0' == '343167')."""
    s = str(v or '').strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

def _safe_int(v):
    """int() that treats None, NaN, blanks and junk text as 0
    ('x or 0' does NOT catch NaN — NaN is truthy)."""
    try:
        if v is None or (isinstance(v, float) and v != v):
            return 0
        return int(float(v))
    except (ValueError, TypeError):
        return 0

def _apply_device_rules(df):
    """Drop EXCLUDED_DEVICE_IDS rows and re-home DEVICE_RELOCATIONS rows.
    Works on any sheet: uses the DeviceID column when present, otherwise
    the ';serial' tail of Device Name."""
    if df is None or df.empty:
        return df
    id_cols = [c for c in df.columns if str(c).strip().lower().replace(' ', '') == 'deviceid']
    def did_of(row):
        for c in id_cols:
            s = _norm_did(row.get(c))
            if s and s.lower() != 'nan':
                return s
        if 'Device Name' in df.columns:
            return _extract_did(row.get('Device Name'))
        return None
    ids = df.apply(did_of, axis=1)
    df = df[~ids.isin(EXCLUDED_DEVICE_IDS)].copy()
    ids = ids.loc[df.index]
    for did, fields in DEVICE_RELOCATIONS.items():
        mask = ids == did
        if mask.any():
            for col, val in fields.items():
                if col in df.columns:   # only rewrite columns the sheet has
                    df.loc[mask, col] = val
    return df

def load_data(excel_path):
    wb = openpyxl.load_workbook(excel_path, data_only=True)

    # ── Usage sheet ────────────────────────────────────────────────
    ws_use  = _find_sheet(wb, _USE_KEYS, ['sim', 'error', 'event', 'user'])
    df_use  = _load_df(ws_use)
    if 'Hospital Name' not in df_use.columns and 'Hospital' in df_use.columns:
        df_use['Hospital Name'] = df_use['Hospital']
    df_use = df_use[df_use.get('Hospital Name', pd.Series(dtype=str)).notna()].copy()
    df_use = df_use[~df_use['Hospital Name'].astype(str).str.upper().isin(['ALL', 'TOTAL', ''])]
    # Footer/metadata lines (e.g. 'Report Version 210506a') carry text in the
    # hospital column but no device — they are not hospitals.
    if 'Device Name' in df_use.columns:
        df_use = df_use[df_use['Device Name'].notna()]
    df_use = _apply_device_rules(df_use)

    def qcc(row):
        try:
            t = float(row.get('Total Carts') or 0)
            r = float(row.get('Total Res')   or 0)
            return round((t - r) / t * 100, 1) if t else 0.0
        except:
            return 0.0
    df_use['ErrorRate'] = df_use.apply(qcc, axis=1)

    # ── SIM sheet ──────────────────────────────────────────────────
    ws_sim = _find_sheet(wb, _SIM_KEYS, ['error', 'event', 'cart', 'user'])
    df_sim = _load_df(ws_sim)
    df_sim = _apply_device_rules(df_sim)
    df_sim['_r'] = df_sim.get('Result', pd.Series(dtype=str)).astype(str).str.strip()

    hosp_col = next((c for c in df_sim.columns
                     if c.lower() in ('hospital name', 'hospital')
                     and not c.startswith('_blank')), None)
    if hosp_col and hosp_col != 'Hospital Name':
        df_sim['Hospital Name'] = df_sim[hosp_col]
    elif 'Hospital Name' not in df_sim.columns:
        id2h = {}
        for _, r in df_use.iterrows():
            did = _extract_did(r.get('Device Name', ''))
            if did and r.get('Hospital Name'):
                id2h[did] = str(r['Hospital Name'])
        if 'DeviceID' in df_sim.columns:
            df_sim['Hospital Name'] = df_sim['DeviceID'].astype(str).str.strip().map(id2h)
        elif 'Device Name' in df_sim.columns:
            df_sim['Hospital Name'] = df_sim['Device Name'].apply(_extract_did).map(id2h)

    dev_col = 'Device Name' if 'Device Name' in df_sim.columns else None

    sim_pass = df_sim[df_sim['_r'] == '80']
    if 'Hospital Name' in sim_pass.columns and dev_col:
        sim_counts = (sim_pass.dropna(subset=['Hospital Name'])
                      .groupby(['Hospital Name', dev_col]).size()
                      .reset_index(name='SIM_Runs')
                      .rename(columns={dev_col: 'Device Name'}))
    else:
        sim_counts = pd.DataFrame(columns=['Hospital Name', 'Device Name', 'SIM_Runs'])

    ceramic = df_sim[df_sim['_r'].str.upper() == 'FAIL']
    if 'Hospital Name' in ceramic.columns:
        ceramic_counts = (ceramic.dropna(subset=['Hospital Name'])
                          .groupby('Hospital Name').size()
                          .reset_index(name='Ceramic_Count'))
        # Per-device ceramic counts (for per-device reports)
        dev_grp_cols = ['Hospital Name'] + ([dev_col] if dev_col and dev_col in ceramic.columns else [])
        ceramic_by_device = (ceramic.dropna(subset=['Hospital Name'])
                             .groupby(dev_grp_cols).size()
                             .reset_index(name='Ceramic_Count'))
        if dev_col and dev_col in ceramic_by_device.columns and dev_col != 'Device Name':
            ceramic_by_device = ceramic_by_device.rename(columns={dev_col: 'Device Name'})
    else:
        ceramic_counts    = pd.DataFrame(columns=['Hospital Name', 'Ceramic_Count'])
        ceramic_by_device = pd.DataFrame(columns=['Hospital Name', 'Device Name', 'Ceramic_Count'])

    # ── Events / Error sheet ───────────────────────────────────────
    ws_err = _find_sheet(wb, _ERR_KEYS, ['sim', 'cart', 'istat error'])
    df_err = _load_df(ws_err)
    if 'Hospital Name' not in df_err.columns and 'Hospital' in df_err.columns:
        df_err['Hospital Name'] = df_err['Hospital']
    if 'Hospital Name' in df_err.columns:
        df_err = df_err[df_err['Hospital Name'].notna()]
    df_err = _apply_device_rules(df_err)

    # ── Per-user summary sheet (istat errorpermonthperuser) ────────────────
    # Contains PATS column = total cartridges run by each user that month.
    # Load every sheet whose name contains any _PERUSER_KEYS keyword.
    pats_by_staff = {}   # (hospital_upper, surname_upper, firstname_upper) -> pats_int
    print(f'[PATS] All sheets: {wb.sheetnames}')
    for sname in wb.sheetnames:
        slow = sname.lower().replace(' ', '')
        matched = any(k.replace(' ', '') in slow for k in _PERUSER_KEYS)
        print(f'[PATS] Sheet "{sname}" -> slow="{slow}" matched={matched}')
        if matched:
            df_pu = _load_df(wb[sname])
            print(f'[PATS] Loaded sheet "{sname}": {len(df_pu)} rows, columns={list(df_pu.columns)}')
            if df_pu.empty:
                continue
            if 'Hospital Name' not in df_pu.columns and 'Hospital' in df_pu.columns:
                df_pu['Hospital Name'] = df_pu['Hospital']
            # Flexible column detection for surname/first name variants
            _pu_sn  = next((c for c in df_pu.columns
                            if any(k in c.lower() for k in ('surname', 'last name', 'lastname', 'last_name'))), None)
            _pu_fn  = next((c for c in df_pu.columns
                            if any(k in c.lower() for k in ('first name', 'firstname', 'first_name', 'given'))), None)
            # Also try just 'first' if the above didn't match
            if not _pu_fn:
                _pu_fn = next((c for c in df_pu.columns if 'first' in c.lower()), None)
            if not _pu_sn:
                _pu_sn = next((c for c in df_pu.columns if 'surname' in c.lower()), None)
            _pu_pat = next((c for c in df_pu.columns
                            if c.strip().upper() == 'PATS' or 'pats' in c.lower()), None)
            _pu_qc  = next((c for c in df_pu.columns
                            if c.strip().upper() == 'QC'), None)
            _pu_uid = 'User ID' if 'User ID' in df_pu.columns else None
            _pu_hn  = 'Hospital Name' if 'Hospital Name' in df_pu.columns else None
            print(f'[PATS] Detected cols: sn={_pu_sn!r} fn={_pu_fn!r} pats={_pu_pat!r} qc={_pu_qc!r} uid={_pu_uid!r} hosp={_pu_hn!r}')
            if _pu_sn and _pu_fn and _pu_pat:
                def _pu_num(v):
                    try:
                        if pd.notna(v) and str(v).strip() not in ('', 'nan'):
                            return int(float(v))
                    except Exception:
                        pass
                    return 0
                for _, row in df_pu.iterrows():
                    try:
                        hosp_k = str(row[_pu_hn]).strip().upper() if _pu_hn else ''
                        sn_k   = str(row[_pu_sn]).strip().upper()
                        fn_k   = str(row[_pu_fn]).strip().upper()
                        # Pats and QC count SUCCESSFUL cartridges only; failed
                        # cartridges appear solely as error events. Sum across
                        # a person's rows (one per device model) — never max.
                        base = _pu_num(row[_pu_pat]) + (_pu_num(row[_pu_qc]) if _pu_qc else 0)
                        keys = [(hosp_k, sn_k, fn_k)]
                        if _pu_uid and pd.notna(row[_pu_uid]):
                            keys.append(('UID', _norm_uid(row[_pu_uid])))
                        for key in keys:
                            pats_by_staff[key] = pats_by_staff.get(key, 0) + base
                    except Exception:
                        pass
                print(f'[PATS] Built {len(pats_by_staff)} entries: {dict(list(pats_by_staff.items())[:5])}')
            else:
                print(f'[PATS] SKIPPED -- could not find required columns (sn/fn/pats)')

    return df_use, sim_counts, ceramic_counts, ceramic_by_device, df_err, pats_by_staff



# ═══════════════════════════════════════════════════════════════════════════════
#  HISTORY ENCRYPTION (AES via Fernet; key derived from password, never stored)
# ═══════════════════════════════════════════════════════════════════════════════
import base64 as _b64
import glob as _hist_glob

KEYMETA_PATH = os.path.join(APP_DIR, 'history.keymeta')
_ENC_MAGIC   = b'ISTATENC1'
_hist_fernet = None            # set only after a successful unlock

class _HistoryLocked(RuntimeError):
    pass

def _derive_key(password, salt):
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32,
                     salt=salt, iterations=600_000)
    return _b64.urlsafe_b64encode(kdf.derive(password.encode('utf-8')))

def _encryption_initialized():
    return os.path.exists(KEYMETA_PATH)

def _unlocked():
    return _hist_fernet is not None

def _encrypt_existing_plaintext():
    """One-time sweep: encrypt any plaintext history files (incl. .bak copies)."""
    for p in _hist_glob.glob(os.path.join(APP_DIR, '*history*')):
        if not os.path.isfile(p) or os.path.abspath(p) == os.path.abspath(KEYMETA_PATH):
            continue
        try:
            with open(p, 'rb') as fh:
                raw = fh.read()
            if raw.startswith(_ENC_MAGIC):
                continue
            json.loads(raw.decode('utf-8'))       # only touch valid JSON files
            with open(p, 'wb') as fh:
                fh.write(_ENC_MAGIC + _hist_fernet.encrypt(raw))
        except Exception:
            continue                              # logs etc. are left alone

def _set_history_password(password):
    """First-time setup: derive key, store salt+verifier, encrypt existing files."""
    global _hist_fernet
    from cryptography.fernet import Fernet
    salt = os.urandom(16)
    f = Fernet(_derive_key(password, salt))
    meta = {'salt': _b64.b64encode(salt).decode(),
            'verifier': f.encrypt(b'ISTAT-HISTORY-OK').decode()}
    with open(KEYMETA_PATH, 'w') as fh:
        json.dump(meta, fh)
    _hist_fernet = f
    _encrypt_existing_plaintext()

def _try_unlock(password):
    global _hist_fernet
    from cryptography.fernet import Fernet
    try:
        with open(KEYMETA_PATH) as fh:
            meta = json.load(fh)
        salt = _b64.b64decode(meta['salt'])
        f = Fernet(_derive_key(password, salt))
        if f.decrypt(meta['verifier'].encode()) == b'ISTAT-HISTORY-OK':
            _hist_fernet = f
            _encrypt_existing_plaintext()         # catch any strays
            return True
    except Exception:
        pass
    return False

def _read_history_file(path):
    """Parse a history file, transparently decrypting when encrypted."""
    with open(path, 'rb') as fh:
        raw = fh.read()
    if raw.startswith(_ENC_MAGIC):
        if not _unlocked():
            raise _HistoryLocked('History files are encrypted — unlock first.')
        raw = _hist_fernet.decrypt(raw[len(_ENC_MAGIC):])
    return json.loads(raw.decode('utf-8'))

def _write_history_file(path, obj):
    data = json.dumps(obj, indent=2).encode('utf-8')
    if _unlocked():
        data = _ENC_MAGIC + _hist_fernet.encrypt(data)
    with open(path, 'wb') as fh:
        fh.write(data)


# ═══════════════════════════════════════════════════════════════════════════════
#  HISTORY + TREND CHART
# ═══════════════════════════════════════════════════════════════════════════════

def _load_history():
    if os.path.exists(HISTORY_PATH):
        try:
            return _read_history_file(HISTORY_PATH)
        except _HistoryLocked:
            raise
        except Exception:
            pass
    return {}

def _save_history(history):
    try:
        _write_history_file(HISTORY_PATH, history)
    except Exception:
        pass

def _history_key(hospital, device):
    return f"{hospital}|{device}"

def _update_history(history, hospital, device, year, month_num, error_rate,
                    carts=0, types=None):
    key = _history_key(hospital, device)
    entries = [e for e in history.get(key, [])
               if not (e['year'] == int(year) and e['month'] == int(month_num))]
    entry = {'year': int(year), 'month': int(month_num),
             'error_rate': round(float(error_rate), 2)}
    if carts:
        entry['carts'] = int(carts)                       # cartridges run
        entry['types'] = {c: int(n) for c, n in (types or {}).items() if n}
    entries.append(entry)
    entries.sort(key=lambda e: (e['year'], e['month']))
    history[key] = entries[-24:]   # keep rolling 24 months
    return history



def _load_error_type_history():
    if os.path.exists(ERROR_TYPE_HISTORY_PATH):
        try:
            return _read_history_file(ERROR_TYPE_HISTORY_PATH)
        except _HistoryLocked:
            raise
        except Exception:
            pass
    return {}

def _save_error_type_history(eth):
    try:
        _write_history_file(ERROR_TYPE_HISTORY_PATH, eth)
    except Exception:
        pass

def _update_error_type_history(eth, key, year, month_num, counts):
    """Persist monthly error-type counts for a hospital/device key."""
    top_error = max(counts, key=counts.get) if counts else None
    entries   = [e for e in eth.get(key, [])
                 if not (e['year'] == int(year) and e['month'] == int(month_num))]
    entries.append({
        'year':      int(year),
        'month':     int(month_num),
        'top_error': top_error,
        'counts':    counts,
    })
    entries.sort(key=lambda e: (e['year'], e['month']))
    eth[key] = entries[-24:]
    return eth



# ═══════════════════════════════════════════════════════════════════════════════
#  DORMANT ANALYSERS
#  An analyser with no cartridges, no simulator runs AND no ceramic runs for
#  DORMANT_MONTHS consecutive months (ending in the report month) is dormant:
#  it gets no report and is dropped from its site's simulator / ceramic checks,
#  but is listed in Dormant_iSTATs_<Month>.docx so the list can be checked.
#  Activity for every analyser is recorded each run in
#  device_activity_history.json (encrypted like the other history files).
#  A month that was never recorded ends the idle run, so a unit is never
#  removed on incomplete evidence. A single quiet month is normal and never
#  removes anything — the reports exist to catch wards skipping simulators.
# ═══════════════════════════════════════════════════════════════════════════════
DORMANT_MONTHS = 4
ACTIVITY_HISTORY_PATH = os.path.join(APP_DIR, 'device_activity_history.json')
_MON_ABBR = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
             'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

def _load_activity_history():
    if os.path.exists(ACTIVITY_HISTORY_PATH):
        try:
            return _read_history_file(ACTIVITY_HISTORY_PATH)
        except _HistoryLocked:
            raise
        except Exception:
            pass
    return {}

def _save_activity_history(ah):
    try:
        _write_history_file(ACTIVITY_HISTORY_PATH, ah)
    except Exception:
        pass

def _month_key(year, month_num):
    return f'{int(year):04d}-{int(month_num):02d}'

def _prev_month_key(mk):
    y, m = int(mk[:4]), int(mk[5:7])
    return _month_key(y - 1, 12) if m == 1 else _month_key(y, m - 1)

def _month_label(mk):
    return f'{_MON_ABBR[int(mk[5:7])]} {mk[:4]}' if mk else ''

def _device_serial(name):
    s = str(name or '').strip()
    return _extract_did(s) or s

def _activity_key(hospital, device_name):
    return f'{str(hospital).strip()}|{_device_serial(device_name)}'

def _record_activity(ah, df_use, sim_counts, ceramic_by_device, year, month_num):
    """Store [cartridges, simulator runs, ceramic runs] for every analyser in
    the export under 'hospital|serial' for this month (rolling 36 months)."""
    mk = _month_key(year, month_num)

    def _sum(df, col):
        out = {}
        if (df is not None and not df.empty and 'Device Name' in df.columns
                and 'Hospital Name' in df.columns and col in df.columns):
            for _, r in df.iterrows():
                k = _activity_key(r['Hospital Name'], r['Device Name'])
                out[k] = out.get(k, 0) + _safe_int(r[col])
        return out

    sims = _sum(sim_counts, 'SIM_Runs')
    cers = _sum(ceramic_by_device, 'Ceramic_Count')
    if df_use is None or df_use.empty:
        return
    for _, r in df_use.iterrows():
        h = str(r.get('Hospital Name', '') or '').strip()
        d = str(r.get('Device Name', '') or '').strip()
        if not h or not d or h.upper() in ('ALL', 'TOTAL'):
            continue
        k = _activity_key(h, d)
        ent = ah.setdefault(k, {})
        ent['name'] = d
        ent['loc']  = str(r.get('Location', '') or '').strip()
        months = ent.setdefault('months', {})
        months[mk] = [_safe_int(r.get('Total Carts')), sims.get(k, 0), cers.get(k, 0)]
        for old in sorted(months)[:-36]:
            del months[old]

def _find_dormant(ah, year, month_num):
    """Analysers present in the report month whose last DORMANT_MONTHS
    recorded months (report month included) all show zero activity."""
    mk0 = _month_key(year, month_num)
    out = []
    for k, ent in ah.items():
        months = ent.get('months', {})
        if mk0 not in months:
            continue
        hosp, _, serial = k.partition('|')
        idle, mk, last_active = 0, mk0, None
        while mk in months:
            if any(months[mk]):
                last_active = mk
                break
            idle += 1
            mk = _prev_month_key(mk)
        if idle < DORMANT_MONTHS:
            continue
        out.append({'hospital': hosp, 'serial': serial,
                    'name': ent.get('name', ''), 'loc': ent.get('loc', ''),
                    'idle': idle, 'last_active': last_active,
                    'first_seen': min(months)})
    out.sort(key=lambda d: (d['hospital'], d['loc'], d['serial']))
    return out

def _drop_dormant(df, dormant_keys):
    if (df is None or df.empty or not dormant_keys
            or 'Device Name' not in df.columns or 'Hospital Name' not in df.columns):
        return df
    keys = df.apply(lambda r: _activity_key(r['Hospital Name'], r['Device Name']), axis=1)
    return df[~keys.isin(dormant_keys)].copy()

def _write_dormant_report(dormant, out_path, report_month):
    """Companion document listing the analysers left out of this run."""
    doc  = Document(TEMPLATE_PATH)
    body = doc.element.body
    sect = body.find(qn('w:sectPr'))
    for child in list(body):
        if child != sect:
            body.remove(child)
    LOGO_INDENT = Inches(2127 / 1440)
    p = doc.add_paragraph(style='Title')
    p.paragraph_format.left_indent = LOGO_INDENT
    _run(p, 'Pathology Queensland', size=12, color=C_DARK_BLUE)
    p = doc.add_paragraph(style='Title')
    p.paragraph_format.left_indent = LOGO_INDENT
    _run(p, 'i', italic=True, size=20, color=C_DARK_BLUE)
    _run(p, '-Stat Dormant Analysers', size=20, color=C_DARK_BLUE)
    p = doc.add_paragraph(style='Subtitle')
    p.paragraph_format.first_line_indent = LOGO_INDENT
    r = _run(p, f'Removed from the {report_month} run', size=18, color=C_GREEN)
    try:
        r.style = doc.styles['DocSubTitle']
    except Exception:
        pass

    _heading(doc, 'Why these analysers have no report')
    p = doc.add_paragraph()
    _run(p, (f"The {len(dormant)} analyser{'s' if len(dormant) != 1 else ''} listed below recorded "
             f"no cartridges, no simulator runs and no ceramic cleaning cartridges for at least "
             f"{DORMANT_MONTHS} consecutive months up to and including {report_month}. "
             f"They have been left out of this month's reports and do not count against their "
             f"site's simulator or ceramic checks. An analyser that is used again is picked up "
             f"automatically the following month. Please review the list for units that should "
             f"be retired from the device list, and for units that should be in service but are "
             f"not being maintained."), size=10)

    _heading(doc, f'Dormant analysers — {report_month}')
    cols = ['Hospital', 'Location', 'Analyser', 'Serial', 'Idle months', 'Last activity']
    tbl = doc.add_table(rows=1, cols=len(cols))
    _add_borders(tbl)
    for c, lbl in zip(tbl.rows[0].cells, cols):
        _hdr_cell(c, lbl, FILL_BLUE_HDR, C_BLUE)
    for i, d in enumerate(dormant):
        fill = FILL_ALT_ROW if i % 2 else FILL_WHITE
        if d['last_active']:
            last = _month_label(d['last_active'])
        else:
            last = f"none since {_month_label(d['first_seen'])}"
        vals = [d['hospital'], d['loc'], d['name'], d['serial'], str(d['idle']), last]
        cells = tbl.add_row().cells
        for j, (c, v) in enumerate(zip(cells, vals)):
            _data_cell(c, v, fill=fill, color=C_BLUE,
                       align=WD_ALIGN_PARAGRAPH.CENTER if j in (3, 4) else WD_ALIGN_PARAGRAPH.LEFT,
                       size=8)
    _set_col_widths(tbl, [2000, 2000, 2500, 1000, 900, 1960])
    _cell_padding(tbl, top=40, bottom=40)
    doc.save(out_path)
    _patch_white_background(out_path)


def _months_consecutive(e1, e2):
    """True if e2 is the calendar month immediately after e1."""
    y1, m1 = e1['year'], e1['month']
    y2, m2 = e2['year'], e2['month']
    if m1 == 12:
        return y2 == y1 + 1 and m2 == 1
    return y2 == y1 and m2 == m1 + 1


def _recurring_now(ents, report_year, report_month_num):
    """True only if the person is flagged in the report month AND the month
    immediately before it — old streaks must not keep resurfacing in later
    months' reports, and re-generating a past month must use that month's
    own streak, not the newest entries."""
    ry, rm = int(report_year), int(report_month_num)
    py, pm = (ry - 1, 12) if rm == 1 else (ry, rm - 1)
    have_cur  = any(int(e['year']) == ry and int(e['month']) == rm for e in ents)
    have_prev = any(int(e['year']) == py and int(e['month']) == pm for e in ents)
    return have_cur and have_prev


# ── i-STAT error code lookup ────────────────────────────────────────────────
_ISTAT_BY_CODE    = {}   # "24" → entry dict
_ISTAT_BY_MSG     = {}   # normalised display text fragment → entry dict
_ISTAT_LOADED     = False

def _load_istat_codes():
    global _ISTAT_BY_CODE, _ISTAT_BY_MSG, _ISTAT_LOADED
    if _ISTAT_LOADED:
        return
    try:
        with open(ISTAT_CODES_PATH) as f:
            data = json.load(f)
        for entry in data.get('entries', []):
            for code in entry.get('codes', []):
                _ISTAT_BY_CODE[str(code).strip()] = entry
            # Index first few significant words of the display message
            msg_key = entry.get('display', '').lower().split('/')[0].strip()
            if msg_key:
                _ISTAT_BY_MSG[msg_key] = entry
    except Exception:
        pass
    _ISTAT_LOADED = True

def _pick_err_col(cols):
    """Column holding the specific error identifier for the summary sections.
    Prefer the specific sub-code (historical behaviour — 'CODE nn' values the
    code lookup resolves), then the descriptive TEXT column, then anything
    error-shaped. Never the P/Q category column by accident."""
    for pref in ('Sub Code', 'TEXT'):
        if pref in cols:
            return pref
    return next((c for c in cols
                 if 'description' in str(c).lower() or str(c).lower() == 'text'
                 or 'error code' in str(c).lower()), None)

def _lookup_istat_error(text):
    """Return best-matching entry dict for an error text, or None."""
    _load_istat_codes()
    if not text:
        return None
    t = str(text).strip()
    # P/Q category codes (from the events category column) resolve to their
    # category names so 'Likely Cause' is never blank.
    m0 = re.match(r'^([PQ][1-8])$', t.upper())
    if m0 and m0.group(1) in ISTAT_TYPE_NAMES:
        name = ISTAT_TYPE_NAMES[m0.group(1)]
        return {'code': m0.group(1), 'short': name, 'display': name,
                'explanation': f'{name} events (error category {m0.group(1)}) '
                               'recorded on the analyser this month.'}
    # 1. Extract numeric code from text (e.g. "Code 24", "(24)", "24")
    m = re.search(r'\b(\d{1,3})\b', t)
    if m:
        entry = _ISTAT_BY_CODE.get(m.group(1))
        if entry:
            return entry
    # 2. Single uppercase/letter code (L, G, R, r, t, B)
    m2 = re.match(r'^([LGRrtB])$', t.strip())
    if m2:
        entry = _ISTAT_BY_CODE.get(m2.group(1))
        if entry:
            return entry
    tl = t.lower()
    # 3. Exact key match
    if tl in _ISTAT_BY_MSG:
        return _ISTAT_BY_MSG[tl]
    # 4. Longest-key-first substring match (most specific wins)
    candidates = [(key, entry) for key, entry in _ISTAT_BY_MSG.items()
                  if key in tl or tl in key]
    if candidates:
        return max(candidates, key=lambda x: len(x[0]))[1]
    # 5. Partial keyword match — longest matching key wins
    keywords = [w for w in re.split(r'\W+', tl) if len(w) > 4]
    kw_cands = [(key, entry) for key, entry in _ISTAT_BY_MSG.items()
                if any(kw in key for kw in keywords)]
    if kw_cands:
        return max(kw_cands, key=lambda x: len(x[0]))[1]
    return None

def _load_staff_history():
    if os.path.exists(STAFF_HISTORY_PATH):
        try:
            sh = _read_history_file(STAFF_HISTORY_PATH)
            # One-time migration: v1 entries stored total_carts as successful
            # patient cartridges only (could even be below the error count,
            # producing >100% rates). v2 stores ATTEMPTS = successes + errors.
            # Entries with no recorded denominator at all stay at 0 so the
            # report shows '—' (not assessable) rather than a made-up rate.
            for ents in sh.values():
                if isinstance(ents, list):
                    for e in ents:
                        if isinstance(e, dict) and not e.get('v2'):
                            if int(e.get('total_carts', 0) or 0) > 0:
                                e['total_carts'] = (int(e['total_carts'])
                                                    + int(e.get('error_count', 0)))
                            e['v2'] = True
            return sh
        except _HistoryLocked:
            raise
        except Exception:
            pass
    return {}

def _save_staff_history(sh):
    try:
        _write_history_file(STAFF_HISTORY_PATH, sh)
    except Exception:
        pass

def _staff_key(hospital, surname, first_name):
    return f"{hospital}|{str(surname).strip().upper()}|{str(first_name).strip().upper()}"

def _update_staff_history(sh, hospital, surname, first_name, year, month_num, error_count, total_carts=0):
    """Record month entries only for staff with ≥3 errors."""
    if int(error_count) < 3:
        return sh
    key     = _staff_key(hospital, surname, first_name)
    entries = [e for e in sh.get(key, [])
               if not (e['year'] == int(year) and e['month'] == int(month_num))]
    entries.append({
        'year':         int(year),
        'month':        int(month_num),
        'error_count':  int(error_count),
        'total_carts':  int(total_carts),   # attempts: successes + errors
        'surname':      str(surname).strip(),
        'first_name':   str(first_name).strip(),
        'v2':           True,
    })
    entries.sort(key=lambda e: (e['year'], e['month']))
    sh[key] = entries[-24:]
    return sh

def _plot_error_trend(entries, device_label):
    """Return a BytesIO PNG of the error-rate trend, or None if matplotlib missing."""
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        import matplotlib.patches as mpatches
        import calendar

        if not entries:
            return None
        entries = sorted(entries, key=lambda e: (int(e['year']), int(e['month'])))[-CHART_MONTHS:]

        labels = []
        values = []
        for e in entries:
            labels.append(f"{calendar.month_abbr[e['month']]}\n{str(e['year'])[2:]}")
            values.append(e['error_rate'])

        fig, ax = plt.subplots(figsize=(9, 5.0))
        fig.patch.set_facecolor('white')
        ax.set_facecolor('white')

        # Performance bands
        y_ceil = max(max(values) * 1.35, 14)
        ax.axhspan(0,  4,       alpha=0.07, color='#538135', zorder=0)
        ax.axhspan(4,  6,       alpha=0.07, color='#2F5496', zorder=0)
        ax.axhspan(6,  10,      alpha=0.07, color='#FF8000', zorder=0)
        ax.axhspan(10, y_ceil,  alpha=0.07, color='#FF0000', zorder=0)

        # Threshold lines
        for thresh, col in [(4, '#538135'), (6, '#2F5496'), (10, '#FF0000')]:
            ax.axhline(thresh, color=col, linewidth=0.8, linestyle='--', alpha=0.55)

        # Main line
        xs = list(range(len(labels)))
        ax.plot(xs, values, color='#305496', linewidth=3, marker='o',
                markersize=8, markerfacecolor='white',
                markeredgecolor='#305496', markeredgewidth=2, zorder=3)

        # Annotate each point with coloured value
        for xi, yi in zip(xs, values):
            _, lc = get_perf_label(yi)
            rgb = (lc[0] / 255, lc[1] / 255, lc[2] / 255)
            ax.annotate(f'{yi:.1f}%', (xi, yi),
                        textcoords='offset points', xytext=(0, 9),
                        ha='center', fontsize=11, color=rgb, fontweight='bold')

        ax.set_xticks(xs)
        ax.set_xticklabels(labels, fontsize=12)
        ax.set_xlim(-0.6, CHART_MONTHS - 0.4)      # 12 fixed slots
        ax.set_ylabel('Error Rate %', fontsize=13.5)
        ax.set_ylim(0, y_ceil)
        ax.set_title(f'Error Rate Trend  —  {device_label}  (last {CHART_MONTHS} months)',
                     fontsize=15, color='#1F497D', fontweight='bold', pad=10)
        ax.yaxis.set_major_formatter(
            plt.FuncFormatter(lambda v, _: f'{v:.0f}%'))
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.tick_params(axis='both', labelsize=8)

        legend = [
            mpatches.Patch(facecolor='#538135', alpha=0.4, label='Champions  <4%'),
            mpatches.Patch(facecolor='#2F5496', alpha=0.4, label='Excellent  4–6%'),
            mpatches.Patch(facecolor='#FF8000', alpha=0.4, label='Acceptable  6–10%'),
            mpatches.Patch(facecolor='#FF0000', alpha=0.4, label='Needs attention  >10%'),
        ]
        fig.legend(handles=legend, loc='lower center', ncol=4, fontsize=10.5,
                   frameon=False, bbox_to_anchor=(0.5, -0.01))
        fig.subplots_adjust(bottom=0.2)
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


# ── Cartridge volume & error-mix chart (ported from the ABL report) ───────────
ISTAT_ERR_TYPES = ['P1', 'P2', 'P3', 'P4', 'P5', 'P6', 'P7', 'P8',
                   'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8']
ISTAT_TYPE_NAMES = {
    'P1': 'Other',               'P2': 'Environment',
    'P3': 'Cartridge Handling',  'P4': 'Overfilled',
    'P5': 'Unable to Position',  'P6': 'Underfilled',
    'P7': 'Insufficient Sample', 'P8': 'Thermal Contact',
    'Q1': 'QC Other',               'Q2': 'QC Environment',
    'Q3': 'QC Cartridge Handling',  'Q4': 'QC Overfilled',
    'Q5': 'QC Unable to Position',  'Q6': 'QC Underfilled',
    'Q7': 'QC Insufficient Sample', 'Q8': 'QC Thermal Contact',
}
ISTAT_TYPE_COLORS = {
    'P1': '#9E9E9E', 'P2': '#35978F', 'P3': '#01485C', 'P4': '#E08214',
    'P5': '#8DA0CB', 'P6': '#C51B8A', 'P7': '#A6D854', 'P8': '#FFD92F',
    'Q1': '#5A5A5A', 'Q2': '#1B7837', 'Q3': '#11B5AE', 'Q4': '#B34700',
    'Q5': '#4A5FA5', 'Q6': '#7A1466', 'Q7': '#6B8E23', 'Q8': '#B8860B',
}
_TYPE_DARK_TEXT = {'P7', 'P8'}   # light segment colours need dark labels
CHART_MONTHS = 12                 # every trend chart/table shows the last 12 months

# Sidecar with device volumes/error mixes for months generated before this
# feature existed (rebuilt from the retained monthly exports; device-level
# only — no staff data, so it stays a plain file).
VOLUME_BACKFILL_PATH = os.path.join(APP_DIR, 'device_volume_history.json')
_VOL_BF = None

def _merge_volume_backfill(key, entries):
    """Fill carts/types into history entries recorded before volumes were kept."""
    global _VOL_BF
    try:
        if any(not e.get('carts') for e in entries):
            if _VOL_BF is None:
                _VOL_BF = {}
                if os.path.exists(VOLUME_BACKFILL_PATH):
                    # The unlock sweep encrypts every *history* file, this
                    # sidecar included — read through the decrypting loader.
                    _VOL_BF = _read_history_file(VOLUME_BACKFILL_PATH)
            bf = _VOL_BF.get(key, {})
            for e in entries:
                if not e.get('carts'):
                    mk = f"{int(e['year'])}-{int(e['month']):02d}"
                    if mk in bf:
                        e['carts'] = int(bf[mk].get('carts', 0))
                        e['types'] = bf[mk].get('types', {})
    except Exception:
        pass
    return entries

def _plot_volume_trend(entries, device_label):
    """Two-panel chart: cartridges-per-month line on top, and beneath it
    stacked monthly bars of each error category as % of that month's
    cartridges. Always laid out as 12 monthly slots (last 12 months of data),
    fixed bar width, and EVERY segment carries its percentage — inside the
    segment when it fits, beside the bar in the segment's colour when not."""
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        import matplotlib.patches as mpatches
        import calendar

        pts = sorted((e for e in entries if int(e.get('carts', 0) or 0) > 0),
                     key=lambda e: (int(e['year']), int(e['month'])))[-CHART_MONTHS:]
        if not pts:
            return None

        labels = [f"{calendar.month_abbr[e['month']]}\n{str(e['year'])[2:]}" for e in pts]
        values = [int(e['carts']) for e in pts]
        xs     = list(range(len(labels)))
        BAR_W  = 0.62

        codes_present = [c for c in ISTAT_ERR_TYPES
                         if any((e.get('types') or {}).get(c) for e in pts)]

        fig, (ax, ax2) = plt.subplots(2, 1, figsize=(9, 7.2), sharex=True,
                                      gridspec_kw={'height_ratios': [1.0, 1.45], 'hspace': 0.12})
        fig.patch.set_facecolor('white')
        for a in (ax, ax2):
            a.set_facecolor('white')
            a.set_xlim(-0.6, CHART_MONTHS - 0.4)      # 12 fixed slots
            a.spines['top'].set_visible(False)
            a.spines['right'].set_visible(False)

        # ── Top panel: cartridges-per-month line ─────────────────────────
        ax.set_ylim(0, max(values) * 1.30)
        ax.plot(xs, values, color='#538135', linewidth=3, marker='o',
                markersize=8, markerfacecolor='white',
                markeredgecolor='#538135', markeredgewidth=2, zorder=4)
        for xi, yi in zip(xs, values):
            ax.annotate(f'{yi:,}', (xi, yi), textcoords='offset points', xytext=(0, 9),
                        ha='center', fontsize=11, color='#538135', fontweight='bold')
        ax.set_ylabel('Cartridges run', fontsize=13, color='#538135')
        ax.set_title(f'Cartridges Run & Error Mix  —  {device_label}  (last {CHART_MONTHS} months)',
                     fontsize=15, color='#1F497D', fontweight='bold', pad=10)
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f'{v:,.0f}'))
        ax.tick_params(axis='y', labelsize=10, colors='#538135')
        ax.grid(axis='y', color='#D9D9D9', linewidth=0.6, alpha=0.6, zorder=0)

        # ── Bottom panel: error mix, % of that month's cartridges ────────
        month_tot = []
        for e in pts:
            codes  = {c: n for c, n in (e.get('types') or {}).items() if n}
            month_tot.append(sum(codes.values()) / e['carts'] * 100 if codes else 0.0)
        r_max = max(max(month_tot) * 1.22, 1.0)
        ax2.set_ylim(0, r_max)
        # a label needs roughly this much bar height (in % units) to sit inside
        panel_pts = 7.2 * 72 * (1.45 / 2.45) * 0.80
        min_inside = r_max * (10.0 / panel_pts)

        for i, e in enumerate(pts):
            codes  = {c: n for c, n in (e.get('types') or {}).items() if n}
            if not codes:
                continue
            bottom = 0.0
            side_y = -1.0                                # last side-label y (data units)
            for c in ISTAT_ERR_TYPES:
                n = codes.get(c, 0)
                if not n:
                    continue
                seg = n / e['carts'] * 100
                col = ISTAT_TYPE_COLORS.get(c, '#888780')
                ax2.bar([i], [seg], bottom=[bottom], width=BAR_W, color=col, zorder=2,
                        edgecolor='white', linewidth=0.5)
                mid = bottom + seg / 2
                if seg >= min_inside:
                    ax2.annotate(f'{seg:.1f}%', (i, mid), ha='center', va='center',
                                 fontsize=8.6, fontweight='bold', zorder=3,
                                 color='#3A3A00' if c in _TYPE_DARK_TEXT else 'white')
                else:
                    # too thin for an inside label: write it beside the bar, in the
                    # segment's own colour, nudged up if it would overlap the last one
                    y = max(mid, side_y + min_inside * 0.9)
                    ax2.annotate(f'{seg:.1f}%', (i + BAR_W / 2 + 0.03, y),
                                 ha='left', va='center', fontsize=7.2, fontweight='bold',
                                 color=col if c not in _TYPE_DARK_TEXT else '#8A7A00', zorder=3)
                    side_y = y
                bottom += seg
            ax2.annotate(f'{month_tot[i]:.1f}%', (i, bottom), textcoords='offset points',
                         xytext=(0, 4), ha='center', fontsize=10.5, color='#444444',
                         fontweight='bold', zorder=3)

        ax2.set_ylabel('Errors, % of cartridges', fontsize=13, color='#7A7A7A')
        ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f'{v:g}%'))
        ax2.tick_params(axis='y', labelsize=10, colors='#7A7A7A')
        ax2.set_xticks(xs)
        ax2.set_xticklabels(labels, fontsize=11)
        ax2.grid(axis='y', color='#D9D9D9', linewidth=0.6, alpha=0.6, zorder=0)

        if codes_present:
            handles = [mpatches.Patch(facecolor=ISTAT_TYPE_COLORS.get(c, '#888780'),
                                      label=f'{c}  {ISTAT_TYPE_NAMES.get(c, "")}')
                       for c in codes_present]
            fig.legend(handles=handles, loc='lower center', ncol=min(4, len(handles)),
                       fontsize=10, frameon=False, bbox_to_anchor=(0.5, -0.005))
            fig.subplots_adjust(bottom=0.17)

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
#  REPORT GENERATION — TEMPLATE-BASED
# ═══════════════════════════════════════════════════════════════════════════════

def generate_report(hospital, df_use, sim_counts, ceramic_counts, df_err,
                    report_month, report_month_num, report_year,
                    history=None, staff_history=None, error_type_history=None,
                    pats_by_staff=None, ceramic_by_device=None, title_label=None):

    hosp_use = df_use[df_use['Hospital Name'] == hospital].copy()
    if hosp_use.empty:
        return None

    # ── Per-analyser ceramic cleaning counts ─────────────────────────────
    # Every analyser in this report is listed (0 when nothing recorded) and
    # each one needs CERAMIC_REQUIRED cartridges in the month — a shortfall
    # on ANY analyser fails the whole department.
    _cer_by_id = {}
    if (ceramic_by_device is not None and not ceramic_by_device.empty
            and 'Hospital Name' in ceramic_by_device.columns
            and 'Device Name' in ceramic_by_device.columns):
        _cbd_h = ceramic_by_device[ceramic_by_device['Hospital Name'] == hospital]
        for _, _cr in _cbd_h.iterrows():
            _ck = _extract_did(str(_cr['Device Name'])) or str(_cr['Device Name']).strip()
            _cer_by_id[_ck] = _cer_by_id.get(_ck, 0) + _safe_int(_cr['Ceramic_Count'])
    _cer_devices = []
    if 'Device Name' in hosp_use.columns:
        for _cdev in sorted(str(d) for d in hosp_use['Device Name'].dropna().unique()):
            _ck = _extract_did(_cdev) or _cdev.strip()
            _cer_devices.append((_ck, _cer_by_id.get(_ck, 0)))
    _cer_failed = [(d, n) for d, n in _cer_devices if n < CERAMIC_REQUIRED]

    short = hospital.split('_', 1)[-1] if '_' in hospital else hospital
    hhs   = hospital.split('_')[0]     if '_' in hospital else hospital

    # ═══════════════════════════════════════════════════════════════
    #  PRE-COMPUTE SNAPSHOT METRICS (used in at-a-glance card)
    # ═══════════════════════════════════════════════════════════════
    _snap_total_c = 0; _snap_total_r = 0
    for _, _row in hosp_use.iterrows():
        _snap_total_c += _safe_int(_row.get('Total Carts'))
        _snap_total_r += _safe_int(_row.get('Total Res'))
    _snap_oer    = round((_snap_total_c - _snap_total_r) / _snap_total_c * 100, 1) if _snap_total_c else 0.0
    _snap_label, _snap_lc = get_perf_label(_snap_oer)
    _snap_total_err = _snap_total_c - _snap_total_r

    # Simulator
    _snap_sim   = sim_counts[sim_counts['Hospital Name'] == hospital]
    _snap_sim_n = int(_snap_sim['SIM_Runs'].sum()) if not _snap_sim.empty else 0

    # Ceramic
    _snap_cer_row = ceramic_counts[ceramic_counts['Hospital Name'] == hospital]
    _snap_cer_n   = int(_snap_cer_row['Ceramic_Count'].iloc[0]) if not _snap_cer_row.empty else 0
    if _cer_devices:
        _snap_cer_ok = not _cer_failed            # every analyser met the requirement
    else:
        _snap_cer_ok = _snap_cer_n >= CERAMIC_REQUIRED

    # Top error & repeat flag
    _snap_hosp_err_col = 'Hospital Name' if 'Hospital Name' in df_err.columns else None
    _snap_hosp_err = (df_err[df_err[_snap_hosp_err_col] == hospital].copy()
                      if _snap_hosp_err_col and not df_err.empty else pd.DataFrame())
    _snap_err_col  = _pick_err_col(_snap_hosp_err.columns) \
                      if not _snap_hosp_err.empty else None
    _snap_top_error   = None
    _snap_top_repeats = False
    if not _snap_hosp_err.empty and _snap_err_col:
        _snap_vc = _snap_hosp_err[_snap_err_col].astype(str).str.strip().value_counts()
        if not _snap_vc.empty:
            _snap_top_error = str(_snap_vc.index[0])
            if error_type_history:
                _snap_devs = (sorted(str(d) for d in hosp_use['Device Name'].dropna().unique())
                              if 'Device Name' in hosp_use.columns else [])
                _snap_et_key = f'{hospital}|{_snap_devs[0]}' if len(_snap_devs) == 1 else hospital
                _snap_prior  = [e for e in error_type_history.get(_snap_et_key, [])
                                if not (e['year'] == int(report_year)
                                        and e['month'] == int(report_month_num))]
                _snap_top_repeats = bool(_snap_prior
                                         and _snap_prior[-1].get('top_error') == _snap_top_error)

    # Staff flags
    _snap_flagged_count = 0
    if staff_history:
        _snap_prefix = hospital + '|'
        for _sk, _se in staff_history.items():
            if (_sk.startswith(_snap_prefix)
                    and _recurring_now(_se, report_year, report_month_num)):
                _snap_flagged_count += 1

    # Staff with 3+ errors THIS month (same rule as the follow-up table)
    _snap_cur_flagged = 0
    if not _snap_hosp_err.empty:
        _sn_c = next((c for c in _snap_hosp_err.columns if 'surname' in str(c).lower()), None)
        _fn_c = next((c for c in _snap_hosp_err.columns if 'first' in str(c).lower()), None)
        if _sn_c and _fn_c:
            _snap_cur_flagged = int((_snap_hosp_err.groupby([_sn_c, _fn_c]).size() >= 3).sum())

    # ── Open template and clear body (keeps header/footer/styles) ──
    doc  = Document(TEMPLATE_PATH)
    body = doc.element.body
    sect = body.find(qn('w:sectPr'))
    for child in list(body):
        if child != sect:
            body.remove(child)

    # ── Force page background to white (overrides any theme lt2 default) ──
    root = doc.element
    existing_bg = root.find(qn('w:background'))
    if existing_bg is not None:
        root.remove(existing_bg)
    bg = OxmlElement('w:background')
    bg.set(qn('w:color'), 'FFFFFF')
    bg.set(qn('w:themeColor'), 'background1')
    root.insert(0, bg)

    # ═══════════════════════════════════════════════════════════════
    #  TITLE BLOCK  (indented to clear the logo on the left)
    # ═══════════════════════════════════════════════════════════════
    LOGO_INDENT = Inches(2127 / 1440)   # 2127 DXA = 1.479"

    # "Pathology Queensland"
    p = doc.add_paragraph(style='Title')
    p.paragraph_format.left_indent = LOGO_INDENT
    _run(p, 'Pathology Queensland', size=12, color=C_DARK_BLUE)

    # "i-Stat Summary Report"  (italic i)
    p = doc.add_paragraph(style='Title')
    p.paragraph_format.left_indent = LOGO_INDENT
    _run(p, 'i', italic=True, size=20, color=C_DARK_BLUE)
    _run(p, '-Stat Summary Report', size=20, color=C_DARK_BLUE)

    # "[Hospital] – [Month]"  (green subtitle — STYLEREF DocSubTitle picks this up for footer)
    p = doc.add_paragraph(style='Subtitle')
    p.paragraph_format.first_line_indent = LOGO_INDENT
    r = _run(p, f'{title_label or short} – {report_month}', size=18, color=C_GREEN)
    try:
        r.style = doc.styles['DocSubTitle']
    except Exception:
        pass

    # ═══════════════════════════════════════════════════════════════
    #  MONTHLY SNAPSHOT  (at-a-glance status card)
    # ═══════════════════════════════════════════════════════════════
    _heading(doc, 'Monthly Snapshot')

    snap_tbl = doc.add_table(rows=0, cols=3)
    _add_borders(snap_tbl)
    _set_col_widths(snap_tbl, [2700, 5100, 2560])   # full printable width
    _cell_padding(snap_tbl)

    # Header row
    _snap_hdr = snap_tbl.add_row().cells
    for _sc, _sl in zip(_snap_hdr, ['Indicator', 'Value', 'Status']):
        _hdr_cell(_sc, _sl, FILL_GREY_HDR)

    def _snap_row(tbl, label, value_str, status_str, status_color=None, fill=FILL_WHITE):
        _r = tbl.add_row().cells
        _data_cell(_r[0], label,      fill, C_BLUE,  bold=True, size=9)
        _data_cell(_r[1], value_str,  fill, C_BLUE,  size=9)
        _data_cell(_r[2], status_str, fill, status_color or C_BLUE, bold=True, size=9)

    # Row 1 – Error Rate
    _snap_row(snap_tbl,
              'Overall Error Rate',
              f'{_snap_oer:.1f}%  ({_snap_total_err} errors / {_snap_total_c} cartridges)',
              _snap_label,
              _snap_lc,
              fill=FILL_ALT_ROW)

    # Row 2 – Simulator, one line per analyser
    _sim_by_id = {}
    if not _snap_sim.empty and 'Device Name' in _snap_sim.columns:
        for _, _sr in _snap_sim.iterrows():
            _sk = _extract_did(str(_sr['Device Name'])) or str(_sr['Device Name']).strip()
            _sim_by_id[_sk] = _sim_by_id.get(_sk, 0) + _safe_int(_sr['SIM_Runs'])
    _sim_lines = [(_ck, _sim_by_id.get(_ck, 0)) for _ck, _ in _cer_devices] \
                 or [(_k, _v) for _k, _v in sorted(_sim_by_id.items())]
    if _sim_lines:
        _sim_val = '\n'.join(
            f'{_ck} — {_n} run{"s" if _n != 1 else ""}' if _n else f'{_ck} — no runs recorded'
            for _ck, _n in _sim_lines)
        _sim_none = [_ck for _ck, _n in _sim_lines if not _n]
        if not _sim_none:
            _sim_status, _sim_color = '✓ All analysers', C_GREEN_PASS
        else:
            _sim_status = (f'⚠ {len(_sim_none)} analyser'
                           f'{"s" if len(_sim_none) != 1 else ""} with no runs')
            _sim_color  = C_RED_TEXT
    else:
        _sim_val, _sim_status, _sim_color = 'No records found', 'Data unavailable', C_RED_TEXT
    _snap_row(snap_tbl, 'Simulator', _sim_val, _sim_status, _sim_color)

    # Row 3 – Ceramic
    _n_cer_dev = len(_cer_devices)
    if _snap_cer_ok:
        if _n_cer_dev > 1:
            _cer_val = (f'All {_n_cer_dev} analysers completed '
                        f'({CERAMIC_REQUIRED} cleaning cartridges each)')
        else:
            _cer_val = f'{_snap_cer_n} cleaning cartridges recorded'
        _cer_status = '✓ Completed'
        _cer_color  = C_GREEN_PASS
    else:
        if _cer_devices:
            _short = ', '.join(f'{d} ({n} of {CERAMIC_REQUIRED})' for d, n in _cer_failed[:5])
            if len(_cer_failed) > 5:
                _short += '…'
            _plural = 's' if _n_cer_dev != 1 else ''
            _cer_val = (f'{len(_cer_failed)} of {_n_cer_dev} analyser{_plural} '
                        f'below the {CERAMIC_REQUIRED} required: {_short}')
        else:
            _cer_val = f'{_snap_cer_n} of {CERAMIC_REQUIRED} required'
        _cer_status = '✗ Fail'
        _cer_color  = C_RED_TEXT
    _snap_row(snap_tbl, 'Ceramic Cleaning', _cer_val, _cer_status, _cer_color,
              fill=FILL_ALT_ROW)

    # Row 4 – Top error
    _load_istat_codes()
    if _snap_top_error:
        _te_entry  = _lookup_istat_error(_snap_top_error)
        _te_label  = (_te_entry['short'] if _te_entry else _snap_top_error)[:60]
        _te_val    = f'{_te_label}  [{_snap_top_error}]' if _te_entry else _snap_top_error
        _te_status = '⚠ Repeating error' if _snap_top_repeats else 'New this month'
        _te_color  = C_RED_TEXT if _snap_top_repeats else C_BLUE
    else:
        _te_val    = 'No errors recorded'
        _te_status = '✓ Clear'
        _te_color  = C_GREEN_PASS
    _snap_row(snap_tbl, 'Top Error Code', _te_val, _te_status, _te_color)

    # Row 5a – Staff flagged this month
    if _snap_cur_flagged:
        _cf_val    = (f'{_snap_cur_flagged} staff member'
                      f'{"s" if _snap_cur_flagged != 1 else ""} with 3+ errors this month')
        _cf_status = '⚠ Review'
        _cf_color  = C_RED_TEXT
    else:
        _cf_val    = 'No staff with 3+ errors this month'
        _cf_status = '✓ Clear'
        _cf_color  = C_GREEN_PASS
    _snap_row(snap_tbl, 'Staff Flagged This Month', _cf_val, _cf_status, _cf_color)

    # Row 5b – Recurring staff alerts
    if _snap_flagged_count:
        _sf_val    = (f'{_snap_flagged_count} staff member'
                      f'{"s" if _snap_flagged_count != 1 else ""} flagged in consecutive months')
        _sf_status = '⚠ Follow-up required'
        _sf_color  = C_RED_TEXT
    else:
        _sf_val    = 'No recurring issues detected'
        _sf_status = '✓ Clear'
        _sf_color  = C_GREEN_PASS
    _snap_row(snap_tbl, 'Recurring Staff Alerts', _sf_val, _sf_status, _sf_color,
              fill=FILL_ALT_ROW)

    doc.add_paragraph()

    # ═══════════════════════════════════════════════════════════════
    #  HISTORICAL PERFORMANCE (placed here for quick snapshot view)
    # ═══════════════════════════════════════════════════════════════
    _heading(doc, 'Historical Performance')

    if history:
        devices_plotted = []
        dev_names = (hosp_use['Device Name'].dropna().unique()
                     if 'Device Name' in hosp_use.columns else [])
        for dev in sorted(str(d) for d in dev_names):
            key     = _history_key(hospital, dev)
            entries = history.get(key, [])
            if not entries:
                continue
            did     = _extract_did(dev) or dev
            entries = _merge_volume_backfill(key, entries)
            chart   = _plot_error_trend(entries, did)
            if chart:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.left_indent = Inches(0)
                run = p.add_run()
                run.add_picture(chart, width=Inches(7.2))   # full printable width
                devices_plotted.append(did)
            vchart = _plot_volume_trend(entries, did)
            if vchart:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.left_indent = Inches(0)
                run = p.add_run()
                run.add_picture(vchart, width=Inches(7.2))   # full printable width
        if not devices_plotted:
            p = doc.add_paragraph()
            _run(p, 'Trend data will appear here as monthly reports are generated.',
                 size=9, color=C_GREY_TEXT)
    else:
        p = doc.add_paragraph()
        _run(p, 'Trend data will appear here as monthly reports are generated.',
             size=9, color=C_GREY_TEXT)

    doc.add_paragraph()

    # ═══════════════════════════════════════════════════════════════
    #  PERFORMANCE SUMMARY
    # ═══════════════════════════════════════════════════════════════
    _heading(doc, 'Performance Summary')

    # Table: Hospital Name | Year | Month | Device Name | Total Carts | Error Count | Error Rate %
    COL_W = [2000, 800, 800, 2100, 1200, 1200, 1500]   # DXA (7 cols)
    tbl   = doc.add_table(rows=1, cols=7)
    _add_borders(tbl)
    _set_col_widths(tbl, COL_W)

    hdr_labels = ['Hospital Name', 'Year', 'Month', 'Device Name', 'Total Carts', 'Error Count', 'Error Rate %']
    for i, (c, lbl) in enumerate(zip(tbl.rows[0].cells, hdr_labels)):
        fill = FILL_RED_HDR if i >= 5 else FILL_GREY_HDR
        _hdr_cell(c, lbl, fill)

    total_c = total_r_val = 0
    for row_i, (_, row) in enumerate(hosp_use.iterrows()):
        tc  = _safe_int(row.get('Total Carts'))
        tr  = _safe_int(row.get('Total Res'))
        er  = float(row.get('ErrorRate') or 0)
        total_c    += tc
        total_r_val += tr
        label, lc = get_perf_label(er)

        dr   = tbl.add_row().cells
        fill = FILL_BLUE_HDR if row_i % 2 == 1 else FILL_WHITE

        ec = tc - tr   # error count
        _data_cell(dr[0], hospital  if row_i == 0 else '', fill, C_BLUE)
        _data_cell(dr[1], str(report_year)      if row_i == 0 else '', fill, C_BLUE, align=WD_ALIGN_PARAGRAPH.CENTER)
        _data_cell(dr[2], str(report_month_num) if row_i == 0 else '', fill, C_BLUE, align=WD_ALIGN_PARAGRAPH.CENTER)
        _data_cell(dr[3], str(row.get('Device Name', '')).strip(), fill, C_BLUE)
        _data_cell(dr[4], str(tc), fill, C_BLUE, align=WD_ALIGN_PARAGRAPH.CENTER)
        _data_cell(dr[5], str(ec), fill, C_BLUE, align=WD_ALIGN_PARAGRAPH.CENTER)

        # Error Rate cell: coloured label
        _shd(dr[6], fill)
        p5 = dr[6].paragraphs[0]; p5.clear()
        p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r5a = p5.add_run(f'{er:.1f}%  ')
        r5a.font.size = Pt(9); r5a.font.color.rgb = lc
        r5b = p5.add_run(label)
        r5b.font.size = Pt(9); r5b.bold = True; r5b.font.color.rgb = lc

    # Compute overall rate (always needed for legend, even if TOTAL row is hidden)
    oer    = round((total_c - total_r_val) / total_c * 100, 1) if total_c else 0.0
    olabel, olc = get_perf_label(oer)

    # Overall total row — only shown when there are multiple devices
    if total_c and len(hosp_use) > 1:
        gr = tbl.add_row().cells
        total_ec = total_c - total_r_val
        _data_cell(gr[0], 'TOTAL', FILL_ALT_ROW, C_BLUE, bold=True)
        _data_cell(gr[1], '', FILL_ALT_ROW)
        _data_cell(gr[2], '', FILL_ALT_ROW)
        _data_cell(gr[3], '', FILL_ALT_ROW)
        _data_cell(gr[4], str(total_c),  FILL_ALT_ROW, C_BLUE, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        _data_cell(gr[5], str(total_ec), FILL_ALT_ROW, C_BLUE, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        _shd(gr[6], FILL_ALT_ROW)
        p5 = gr[6].paragraphs[0]; p5.clear()
        p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r5a = p5.add_run(f'{oer:.1f}%  ')
        r5a.bold = True; r5a.font.size = Pt(9); r5a.font.color.rgb = olc
        r5b = p5.add_run(olabel)
        r5b.bold = True; r5b.font.size = Pt(9); r5b.font.color.rgb = olc

    # Cartridge Error rate legend
    doc.add_paragraph()
    p = doc.add_paragraph()
    _run(p, 'Cartridge Error rate: ', bold=True, size=10)
    _run(p, f'{get_perf_label(oer if total_c else 0)[0]}  ', bold=True, size=10,
         color=get_perf_label(oer if total_c else 0)[1])
    _run(p, 'Champions', size=9, color=RGBColor(0x53, 0x81, 0x35))
    _run(p, ':<4%   ', size=9)
    _run(p, 'Excellent', size=9, color=RGBColor(0x2F, 0x54, 0x96))
    _run(p, ':4-6%   ', size=9)
    _run(p, 'Acceptable', size=9, color=RGBColor(0xFF, 0x80, 0x00))
    _run(p, ':6-10%   ', size=9)
    _run(p, 'Needs attention', size=9, color=RGBColor(0xFF, 0x00, 0x00))
    _run(p, ':>10%', size=9)

    # ═══════════════════════════════════════════════════════════════
    #  SIMULATORS RUN
    # ═══════════════════════════════════════════════════════════════
    _heading(doc, 'Simulators run:')

    hosp_sim = sim_counts[sim_counts['Hospital Name'] == hospital]
    if not hosp_sim.empty:
        # Template sim table: DeviceID label + count, full width, no visible header
        sim_tbl = doc.add_table(rows=0, cols=2)
        _add_borders(sim_tbl)
        _set_col_widths(sim_tbl, [7500, 2460])
        for _, sr in hosp_sim.iterrows():
            row = sim_tbl.add_row().cells
            did = str(sr['Device Name'])
            # Extract just the numeric device ID for display (matching template)
            did_num = _extract_did(did) or did
            label = f'{did_num} Total'
            _data_cell(row[0], label, FILL_WHITE, C_BLUE, size=9)
            _data_cell(row[1], str(int(sr['SIM_Runs'])), FILL_WHITE, C_BLUE, size=9,
                       align=WD_ALIGN_PARAGRAPH.RIGHT)
    else:
        p = doc.add_paragraph()
        _run(p, 'Simulator data unavailable.', size=9, color=C_GREY_TEXT)

    p_note = doc.add_paragraph()
    _run(p_note, 'Please note: ', bold=True, size=9)
    _run(p_note, 'Simulator must be run ', size=9)
    _run(p_note, 'once', bold=True, italic=True, size=9)
    _run(p_note, ' daily or before a patient specimen if i-STAT is used infrequently.', size=9)

    doc.add_paragraph()   # space before Monthly Ceramic section

    # ═══════════════════════════════════════════════════════════════
    #  MONTHLY CERAMIC CLEANING PROCEDURE
    # ═══════════════════════════════════════════════════════════════
    cer_row  = ceramic_counts[ceramic_counts['Hospital Name'] == hospital]
    ceramic_n = int(cer_row['Ceramic_Count'].iloc[0]) if not cer_row.empty else 0
    completed = (not _cer_failed) if _cer_devices else (ceramic_n >= CERAMIC_REQUIRED)

    ph = doc.add_paragraph()
    _run(ph, 'Monthly Ceramic Cleaning Procedure', bold=True, size=12, underline=True, color=C_BLUE)
    _run(ph, ': ', size=12)
    if completed:
        _run(ph, 'Completed ✓', bold=True, size=10, color=C_GREEN_PASS)
    else:
        _run(ph, 'FAIL', bold=True, size=10, color=C_RED_TEXT)
        if _cer_devices:
            _pl = 's' if len(_cer_devices) != 1 else ''
            _run(ph, f'  ({len(_cer_failed)} of {len(_cer_devices)} analyser{_pl} '
                     f'below the {CERAMIC_REQUIRED} required)', size=9, color=C_GREY_TEXT)
        else:
            _run(ph, f'  ({ceramic_n} of {CERAMIC_REQUIRED} required)', size=9, color=C_GREY_TEXT)

    # One row per analyser, mirroring the simulator table
    if _cer_devices:
        cer_tbl = doc.add_table(rows=0, cols=3)
        _add_borders(cer_tbl)
        _set_col_widths(cer_tbl, [5500, 1700, 2760])
        for _cd, _cn in _cer_devices:
            _ok  = _cn >= CERAMIC_REQUIRED
            _col = C_BLUE if _ok else C_RED_TEXT
            crow = cer_tbl.add_row().cells
            _data_cell(crow[0], f'{_cd} Total', FILL_WHITE, _col, bold=not _ok, size=9)
            _data_cell(crow[1], str(_cn), FILL_WHITE, _col, bold=not _ok, size=9,
                       align=WD_ALIGN_PARAGRAPH.RIGHT)
            _data_cell(crow[2],
                       '✓ Completed' if _ok else f'✗ Fail ({_cn} of {CERAMIC_REQUIRED})',
                       FILL_WHITE, C_GREEN_PASS if _ok else C_RED_TEXT, bold=True, size=9,
                       align=WD_ALIGN_PARAGRAPH.CENTER)

    pn2 = doc.add_paragraph()
    _run(pn2, 'Please note: ', bold=True, size=9)
    _run(pn2, 'Ceramic cleaning procedure must be completed ', size=9)
    _run(pn2, 'once', bold=True, italic=True, size=9)
    _run(pn2, ' per month on ', size=9)
    _run(pn2, 'every', bold=True, italic=True, size=9)
    _run(pn2, ' i-STAT ', size=9)
    _run(pn2, '(3 x cleaning cartridges in succession)', bold=True, size=9)
    _run(pn2, ' and be followed by a successful simulator.', size=9)

    doc.add_paragraph()   # space before Cartridge Error Details section

    # ═══════════════════════════════════════════════════════════════
    #  CARTRIDGE ERROR DETAILS
    # ═══════════════════════════════════════════════════════════════
    hosp_err_col = 'Hospital Name' if 'Hospital Name' in df_err.columns else None
    hosp_err = (df_err[df_err[hosp_err_col] == hospital].copy()
                if hosp_err_col and not df_err.empty else pd.DataFrame())

    # Column name resolution
    loc_col  = next((c for c in hosp_err.columns if 'location' in c.lower()), None)
    sn_col   = next((c for c in hosp_err.columns if 'surname'  in c.lower()), None)
    fn_col   = next((c for c in hosp_err.columns if 'first'    in c.lower()), None)
    err_col  = _pick_err_col(hosp_err.columns)
    hhs_col  = next((c for c in hosp_err.columns if c.upper() == 'HHS'), None)

    if not hosp_err.empty and loc_col and sn_col and fn_col and err_col:
        total_errors = len(hosp_err)
        ph = doc.add_paragraph()
        _run(ph, 'Cartridge Error Details', bold=True, size=12, underline=True, color=C_BLUE)
        _run(ph, f': Total = {total_errors}', bold=True, size=12)

        # ── Staff Requiring Follow-up ────────────────────────────────────────────
        import calendar as _cal

        # Per-staff PATS lookup (loaded from the permonthperuser sheet in load_data)
        _pats_lookup = pats_by_staff or {}

        print(f'[PATS] generate_report for "{hospital}": pats_lookup has {len(_pats_lookup)} entries')
        if _pats_lookup:
            print(f'[PATS] Sample keys: {list(_pats_lookup.keys())[:5]}')

        def _get_attempt_base(surname, firstname, uid=None):
            """Successful-cartridge count (Pats + QC) for one person from the
            permonthperuser sheet. Failed cartridges are NOT in that sheet — they
            appear only as error events — so error % = errors / (base + errors)."""
            if uid is not None and str(uid).strip():
                by_uid = _pats_lookup.get(('UID', _norm_uid(uid)))
                if by_uid:
                    return by_uid
            hosp_k = hospital.strip().upper()
            sn_k   = str(surname).strip().upper()
            fn_k   = str(firstname).strip().upper()
            # Try hospital-qualified key first, then name-only fallback
            result = (_pats_lookup.get((hosp_k, sn_k, fn_k))
                      or _pats_lookup.get(('', sn_k, fn_k))
                      or 0)
            print(f'[PATS] lookup ({hosp_k!r},{sn_k!r},{fn_k!r}) -> {result}')
            return result

        def _group_uid(grp):
            """First non-blank User ID in a grouped set of event rows, if any."""
            if 'User ID' in grp.columns:
                ids = grp['User ID'].dropna()
                if len(ids):
                    return ids.iloc[0]
            return None

        # Group 1 — current month: any staff with ≥3 errors this month
        _cur_flagged = {}   # (sn, fn) -> entry dict
        if not hosp_err.empty and sn_col and fn_col:
            for (_csn, _cfn), _cgrp in hosp_err.groupby([sn_col, fn_col]):
                _csn, _cfn = str(_csn).strip(), str(_cfn).strip()
                _cec = len(_cgrp)
                if _cec >= 3:
                    # total_carts = ATTEMPTS: successful cartridges + the failed ones
                    _ctc = _get_attempt_base(_csn, _cfn, _group_uid(_cgrp)) + _cec
                    _cur_flagged[(_csn, _cfn)] = {
                        'year': int(report_year), 'month': int(report_month_num),
                        'error_count': _cec, 'total_carts': _ctc,
                        'surname': _csn, 'first_name': _cfn,
                    }

        # Group 2 — recurring: ≥3 errors in two consecutive months (history-based)
        _consec_flagged = {}   # (sn, fn) -> entries list
        if staff_history:
            _sfprefix = hospital + '|'
            for _sfkey, _sfents in staff_history.items():
                if not _sfkey.startswith(_sfprefix):
                    continue
                if _recurring_now(_sfents, report_year, report_month_num):
                    _rfsn = _sfents[-1]['surname']
                    _rffn = _sfents[-1]['first_name']
                    _consec_flagged[(_rfsn, _rffn)] = _sfents

        _any_flagged = bool(_cur_flagged or _consec_flagged)

        if _any_flagged:
            doc.add_paragraph()
            _heading(doc, 'Staff Requiring Follow-up')

            def _render_fu_table(rows_data):
                tbl = doc.add_table(rows=1, cols=6)
                _add_borders(tbl)
                _set_col_widths(tbl, [1800, 1600, 1300, 700, 700, 1360])
                for _hc, _hl in zip(tbl.rows[0].cells,
                                     ['Surname', 'First Name', 'Month', 'Year', 'Errors', 'Error %']):
                    _hdr_cell(_hc, _hl, FILL_BLUE_HDR, C_BLUE)
                # Live attempts lookup for the current month: successes + failures
                _live_pats = {}
                if not hosp_err.empty and sn_col and fn_col:
                    for (_lsn, _lfn), _lgrp in hosp_err.groupby([sn_col, fn_col]):
                        _lsn2, _lfn2 = str(_lsn).strip(), str(_lfn).strip()
                        _live_pats[(_lsn2.upper(), _lfn2.upper())] = (
                            _get_attempt_base(_lsn2, _lfn2, _group_uid(_lgrp)) + len(_lgrp))

                for _rsn, _rfn, _rents in sorted(rows_data):
                    for _ri, _re in enumerate(_rents):
                        _iscur = (_re['year'] == int(report_year)
                                  and _re['month'] == int(report_month_num))
                        _fill  = FILL_WHITE if _ri % 2 == 0 else FILL_BLUE_HDR
                        _color = C_RED_TEXT if _iscur else C_BLUE
                        # For the current month row, prefer the live PATS value over stored
                        _tc = (_live_pats.get((_rsn.upper(), _rfn.upper()), 0)
                               if _iscur else _re.get('total_carts', 0))
                        _ec    = _re['error_count']
                        _pct   = f'{_ec / _tc * 100:.1f}%' if _tc else '—'
                        _dr    = tbl.add_row().cells
                        _data_cell(_dr[0], _rsn if _ri == 0 else '', _fill, _color, size=9)
                        _data_cell(_dr[1], _rfn if _ri == 0 else '', _fill, _color, size=9)
                        _data_cell(_dr[2], _cal.month_name[_re['month']], _fill, _color, size=9)
                        _data_cell(_dr[3], str(_re['year']), _fill, _color, size=9,
                                   align=WD_ALIGN_PARAGRAPH.CENTER)
                        _data_cell(_dr[4], str(_ec), _fill, _color, size=9,
                                   align=WD_ALIGN_PARAGRAPH.CENTER)
                        _data_cell(_dr[5], _pct, _fill, _color, size=9,
                                   align=WD_ALIGN_PARAGRAPH.CENTER)

            # Sub-section A: this month ≥3 errors
            if _cur_flagged:
                pa = doc.add_paragraph()
                _run(pa, 'This Month  ', bold=True, size=9, color=C_RED_TEXT)
                _run(pa, '— Staff with 3 or more errors recorded this period:',
                     size=9, color=C_GREY_TEXT)
                _render_fu_table([(_k[0], _k[1], [_v]) for _k, _v in sorted(_cur_flagged.items())])

            # Sub-section B: recurring across consecutive months
            if _consec_flagged:
                if _cur_flagged:
                    doc.add_paragraph()   # breathing room between the two tables
                pb = doc.add_paragraph()
                _run(pb, 'Recurring  ', bold=True, size=9, color=C_RED_TEXT)
                _run(pb, '— Staff with 3 or more errors across consecutive months:',
                     size=9, color=C_GREY_TEXT)
                _render_fu_table([(_k[0], _k[1], _v) for _k, _v in sorted(_consec_flagged.items())])

        doc.add_paragraph()
        pl = doc.add_paragraph()
        _run(pl, 'All cartridge errors recorded this month', bold=True, size=10, color=C_BLUE)

        # Group: HHS | Location | Surname | FirstName | ErrorType → count
        group_cols = [c for c in [hhs_col, loc_col, sn_col, fn_col, err_col] if c]
        grouped = (hosp_err.groupby(group_cols, dropna=False)
                   .size().reset_index(name='Count'))
        grouped[err_col] = grouped[err_col].astype(str).str.strip()
        grouped = grouped.sort_values(group_cols)

        # Pre-compute total errors per person for red-highlighting rule
        person_totals = (hosp_err.groupby([sn_col, fn_col])
                         .size().to_dict())

        # Table: HHS | Location | Year | Month | Surname | First Name | Error | Count
        # Column widths from template: 1247, 1083, 920, 800, 1447, 1308, 2617, 730
        ERR_COLS = [1100,  900, 720, 700, 1200, 1150, 3652, 730]
        et = doc.add_table(rows=1, cols=8)
        _add_borders(et)
        _set_col_widths(et, ERR_COLS)

        hdr_lbls = ['HHS', 'Location', 'Year', 'Month', 'Surname', 'First Name', 'Error', 'Count']
        for c, lbl in zip(et.rows[0].cells, hdr_lbls):
            _hdr_cell(c, lbl, FILL_BLUE_HDR, C_BLUE)

        _load_istat_codes()   # ensure index is ready
        prev_hhs = prev_loc = prev_sn = prev_fn = None
        for row_idx, (_, er) in enumerate(grouped.iterrows()):
            cur_hhs = str(er[hhs_col]).strip() if hhs_col else hhs
            cur_loc = str(er[loc_col]).strip() if loc_col else ''
            cur_sn  = str(er[sn_col]).strip()
            cur_fn  = str(er[fn_col]).strip()
            cur_err = str(er[err_col]).strip()
            cur_cnt = int(er['Count'])
            _ec     = _lookup_istat_error(cur_err)
            cur_err_disp = (f"{_ec['short']}  [{cur_err}]" if _ec
                            else cur_err)

            new_hhs = cur_hhs != prev_hhs
            new_loc = new_hhs or cur_loc != prev_loc
            new_sn  = new_loc or cur_sn  != prev_sn
            new_fn  = new_sn  or cur_fn  != prev_fn

            # Red text for staff with ≥3 total errors
            person_total = person_totals.get((cur_sn, cur_fn), 0)
            hi_color = C_RED_TEXT if person_total >= 3 else None

            fill = FILL_ALT_ROW if row_idx % 2 == 0 else FILL_WHITE

            dr = et.add_row().cells
            _data_cell(dr[0], cur_hhs           if new_hhs else '', fill, hi_color, size=9)
            _data_cell(dr[1], cur_loc           if new_loc else '', fill, hi_color, size=9)
            _data_cell(dr[2], str(report_year)  if new_loc else '', fill, size=9,
                       align=WD_ALIGN_PARAGRAPH.CENTER)
            _data_cell(dr[3], str(report_month_num) if new_loc else '', fill, size=9,
                       align=WD_ALIGN_PARAGRAPH.CENTER)
            _data_cell(dr[4], cur_sn            if new_sn  else '', fill, hi_color, size=9)
            _data_cell(dr[5], cur_fn            if new_fn  else '', fill,
                       hi_color if hi_color else None, bold=(hi_color is not None), size=9)
            _data_cell(dr[6], cur_err_disp, fill, hi_color, size=9)
            _data_cell(dr[7], str(cur_cnt), fill, hi_color, size=9,
                       align=WD_ALIGN_PARAGRAPH.CENTER)

            prev_hhs, prev_loc, prev_sn, prev_fn = cur_hhs, cur_loc, cur_sn, cur_fn

        # Grand total row
        gr = et.add_row().cells
        _data_cell(gr[0], 'Grand Total', FILL_ALT_ROW, bold=True, size=9)
        for i in range(1, 7):
            _data_cell(gr[i], '', FILL_ALT_ROW, size=9)
        _data_cell(gr[7], str(total_errors), FILL_ALT_ROW, bold=True, size=9,
                   align=WD_ALIGN_PARAGRAPH.CENTER)

        # ── Error Type Summary ──────────────────────────────────────────
        import calendar as _cal2
        doc.add_paragraph()
        ps = doc.add_paragraph()
        _run(ps, f'Error Type Summary — {report_month}', bold=True, size=11, underline=True, color=C_BLUE)

        err_summary = (
            hosp_err[err_col].astype(str).str.strip()
            .value_counts()
            .reset_index()
        )
        err_summary.columns = ['Error Type', 'Count']
        err_summary = err_summary.sort_values('Count', ascending=False).head(3)

        # Derive history key (per-device for single-device reports, else hospital)
        if 'Device Name' in hosp_use.columns:
            _devs = sorted(str(d) for d in hosp_use['Device Name'].dropna().unique())
            et_key = f'{hospital}|{_devs[0]}' if len(_devs) == 1 else hospital
        else:
            et_key = hospital

        # Current top error
        cur_top = str(err_summary.iloc[0]['Error Type']) if not err_summary.empty else None
        cur_counts = {str(r['Error Type']): int(r['Count'])
                      for _, r in err_summary.iterrows()}

        # Persist to history (mutates dict in place — saved by run_generation)
        if error_type_history is not None:
            _update_error_type_history(error_type_history, et_key,
                                       report_year, report_month_num, cur_counts)
            et_entries   = error_type_history.get(et_key, [])
            prior_months = [e for e in et_entries
                            if not (e['year'] == int(report_year)
                                    and e['month'] == int(report_month_num))]
            prev_top = prior_months[-1]['top_error'] if prior_months else None
        else:
            et_entries = []
            prev_top   = None

        # Flag if top error repeats from previous month
        top_repeats = (cur_top and prev_top and cur_top == prev_top)

        # Pre-fetch code lookups for top errors
        _load_istat_codes()
        err_lookup = {str(r['Error Type']): _lookup_istat_error(str(r['Error Type']))
                      for _, r in err_summary.iterrows()}

        # Table: Error Type | Cause (truncated) | Count
        st = doc.add_table(rows=1, cols=3)
        _add_borders(st)
        _set_col_widths(st, [2400, 5100, 960])
        for c, lbl in zip(st.rows[0].cells, ['Error Type', 'Likely Cause', 'Count']):
            _hdr_cell(c, lbl, FILL_BLUE_HDR, C_BLUE)

        for si, (_, sr) in enumerate(err_summary.iterrows()):
            fill       = FILL_ALT_ROW if si % 2 == 0 else FILL_WHITE
            is_top_rep = (si == 0 and top_repeats)
            row_color  = C_RED_TEXT if is_top_rep else C_BLUE
            dr2        = st.add_row().cells
            err_text   = str(sr['Error Type'])
            if is_top_rep:
                err_text += '  ⚠'
            entry      = err_lookup.get(str(sr['Error Type']))
            short_desc = (entry['short'] if entry else '') [:80]
            _data_cell(dr2[0], err_text,         fill, row_color, size=9)
            _data_cell(dr2[1], short_desc,       fill, row_color, size=9)
            _data_cell(dr2[2], str(sr['Count']), fill, row_color, size=9,
                       align=WD_ALIGN_PARAGRAPH.CENTER)

        # Total row
        tr2 = st.add_row().cells
        _data_cell(tr2[0], 'Total', FILL_ALT_ROW, C_BLUE, bold=True, size=9)
        _data_cell(tr2[1], '', FILL_ALT_ROW, size=9)
        _data_cell(tr2[2], str(total_errors), FILL_ALT_ROW, C_BLUE, bold=True, size=9,
                   align=WD_ALIGN_PARAGRAPH.CENTER)

        # ── Full explanation of top error codes ─────────────────────
        explained = [(str(r['Error Type']), int(r['Count']), err_lookup.get(str(r['Error Type'])))
                     for _, r in err_summary.iterrows()
                     if err_lookup.get(str(r['Error Type']))]
        if explained:
            doc.add_paragraph()
            ph3 = doc.add_paragraph()
            _run(ph3, 'Top Error Code Analysis', bold=True, size=11,
                 underline=True, color=C_BLUE)
            for err_text, cnt, entry in explained:
                is_rep = (err_text == cur_top and top_repeats)
                p_err  = doc.add_paragraph()
                _run(p_err, f'{err_text}  ({cnt} occurrences)',
                     bold=True, size=10,
                     color=C_RED_TEXT if is_rep else C_BLUE)
                p_disp = doc.add_paragraph()
                _run(p_disp, f'Display message: {entry["display"]}',
                     italic=True, size=9, color=C_GREY_TEXT)
                p_exp  = doc.add_paragraph()
                _run(p_exp, entry['explanation'], size=9, color=C_GREY_TEXT)

        # ── Month-to-month top error trend (last 12 months) ─────────────
        et_entries = sorted(et_entries, key=lambda e: (int(e['year']), int(e['month'])))[-CHART_MONTHS:]
        if len(et_entries) > 1:
            doc.add_paragraph()
            ph2 = doc.add_paragraph()
            _run(ph2, f'Error Type Trend — last {CHART_MONTHS} months', bold=True, size=11,
                 underline=True, color=C_BLUE)

            trend_tbl = doc.add_table(rows=1, cols=4)
            _add_borders(trend_tbl)
            _set_col_widths(trend_tbl, [1400, 800, 4800, 960])
            for c, lbl in zip(trend_tbl.rows[0].cells,
                               ['Month', 'Year', 'Top Error Type', 'Count']):
                _hdr_cell(c, lbl, FILL_BLUE_HDR, C_BLUE)

            for ti, te in enumerate(et_entries):
                is_cur  = (te['year'] == int(report_year)
                           and te['month'] == int(report_month_num))
                prev_te = et_entries[ti - 1] if ti > 0 else None
                repeat  = (prev_te and te['top_error']
                           and te['top_error'] == prev_te.get('top_error'))
                color   = C_RED_TEXT if repeat else C_BLUE
                fill    = FILL_ALT_ROW if ti % 2 == 0 else FILL_WHITE
                top_cnt  = te['counts'].get(te['top_error'], '') if te['top_error'] else ''
                _te      = _lookup_istat_error(te['top_error']) if te['top_error'] else None
                _te_eng  = _te['short'] if _te else te['top_error'] or ''
                _te_raw  = f'  [{te["top_error"]}]' if te['top_error'] else ''
                top_lbl  = (_te_eng + _te_raw + '  ⚠') if repeat else (_te_eng + _te_raw)
                tr3 = trend_tbl.add_row().cells
                _data_cell(tr3[0], _cal2.month_name[te['month']], fill, color, size=9)
                _data_cell(tr3[1], str(te['year']), fill, color, size=9,
                           align=WD_ALIGN_PARAGRAPH.CENTER)
                _data_cell(tr3[2], top_lbl, fill, color, size=9)
                _data_cell(tr3[3], str(top_cnt), fill, color, size=9,
                           align=WD_ALIGN_PARAGRAPH.CENTER)

    else:
        ph = doc.add_paragraph()
        _run(ph, 'Cartridge Error Details: No errors recorded this period. ✓',
             bold=True, size=12, color=C_GREEN_PASS)


    doc.add_paragraph()
    p = doc.add_paragraph()
    _run(p, 'Online training resources: ', bold=True, size=9)
    url = 'http://qheps.health.qld.gov.au/hsq/pathology/testing/point-of-care.htm'
    _run(p, url, size=9, color=RGBColor(0x00, 0x00, 0xFF))

    return doc



# ═══════════════════════════════════════════════════════════════════════════════
#  SITE OVERVIEW + AREA SUMMARY
#  Site overview: for a per-device site (Townsville, Cairns, …) one page listing
#  every ward's analyser with its error %, simulator runs and ceramic status.
#  Area summary: one page per HHS folder listing every hospital in that area
#  with the same figures, plus the analysers that missed ceramic / simulator.
# ═══════════════════════════════════════════════════════════════════════════════
HHS_NAMES = {
    'CA': 'Cairns and Hinterland', 'CQLD': 'Central Queensland', 'CWest': 'Central West',
    'ChildQLD': "Children's Health Queensland", 'DaDo': 'Darling Downs',
    'GoCo': 'Gold Coast', 'Mackay': 'Mackay', 'MeNo': 'Metro North',
    'MeSo': 'Metro South', 'NoWe': 'North West', 'SoWe': 'South West',
    'SuCo': 'Sunshine Coast', 'ToCa': 'Torres and Cape', 'Townsv': 'Townsville',
    'WeMo': 'West Moreton', 'WiBa': 'Wide Bay',
}

def _folder_prefix(hosp):
    """(output folder, short site name) for a hospital, honouring overrides."""
    prefix = hosp.split('_')[0] if '_' in hosp else 'Other'
    short  = hosp.split('_', 1)[-1] if '_' in hosp else hosp
    return OUTPUT_FOLDER_OVERRIDES.get(short.strip().lower(), prefix), short

def _per_serial_sum(df, hospital, col):
    out = {}
    if (df is not None and not df.empty and 'Device Name' in df.columns
            and 'Hospital Name' in df.columns and col in df.columns):
        for _, r in df[df['Hospital Name'] == hospital].iterrows():
            k = _device_serial(r['Device Name'])
            out[k] = out.get(k, 0) + _safe_int(r[col])
    return out

def _site_device_rows(hosp, short, df_use, sim_counts, ceramic_by_device):
    """One dict per analyser at a site: ward, serial, carts, errors, rate, sims, cer."""
    hu = df_use[df_use['Hospital Name'] == hosp]
    if hu.empty:
        return []
    sims = _per_serial_sum(sim_counts, hosp, 'SIM_Runs')
    cers = _per_serial_sum(ceramic_by_device, hosp, 'Ceramic_Count')
    loc_col = next((c for c in hu.columns if 'location' in str(c).lower()), None)
    rows = []
    for _, r in hu.iterrows():
        dev = str(r.get('Device Name', '') or '').strip()
        if not dev or dev.lower() == 'nan':
            continue
        did   = _device_serial(dev)
        carts = _safe_int(r.get('Total Carts'))
        res   = _safe_int(r.get('Total Res'))
        err   = max(carts - res, 0)
        rows.append({'ward': _device_title_label(short, dev, r.get(loc_col, '') if loc_col else ''),
                     'serial': did, 'carts': carts, 'errors': err,
                     'rate': round(err / carts * 100, 1) if carts else 0.0,
                     'sims': sims.get(did, 0), 'cer': cers.get(did, 0)})
    rows.sort(key=lambda d: (d['ward'].lower(), d['serial']))
    return rows

def _hospital_summary_row(hosp, df_use, sim_counts, ceramic_by_device):
    _, short = _folder_prefix(hosp)
    rows = _site_device_rows(hosp, short, df_use, sim_counts, ceramic_by_device)
    if not rows:
        return None
    carts = sum(r['carts'] for r in rows)
    err   = sum(r['errors'] for r in rows)
    return {'hospital': hosp, 'short': short, 'n': len(rows), 'carts': carts, 'errors': err,
            'rate': round(err / carts * 100, 1) if carts else 0.0,
            'cer_fail': [r for r in rows if r['cer'] < CERAMIC_REQUIRED],
            'sim_none': [r for r in rows if r['sims'] <= 0]}

def _titled_doc(kind, subtitle):
    """Blank PQ document with the standard title block ('i-Stat <kind>')."""
    doc  = Document(TEMPLATE_PATH)
    body = doc.element.body
    sect = body.find(qn('w:sectPr'))
    for child in list(body):
        if child != sect:
            body.remove(child)
    root = doc.element
    existing_bg = root.find(qn('w:background'))
    if existing_bg is not None:
        root.remove(existing_bg)
    bg = OxmlElement('w:background')
    bg.set(qn('w:color'), 'FFFFFF')
    bg.set(qn('w:themeColor'), 'background1')
    root.insert(0, bg)
    LOGO_INDENT = Inches(2127 / 1440)
    p = doc.add_paragraph(style='Title')
    p.paragraph_format.left_indent = LOGO_INDENT
    _run(p, 'Pathology Queensland', size=12, color=C_DARK_BLUE)
    p = doc.add_paragraph(style='Title')
    p.paragraph_format.left_indent = LOGO_INDENT
    _run(p, 'i', italic=True, size=20, color=C_DARK_BLUE)
    _run(p, f'-Stat {kind}', size=20, color=C_DARK_BLUE)
    p = doc.add_paragraph(style='Subtitle')
    p.paragraph_format.first_line_indent = LOGO_INDENT
    r = _run(p, subtitle, size=18, color=C_GREEN)
    try:
        r.style = doc.styles['DocSubTitle']
    except Exception:
        pass
    return doc

def _snapshot_table(doc, rows):
    """Indicator / Value / Status card; rows = (label, value, status, colour)."""
    tbl = doc.add_table(rows=0, cols=3)
    _add_borders(tbl)
    hdr = tbl.add_row().cells
    for c, l in zip(hdr, ['Indicator', 'Value', 'Status']):
        _hdr_cell(c, l, FILL_GREY_HDR)
    for label, value, status, colour in rows:
        c = tbl.add_row().cells
        _data_cell(c[0], label,  FILL_WHITE, C_BLUE, bold=True, size=9)
        _data_cell(c[1], value,  FILL_WHITE, C_BLUE, size=9)
        _data_cell(c[2], status, FILL_WHITE, colour or C_BLUE, bold=True, size=9)
    _set_col_widths(tbl, [2700, 5100, 2560])
    _cell_padding(tbl)
    return tbl

def _sim_cell_text(n):
    return (str(n), C_BLUE) if n > 0 else ('⚠ none', C_RED_FAIL)

def _cer_cell_text(n):
    if n >= CERAMIC_REQUIRED:
        return (f'✓ {n}', C_GREEN_PASS)
    return (f'✗ {n} of {CERAMIC_REQUIRED}', C_RED_FAIL)

def _write_site_overview(short, rows, out_path, report_month):
    doc = _titled_doc('Site Overview', f'{short} – {report_month}')
    carts = sum(r['carts'] for r in rows); err = sum(r['errors'] for r in rows)
    rate  = round(err / carts * 100, 1) if carts else 0.0
    label, lc = get_perf_label(rate)
    cer_fail = [r for r in rows if r['cer'] < CERAMIC_REQUIRED]
    sim_none = [r for r in rows if r['sims'] <= 0]

    _heading(doc, 'Site Snapshot')
    _snapshot_table(doc, [
        ('Analysers Reporting', f'{len(rows)} analysers across the site', '', None),
        ('Overall Error Rate', f'{rate}%  ({err} errors in {carts} cartridges)', label, lc),
        ('Simulator',
         'All analysers ran a simulator' if not sim_none
         else f'{len(sim_none)} of {len(rows)} analysers ran no simulator: '
              + ', '.join(r['ward'] for r in sim_none),
         '✓ OK' if not sim_none else '⚠ Check',
         C_GREEN_PASS if not sim_none else C_RED_FAIL),
        ('Ceramic Cleaning',
         f'All analysers ran {CERAMIC_REQUIRED}+ ceramic cartridges' if not cer_fail
         else f'{len(cer_fail)} of {len(rows)} analysers below {CERAMIC_REQUIRED}: '
              + ', '.join(f"{r['ward']} ({r['cer']})" for r in cer_fail),
         '✓ Pass' if not cer_fail else '✗ Fail',
         C_GREEN_PASS if not cer_fail else C_RED_FAIL),
    ])

    _heading(doc, f'Analysers — {report_month}')
    cols = ['Ward', 'Serial', 'Cartridges', 'Errors', 'Error %', 'Performance', 'Simulator', 'Ceramic']
    tbl = doc.add_table(rows=1, cols=len(cols))
    _add_borders(tbl)
    for c, l in zip(tbl.rows[0].cells, cols):
        _hdr_cell(c, l, FILL_BLUE_HDR, C_BLUE)
    C = WD_ALIGN_PARAGRAPH.CENTER
    for i, r in enumerate(rows):
        fill = FILL_ALT_ROW if i % 2 else FILL_WHITE
        pl, pc = get_perf_label(r['rate'])
        st, sc = _sim_cell_text(r['sims']); ct, cc = _cer_cell_text(r['cer'])
        c = tbl.add_row().cells
        _data_cell(c[0], r['ward'], fill, C_BLUE, size=8)
        _data_cell(c[1], r['serial'], fill, C_BLUE, size=8, align=C)
        _data_cell(c[2], str(r['carts']), fill, C_BLUE, size=8, align=C)
        _data_cell(c[3], str(r['errors']), fill, C_BLUE, size=8, align=C)
        _data_cell(c[4], f"{r['rate']}%", fill, pc, bold=True, size=8, align=C)
        _data_cell(c[5], pl, fill, pc, size=8, align=C)
        _data_cell(c[6], st, fill, sc, bold=(r['sims'] <= 0), size=8, align=C)
        _data_cell(c[7], ct, fill, cc, bold=True, size=8, align=C)
    c = tbl.add_row().cells
    _data_cell(c[0], 'Whole site', FILL_BLUE_HDR, C_BLUE, bold=True, size=8)
    _data_cell(c[1], '', FILL_BLUE_HDR)
    _data_cell(c[2], str(carts), FILL_BLUE_HDR, C_BLUE, bold=True, size=8, align=C)
    _data_cell(c[3], str(err), FILL_BLUE_HDR, C_BLUE, bold=True, size=8, align=C)
    _data_cell(c[4], f'{rate}%', FILL_BLUE_HDR, lc, bold=True, size=8, align=C)
    _data_cell(c[5], label, FILL_BLUE_HDR, lc, bold=True, size=8, align=C)
    _data_cell(c[6], f'{len(sim_none)} none' if sim_none else 'all run', FILL_BLUE_HDR,
               C_RED_FAIL if sim_none else C_GREEN_PASS, bold=True, size=8, align=C)
    _data_cell(c[7], f'{len(cer_fail)} fail' if cer_fail else 'all pass', FILL_BLUE_HDR,
               C_RED_FAIL if cer_fail else C_GREEN_PASS, bold=True, size=8, align=C)
    _set_col_widths(tbl, [2660, 900, 1000, 800, 850, 1350, 1000, 1800])
    _cell_padding(tbl, top=40, bottom=40)
    doc.save(out_path)
    _patch_white_background(out_path)

def _write_area_summary(prefix, hrows, out_path, report_month):
    area = HHS_NAMES.get(prefix, prefix)
    doc  = _titled_doc('Area Summary', f'{area} – {report_month}')
    n_an  = sum(h['n'] for h in hrows)
    carts = sum(h['carts'] for h in hrows); err = sum(h['errors'] for h in hrows)
    rate  = round(err / carts * 100, 1) if carts else 0.0
    label, lc = get_perf_label(rate)
    cer_h = [h for h in hrows if h['cer_fail']]
    sim_h = [h for h in hrows if h['sim_none']]
    n_cer = sum(len(h['cer_fail']) for h in hrows)
    n_sim = sum(len(h['sim_none']) for h in hrows)

    _heading(doc, 'Area Snapshot')
    _snapshot_table(doc, [
        ('Hospitals Reporting', f'{len(hrows)} hospitals, {n_an} analysers', '', None),
        ('Area Error Rate', f'{rate}%  ({err} errors in {carts} cartridges)', label, lc),
        ('Simulator',
         'Every analyser in the area ran a simulator' if not sim_h
         else f'{n_sim} analysers with no simulator runs at: ' + ', '.join(h['short'] for h in sim_h),
         '✓ OK' if not sim_h else '⚠ Check', C_GREEN_PASS if not sim_h else C_RED_FAIL),
        ('Ceramic Cleaning',
         f'Every analyser ran {CERAMIC_REQUIRED}+ ceramic cartridges' if not cer_h
         else f'{len(cer_h)} of {len(hrows)} hospitals failed ({n_cer} analysers): '
              + ', '.join(h['short'] for h in cer_h),
         '✓ Pass' if not cer_h else '✗ Fail', C_GREEN_PASS if not cer_h else C_RED_FAIL),
    ])

    _heading(doc, f'Hospitals — {report_month}')
    cols = ['Hospital', 'Analysers', 'Cartridges', 'Errors', 'Error %', 'Performance', 'Simulator', 'Ceramic']
    tbl = doc.add_table(rows=1, cols=len(cols))
    _add_borders(tbl)
    for c, l in zip(tbl.rows[0].cells, cols):
        _hdr_cell(c, l, FILL_BLUE_HDR, C_BLUE)
    C = WD_ALIGN_PARAGRAPH.CENTER
    for i, h in enumerate(hrows):
        fill = FILL_ALT_ROW if i % 2 else FILL_WHITE
        pl, pc = get_perf_label(h['rate'])
        ns, nc = len(h['sim_none']), len(h['cer_fail'])
        c = tbl.add_row().cells
        _data_cell(c[0], h['short'], fill, C_BLUE, bold=True, size=8)
        _data_cell(c[1], str(h['n']), fill, C_BLUE, size=8, align=C)
        _data_cell(c[2], str(h['carts']), fill, C_BLUE, size=8, align=C)
        _data_cell(c[3], str(h['errors']), fill, C_BLUE, size=8, align=C)
        _data_cell(c[4], f"{h['rate']}%", fill, pc, bold=True, size=8, align=C)
        _data_cell(c[5], pl, fill, pc, size=8, align=C)
        _data_cell(c[6], '✓ all run' if not ns else f'⚠ {ns} of {h["n"]} none', fill,
                   C_GREEN_PASS if not ns else C_RED_FAIL, bold=bool(ns), size=8, align=C)
        _data_cell(c[7], '✓ Pass' if not nc else f'✗ {nc} of {h["n"]} below {CERAMIC_REQUIRED}', fill,
                   C_GREEN_PASS if not nc else C_RED_FAIL, bold=True, size=8, align=C)
    c = tbl.add_row().cells
    _data_cell(c[0], 'Whole area', FILL_BLUE_HDR, C_BLUE, bold=True, size=8)
    _data_cell(c[1], str(n_an), FILL_BLUE_HDR, C_BLUE, bold=True, size=8, align=C)
    _data_cell(c[2], str(carts), FILL_BLUE_HDR, C_BLUE, bold=True, size=8, align=C)
    _data_cell(c[3], str(err), FILL_BLUE_HDR, C_BLUE, bold=True, size=8, align=C)
    _data_cell(c[4], f'{rate}%', FILL_BLUE_HDR, lc, bold=True, size=8, align=C)
    _data_cell(c[5], label, FILL_BLUE_HDR, lc, bold=True, size=8, align=C)
    _data_cell(c[6], f'{n_sim} none' if n_sim else 'all run', FILL_BLUE_HDR,
               C_RED_FAIL if n_sim else C_GREEN_PASS, bold=True, size=8, align=C)
    _data_cell(c[7], f'{n_cer} fail' if n_cer else 'all pass', FILL_BLUE_HDR,
               C_RED_FAIL if n_cer else C_GREEN_PASS, bold=True, size=8, align=C)
    _set_col_widths(tbl, [2300, 900, 1000, 800, 850, 1350, 1300, 1860])
    _cell_padding(tbl, top=40, bottom=40)

    _heading(doc, 'Analysers Needing Attention')
    issues = []
    for h in hrows:
        seen = {}
        for r in h['sim_none']:
            seen.setdefault(r['serial'], [r, []])[1].append('No simulator runs')
        for r in h['cer_fail']:
            seen.setdefault(r['serial'], [r, []])[1].append(f"Ceramic {r['cer']} of {CERAMIC_REQUIRED}")
        for r, why in seen.values():
            issues.append((h['short'], r['ward'], r['serial'], '; '.join(why)))
    if not issues:
        p = doc.add_paragraph()
        _run(p, 'Every analyser in the area ran a simulator and completed ceramic cleaning this month.',
             size=10, color=C_GREEN_PASS)
    else:
        cols = ['Hospital', 'Ward / Analyser', 'Serial', 'Issue']
        tbl = doc.add_table(rows=1, cols=len(cols))
        _add_borders(tbl)
        for c, l in zip(tbl.rows[0].cells, cols):
            _hdr_cell(c, l, FILL_BLUE_HDR, C_BLUE)
        for i, (hs, ward, serial, why) in enumerate(issues):
            fill = FILL_ALT_ROW if i % 2 else FILL_WHITE
            c = tbl.add_row().cells
            _data_cell(c[0], hs, fill, C_BLUE, size=8)
            _data_cell(c[1], ward, fill, C_BLUE, size=8)
            _data_cell(c[2], serial, fill, C_BLUE, size=8, align=C)
            _data_cell(c[3], why, fill, C_RED_FAIL, bold=True, size=8)
        _set_col_widths(tbl, [2300, 3400, 1200, 3460])
        _cell_padding(tbl, top=40, bottom=40)
    doc.save(out_path)
    _patch_white_background(out_path)


def safe_fn(name):
    return re.sub(r'[^\w\s\-]', '', name).strip().replace(' ', '_')


def _patch_white_background(docx_path):
    """Ensure Word renders a white page background by adding displayBackgroundShape to settings.xml."""
    import zipfile, shutil, tempfile
    tmp = docx_path + '.tmp'
    with zipfile.ZipFile(docx_path, 'r') as zin, \
         zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'word/settings.xml':
                xml = data.decode('utf-8')
                if 'displayBackgroundShape' not in xml:
                    # Insert before closing </w:settings>
                    xml = xml.replace(
                        '</w:settings>',
                        '<w:displayBackgroundShape/></w:settings>'
                    )
                data = xml.encode('utf-8')
            zout.writestr(item, data)
    shutil.move(tmp, docx_path)


# =============================================================================
#  BACKGROUND WORKER
# =============================================================================

def run_generation(excel_path, output_dir, report_month, report_month_num, report_year, selected=None):
    global progress
    progress = {"total": 0, "done": 0, "current": "Loading data…",
                "errors": [], "complete": False, "output": output_dir, "reports": 0}
    try:
        df_use, sim_counts, ceramic_counts, ceramic_by_device, df_err, pats_by_staff = load_data(excel_path)
        history       = _load_history()
        staff_history       = _load_staff_history()
        error_type_history  = _load_error_type_history()

        # ── Dormant analysers: record this month's activity, then drop any
        #    unit idle for DORMANT_MONTHS consecutive recorded months ──
        progress["current"] = "Checking for dormant analysers…"
        activity_history = _load_activity_history()
        _record_activity(activity_history, df_use, sim_counts, ceramic_by_device,
                         report_year, report_month_num)
        dormant_all   = _find_dormant(activity_history, report_year, report_month_num)
        _dormant_keys = {f"{d['hospital']}|{d['serial']}" for d in dormant_all}
        if _dormant_keys:
            df_use            = _drop_dormant(df_use, _dormant_keys)
            sim_counts        = _drop_dormant(sim_counts, _dormant_keys)
            ceramic_by_device = _drop_dormant(ceramic_by_device, _dormant_keys)
            print(f'[DORMANT] {len(_dormant_keys)} analyser(s) idle {DORMANT_MONTHS}+ months left out of this run')

        # Record ≥3-error staff for this month
        if not df_err.empty and 'Hospital Name' in df_err.columns:
            _sn = next((c for c in df_err.columns if 'surname'  in c.lower()), None)
            _fn = next((c for c in df_err.columns if 'first'    in c.lower()), None)
            if _sn and _fn:
                for (_hname, _sn_val, _fn_val), _g in df_err.groupby(['Hospital Name', _sn, _fn]):
                    _hname  = str(_hname)
                    _sn_val = str(_sn_val).strip()
                    _fn_val = str(_fn_val).strip()
                    _cnt    = len(_g)
                    _uid    = None
                    if 'User ID' in _g.columns:
                        _ids = _g['User ID'].dropna()
                        if len(_ids):
                            _uid = _ids.iloc[0]
                    # Successful cartridges (Pats + QC) from the permonthperuser sheet
                    _base = (
                        (pats_by_staff.get(('UID', _norm_uid(_uid))) if _uid is not None else None)
                        or pats_by_staff.get((_hname.upper(), _sn_val.upper(), _fn_val.upper()))
                        or pats_by_staff.get(('', _sn_val.upper(), _fn_val.upper()))
                        or 0
                    )
                    # total_carts = ATTEMPTS: successes + the failed cartridges
                    staff_history = _update_staff_history(
                        staff_history, _hname, _sn_val, _fn_val,
                        report_year, report_month_num, int(_cnt),
                        total_carts=_base + int(_cnt))


        # Per-device error-type counts for this month (for the volume chart)
        dev_type_counts = {}
        if not df_err.empty and 'DeviceID' in df_err.columns and 'Error Code' in df_err.columns:
            for _, _er in df_err.iterrows():
                _did  = _norm_did(_er.get('DeviceID'))
                _code = str(_er.get('Error Code') or '').strip().upper()
                if _did and _code:
                    dev_type_counts.setdefault(_did, {})
                    dev_type_counts[_did][_code] = dev_type_counts[_did].get(_code, 0) + 1

        # Record current month's error rate + volume for every device
        for _, row in df_use.iterrows():
            h = row.get('Hospital Name', '')
            d = row.get('Device Name', '')
            er = float(row.get('ErrorRate', 0) or 0)
            if h and d:
                _did = _extract_did(str(d))
                history = _update_history(history, str(h), str(d),
                                          report_year, report_month_num, er,
                                          carts=_safe_int(row.get('Total Carts')),
                                          types=dev_type_counts.get(_did, {}) if _did else {})

        hospitals = sorted(h for h in df_use['Hospital Name'].dropna().unique()
                           if str(h).upper() not in ('ALL', 'TOTAL', ''))
        if selected:
            hospitals = [h for h in hospitals if h in selected]
        progress["total"] = len(hospitals)
        os.makedirs(output_dir, exist_ok=True)

        for hosp in hospitals:
            progress["current"] = hosp
            try:
                prefix = hosp.split('_')[0] if '_' in hosp else 'Other'
                short  = hosp.split('_', 1)[-1] if '_' in hosp else hosp
                prefix = OUTPUT_FOLDER_OVERRIDES.get(short.strip().lower(), prefix)
                rdir   = os.path.join(output_dir, prefix)
                os.makedirs(rdir, exist_ok=True)

                # Determine whether this hospital gets one report per device.
                # PER_DEVICE_HOSPITALS (Townsville, Mt Isa, Redcliffe): split by device.
                # Mackay: only the main Mackay hospital (both HHS prefix AND
                #         short name contain "mackay") — other Mackay-region
                #         hospitals (Proserpine, Bowen, etc.) stay combined.
                _hosp_l  = hosp.lower()
                _short_l = short.lower()
                _is_per_device = (
                    any(k in _hosp_l for k in PER_DEVICE_HOSPITALS)
                    or (any(k in _hosp_l for k in ('mackay', 'mckay'))
                        and any(k in _short_l for k in ('mackay', 'mckay')))
                )
                if _is_per_device:
                    devices = sorted(
                        df_use[df_use['Hospital Name'] == hosp]['Device Name']
                        .dropna().unique()
                    )
                    for device in devices:
                        df_use_dev  = df_use[
                            (df_use['Hospital Name'] == hosp) &
                            (df_use['Device Name'] == device)
                        ].copy()
                        # ── Simulator: match by device ID, fall back to name ──
                        device_id_sim = _extract_did(str(device)) or str(device).strip()
                        def _dev_match(col_series):
                            return col_series.astype(str).apply(
                                lambda x: (_extract_did(x) or x.strip()) == device_id_sim
                            )
                        if 'Device Name' in sim_counts.columns:
                            sim_dev = sim_counts[
                                (sim_counts['Hospital Name'] == hosp) &
                                _dev_match(sim_counts['Device Name'])
                            ].copy()
                        else:
                            sim_dev = pd.DataFrame(columns=['Hospital Name', 'Device Name', 'SIM_Runs'])
                        # No fallback — empty means no data for this device

                        # ── Ceramic: match by device ID only, no fallback ──
                        if 'Device Name' in ceramic_by_device.columns:
                            cer_match = ceramic_by_device[
                                (ceramic_by_device['Hospital Name'] == hosp) &
                                _dev_match(ceramic_by_device['Device Name'])
                            ]
                            cer_dev = cer_match[['Hospital Name', 'Ceramic_Count']].copy() if not cer_match.empty else pd.DataFrame(columns=['Hospital Name', 'Ceramic_Count'])
                        else:
                            cer_dev = pd.DataFrame(columns=['Hospital Name', 'Ceramic_Count'])

                        # Filter errors to this device only.
                        # Strategy 1: match on Device Name column in error sheet.
                        # Strategy 2: match on Location — get locations from usage rows
                        #             for this device, then keep matching error rows.
                        df_err_dev = (df_err[df_err['Hospital Name'] == hosp].copy()
                                      if 'Hospital Name' in df_err.columns else df_err.copy())
                        dev_err_col = next((c for c in df_err_dev.columns
                                            if 'device' in c.lower()), None)
                        if dev_err_col:
                            device_id = _extract_did(str(device)) or str(device)
                            df_err_dev = df_err_dev[
                                df_err_dev[dev_err_col].astype(str)
                                .apply(lambda x: (_extract_did(x) or x.strip()) == device_id)
                            ]
                        else:
                            # Fall back to location matching
                            loc_use_col = next((c for c in df_use_dev.columns
                                                if 'location' in c.lower()), None)
                            loc_err_col = next((c for c in df_err_dev.columns
                                                if 'location' in c.lower()), None)
                            if loc_use_col and loc_err_col:
                                dev_locs = set(
                                    df_use_dev[loc_use_col].dropna()
                                    .astype(str).str.strip().unique()
                                )
                                df_err_dev = df_err_dev[
                                    df_err_dev[loc_err_col].astype(str)
                                    .str.strip().isin(dev_locs)
                                ]

                        _loc_col = next((c for c in df_use_dev.columns
                                         if 'location' in str(c).lower()), None)
                        _loc_vals = (df_use_dev[_loc_col].dropna().astype(str)
                                     if _loc_col else [])
                        _ward = _device_title_label(
                            short, device, _loc_vals.iloc[0] if len(_loc_vals) else '')
                        doc = generate_report(hosp, df_use_dev, sim_dev, cer_dev, df_err_dev,
                                              report_month, report_month_num, report_year,
                                              history=history, staff_history=staff_history,
                                              error_type_history=error_type_history,
                                              pats_by_staff=pats_by_staff,
                                              ceramic_by_device=ceramic_by_device,
                                              title_label=_ward)
                        if doc:
                            did   = _extract_did(str(device)) or safe_fn(str(device))
                            fname = f"i-STAT_{safe_fn(_ward)}_{did}_{report_month.replace(' ', '')}.docx"
                            out_path = os.path.join(rdir, fname)
                            doc.save(out_path)
                            _patch_white_background(out_path)
                            progress["reports"] += 1
                    # ── Site overview: every ward on one page for the site manager ──
                    _ov_rows = _site_device_rows(hosp, short, df_use, sim_counts, ceramic_by_device)
                    if _ov_rows:
                        _ov_path = os.path.join(
                            rdir, f"i-STAT_{safe_fn(short)}_Site_Overview_{report_month.replace(' ', '')}.docx")
                        _write_site_overview(short, _ov_rows, _ov_path, report_month)
                        progress["reports"] += 1
                else:
                    doc = generate_report(hosp, df_use, sim_counts, ceramic_counts, df_err,
                                          report_month, report_month_num, report_year,
                                          history=history, staff_history=staff_history,
                                          error_type_history=error_type_history,
                                          pats_by_staff=pats_by_staff,
                                              ceramic_by_device=ceramic_by_device)
                    if doc:
                        fname = f"i-STAT_{safe_fn(short)}_{report_month.replace(' ', '')}.docx"
                        out_path = os.path.join(rdir, fname)
                        doc.save(out_path)
                        _patch_white_background(out_path)
                        progress["reports"] += 1

            except Exception as e:
                import traceback
                progress["errors"].append(f"{hosp}: {traceback.format_exc(limit=2)}")
            progress["done"] += 1

        # ── Area summaries: one page per HHS folder touched by this run,
        #    covering every hospital in that area (whether selected or not) ──
        progress["current"] = "Area summaries…"
        _area_members = {}
        for _h in df_use['Hospital Name'].dropna().unique():
            if str(_h).upper() in ('ALL', 'TOTAL', ''):
                continue
            _area_members.setdefault(_folder_prefix(str(_h))[0], []).append(str(_h))
        for _pre in sorted({_folder_prefix(h)[0] for h in hospitals}):
            try:
                _hrows = [_hospital_summary_row(_h, df_use, sim_counts, ceramic_by_device)
                          for _h in sorted(_area_members.get(_pre, []))]
                _hrows = sorted((r for r in _hrows if r), key=lambda r: r['short'].lower())
                if _hrows:
                    _adir = os.path.join(output_dir, _pre)
                    os.makedirs(_adir, exist_ok=True)
                    _write_area_summary(_pre, _hrows, os.path.join(
                        _adir, f"i-STAT_{safe_fn(_pre)}_Area_Summary_{report_month.replace(' ', '')}.docx"),
                        report_month)
                    progress["reports"] += 1
            except Exception:
                import traceback
                progress["errors"].append(f"Area summary {_pre}: {traceback.format_exc(limit=2)}")

        _save_history(history)
        _save_staff_history(staff_history)
        _save_error_type_history(error_type_history)
        _save_activity_history(activity_history)

        # Companion list of the analysers left out of this run (selected sites only)
        _dormant_sel = [d for d in dormant_all if d['hospital'] in set(hospitals)]
        progress["dormant"] = len(_dormant_sel)
        if _dormant_sel:
            try:
                _dfile = f"Dormant_iSTATs_{report_month.replace(' ', '')}.docx"
                _write_dormant_report(_dormant_sel, os.path.join(output_dir, _dfile), report_month)
                progress["dormant_file"] = _dfile
                progress["reports"] += 1
            except Exception:
                import traceback
                progress["errors"].append(f"Dormant list: {traceback.format_exc(limit=2)}")

    except Exception as e:
        import traceback
        progress["errors"].append(f"Fatal: {traceback.format_exc(limit=3)}")
    finally:
        progress["complete"] = True
        progress["current"]  = "Done"
    try:
        os.remove(excel_path)
    except Exception:
        pass


# =============================================================================
#  FLASK ROUTES
# =============================================================================

@app.route('/preview', methods=['POST'])
def preview():
    if not _same_origin(request):
        abort(403)
    if 'file' not in request.files:
        return jsonify({"error": "No file"}), 400
    f = request.files['file']
    # Randomised name, OS temp dir, deleted in finally — no data remanence.
    fd, tmp = tempfile.mkstemp(suffix='.xlsx', prefix='istat_preview_')
    os.close(fd)
    try:
        f.save(tmp)
        wb   = openpyxl.load_workbook(tmp, data_only=True)
        ws   = _find_sheet(wb, _USE_KEYS, ['sim', 'error', 'event', 'user'])
        df   = _load_df(ws)
        names = sorted(
            h for h in df.get('Hospital Name', pd.Series(dtype=str)).dropna().unique()
            if str(h).upper() not in ('ALL', 'TOTAL', '')
        )
        hospitals = []
        for h in names:
            hhs   = h.split('_')[0]  if '_' in h else 'Other'
            short = h.split('_', 1)[-1] if '_' in h else h
            hospitals.append({"id": h, "hhs": hhs, "name": short})
        return jsonify({"hospitals": hospitals})
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/lock_status')
def lock_status():
    return jsonify({'initialized': _encryption_initialized(),
                    'unlocked': _unlocked()})

@app.route('/unlock', methods=['POST'])
def unlock():
    if not _same_origin(request):
        abort(403)
    pw = (request.form.get('password') or '').strip()
    if not pw:
        return jsonify({'ok': False, 'error': 'Enter a password.'})
    if not _encryption_initialized():
        if len(pw) < 8:
            return jsonify({'ok': False, 'error': 'Use at least 8 characters.'})
        if pw != (request.form.get('confirm') or '').strip():
            return jsonify({'ok': False, 'error': 'Passwords do not match.'})
        _set_history_password(pw)
        _audit_event('history-encryption', result='enabled, files encrypted')
        return jsonify({'ok': True})
    if _try_unlock(pw):
        _audit_event('history-unlock', result='success')
        return jsonify({'ok': True})
    _audit_event('history-unlock', result='FAILED attempt')
    return jsonify({'ok': False, 'error': 'Incorrect password.'})

@app.route('/generate', methods=['POST'])
def generate():
    global progress
    if _encryption_initialized() and not _unlocked():
        return jsonify({'error': 'History is locked — enter the password first.'}), 403
    if not _same_origin(request):
        abort(403)
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    f = request.files['file']
    if not f.filename.endswith('.xlsx'):
        return jsonify({"error": "Please upload an .xlsx file"}), 400
    if not os.path.exists(TEMPLATE_PATH):
        return jsonify({"error": "template.docx not found in app folder"}), 500

    month_name   = request.form.get('month_name', 'April 2025')
    month_num    = request.form.get('month_num',  '04')
    year         = request.form.get('year',       '2025')
    sel_json     = request.form.get('selected_hospitals', '')
    selected     = json.loads(sel_json) if sel_json else None

    # Randomised name so concurrent runs can't clobber each other; removed by
    # run_generation when processing completes.
    fd, upload_path = tempfile.mkstemp(suffix='.xlsx', prefix='istat_upload_')
    os.close(fd)
    f.save(upload_path)

    _audit_event('generate', period=month_name,
                 hospitals=(len(selected) if selected else 'all'))

    output_dir = _output_dir_for(month_name, month_num, year)
    os.makedirs(output_dir, exist_ok=True)
    progress   = {"total": 0, "done": 0, "current": "Starting...",
                  "errors": [], "complete": False, "output": output_dir, "reports": 0}
    t = threading.Thread(
        target=run_generation,
        args=(upload_path, output_dir, month_name, month_num, year, selected),
        daemon=True
    )
    t.start()
    return jsonify({"status": "started", "output": output_dir})

@app.route('/progress')
def get_progress():
    return jsonify(progress)

@app.route('/open_folder')
def open_folder():
    folder = request.args.get('path', '')
    if not folder:
        return jsonify({"error": "Folder not found"}), 404
    real = os.path.realpath(folder)
    if not real.startswith(os.path.realpath(OUTPUT_ROOT)):
        return jsonify({"error": "Access denied"}), 403
    if sys.platform == 'win32':
        os.startfile(real)
    elif sys.platform == 'darwin':
        subprocess.Popen(['open', real])
    else:
        subprocess.Popen(['xdg-open', real])
    return jsonify({"status": "ok"})

if __name__ == '__main__':
    app.run(debug=False, port=PORT, threaded=True)
