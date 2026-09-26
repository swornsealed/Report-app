"""Pathology Queensland — Operator Report Generator portal.

Single entry point for the combined portable bundle: starts both report
engines (i-STAT on :5757, ABL on :5758) as child processes and serves the
selection page on :5750. Closing the console window stops everything.
Loopback-only, offline — same posture as the report engines themselves.
"""
import os
import sys
import atexit
import subprocess
from flask import Flask, render_template, jsonify, request, abort

APP_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT    = os.path.dirname(APP_DIR)
PYTHON  = os.path.join(ROOT, 'python', 'python.exe')
PORT    = 5750

# Licence check (bundle root). The portal shows the licence status, installs a
# licence file from the PoC lead and lists the licences on the drive; both
# engines check the same files and lock themselves otherwise.
sys.path.insert(0, ROOT)
import pq_licence as lic
lic.record_launch('portal')

def _same_origin(req):
    src = req.headers.get('Origin') or req.headers.get('Referer') or ''
    return src.startswith(f'http://localhost:{PORT}') or src.startswith(f'http://127.0.0.1:{PORT}')

ENGINES = [
    {'key': 'istat', 'name': 'i-STAT — Point of Care', 'port': 5757,
     'script': os.path.join(ROOT, 'iSTAT_App', 'app.py')},
    {'key': 'abl',   'name': 'ABL — Blood Gas',        'port': 5758,
     'script': os.path.join(ROOT, 'ABL_App', 'app.py')},
]

_children = []

def _start_engines():
    env = dict(os.environ, PQ_PORTAL='1')
    for e in ENGINES:
        if os.path.exists(e['script']):
            _children.append(subprocess.Popen([PYTHON, e['script']], env=env))
            print(f"  started {e['name']}  ->  http://localhost:{e['port']}")
        else:
            print(f"  WARNING: {e['script']} not found — {e['name']} unavailable")

def _stop_engines():
    for p in _children:
        try:
            p.terminate()
        except Exception:
            pass

atexit.register(_stop_engines)

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/engines')
def engines():
    return jsonify([{'key': e['key'], 'name': e['name'], 'port': e['port']}
                    for e in ENGINES])

@app.route('/licence_status')
def licence_status():
    return jsonify(lic.status())

@app.route('/licence_install', methods=['POST'])
def licence_install():
    if not _same_origin(request):
        abort(403)
    f = request.files.get('file')
    if f is None:
        return jsonify({'error': 'Choose the .lic file you received from the PoC lead.'}), 400
    raw = f.read(65536)
    try:
        pl = lic.install_licence(raw, f.filename or 'licence.lic')
    except ValueError as exc:
        lic._audit('licence_install_refused', reason=str(exc))
        return jsonify({'error': str(exc)}), 400
    st = lic.status()
    return jsonify({'ok': True, 'site': pl['site'], 'expires': pl['expires'],
                    'covers_this_pc': st['allowed'], 'message': st['message']})

@app.route('/licences')
def licences_page():
    return render_template('licences.html', licences=lic.licences(), me=lic.machine_id(),
                           version=lic.version(), status=lic.status())

if __name__ == '__main__':
    print('\n  Pathology Queensland — Operator Report Generator')
    print(f'  Portal running at http://localhost:{PORT}')
    _start_engines()
    print('  Close this window to stop everything.\n')
    app.run(host='127.0.0.1', port=PORT, debug=False)
