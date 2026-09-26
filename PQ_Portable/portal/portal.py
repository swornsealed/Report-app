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

# PC enrolment (bundle root). The portal is where a PC is enrolled or renewed;
# both engines check the same registry and lock themselves otherwise.
sys.path.insert(0, ROOT)
import pq_enrolment as enrol
enrol.record_launch('portal')

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

@app.route('/enrol_status')
def enrol_status():
    return jsonify(enrol.status())

@app.route('/enrol', methods=['POST'])
def do_enrol():
    if not _same_origin(request):
        abort(403)
    data = request.get_json(silent=True) or {}
    pw   = str(data.get('password') or '')
    try:
        if not enrol.is_initialised():
            if pw != str(data.get('password2') or ''):
                return jsonify({'error': 'The two passwords do not match.'}), 400
            enrol.setup(pw)
        entry = enrol.enrol(pw, note=data.get('note', ''),
                            confirm_network=bool(data.get('confirm_network')))
    except ValueError as exc:
        if str(exc) == 'NETWORK_CONFIRM':
            return jsonify({'error': 'network_confirm', 'network': enrol.network_state()}), 409
        return jsonify({'error': str(exc)}), 400
    return jsonify({'ok': True, 'expires_at': entry['expires_at']})

@app.route('/enrolled')
def enrolled():
    return render_template('enrolled.html', entries=enrol.entries(), me=enrol.machine_id(),
                           version=enrol.version(), term=enrol.TERM_MONTHS)

@app.route('/revoke', methods=['POST'])
def do_revoke():
    if not _same_origin(request):
        abort(403)
    data = request.get_json(silent=True) or {}
    try:
        n = enrol.revoke(str(data.get('machine_id') or ''), str(data.get('password') or ''))
    except ValueError as exc:
        return jsonify({'error': str(exc)}), 400
    return jsonify({'ok': True, 'removed': n})

if __name__ == '__main__':
    print('\n  Pathology Queensland — Operator Report Generator')
    print(f'  Portal running at http://localhost:{PORT}')
    _start_engines()
    print('  Close this window to stop everything.\n')
    app.run(host='127.0.0.1', port=PORT, debug=False)
