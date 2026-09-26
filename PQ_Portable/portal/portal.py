"""Pathology Queensland — Operator Report Generator portal.

Single entry point for the combined portable bundle: starts both report
engines (i-STAT on :5757, ABL on :5758) as child processes and serves the
selection page on :5750. Closing the console window stops everything.
Loopback-only, offline — same posture as the report engines themselves.
"""
import os
import atexit
import subprocess
from flask import Flask, render_template, jsonify

APP_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT    = os.path.dirname(APP_DIR)
PYTHON  = os.path.join(ROOT, 'python', 'python.exe')
PORT    = 5750

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

if __name__ == '__main__':
    print('\n  Pathology Queensland — Operator Report Generator')
    print(f'  Portal running at http://localhost:{PORT}')
    _start_engines()
    print('  Close this window to stop everything.\n')
    app.run(host='127.0.0.1', port=PORT, debug=False)
