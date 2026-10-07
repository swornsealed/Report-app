"""bundle_licence.py — licence check for the PQ Operator Report Generator bundle.

Central control. A PC runs the bundle only under a LICENCE FILE (*.lic at the
drive root) issued by the PoC lead's offline keygen (Licensing\\bundle_keygen.py,
never on the drive). The drive itself cannot authorise anything: it carries
only the PUBLIC key below, so licence files cannot be forged from the drive.

A licence is a signed JSON document naming:
  site          who it was issued to
  machines      the machine IDs allowed to run (id + label), a seat cap
  issued / expires   a 12-month term matching the QIS review cycle
  max_version   the newest bundle version it covers — a newer bundle needs a
                new licence, so re-issue on review = controlled roll-out
The machine ID is a property of the PC (SHA-256 of system UUID, baseboard
serial, system-drive serial and hostname): a copy of the drive on another PC
yields a different ID and stays locked until that ID is licensed.

States, from the term end:
  LICENSED   more than 30 days left        full function
  EXPIRING   30 days or fewer              full function, renewal notice
  GRACE      up to 30 days past the end    full function, prominent notice
  EXPIRED    after the grace period        locked
  LOCKED     no licence covers this PC / bundle newer than the licence allows
A system clock earlier than the last recorded run also locks (dates cannot be
trusted) until the clock is corrected.

Files beside this module (drive root):
  VERSION            bundle version
  *.lic              licence files (signed; edits break the signature)
  licence_seen.json  informational: last launch per machine
  licence_audit.log  every launch, licence install, refusal, invalid file

Stated limit: the check and the public key live on the drive. Someone holding
the BitLocker password and able to edit Python could remove them. It is an
administrative control with an audit trail, not tamper-proof DRM.
"""
import os
import re
import sys
import json
import glob
import ctypes
import getpass
import hashlib
import socket
import subprocess
import datetime as _dt

ROOT         = os.path.dirname(os.path.abspath(__file__))
VERSION_FILE = os.path.join(ROOT, 'VERSION')
SEEN         = os.path.join(ROOT, 'licence_seen.json')
AUDIT        = os.path.join(ROOT, 'licence_audit.log')

# The PoC lead's Ed25519 PUBLIC key (hex). The private half lives only in
# Licensing\private_key.hex on the issuing PC. Replace both together.
PUBLIC_KEY_HEX = 'c81c4a8c04a76f1e1205f90cad10f69cde1ae9b377bf5447e97440e9ba97e07c'

NOTICE_DAYS = 30
GRACE_DAYS  = 30


# ── version ──────────────────────────────────────────────────────────────────
def version():
    try:
        with open(VERSION_FILE, encoding='utf-8') as fh:
            return fh.read().strip() or '0'
    except Exception:
        return '0'

def _vtuple(v):
    return tuple(int(x) for x in re.findall(r'\d+', str(v))) or (0,)


# ── audit ────────────────────────────────────────────────────────────────────
def _user():
    try:
        return getpass.getuser()
    except Exception:
        return '?'

def _audit(action, **fields):
    line = ' '.join([_dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), action,
                     f'user={_user()}', f'machine={machine_id()}', f'version={version()}'] +
                    [f'{k}={v}' for k, v in fields.items()])
    try:
        with open(AUDIT, 'a', encoding='utf-8') as fh:
            fh.write(line + '\n')
    except Exception:
        pass


# ── machine identity ─────────────────────────────────────────────────────────
_PS = ("$u=(Get-CimInstance Win32_ComputerSystemProduct -ErrorAction SilentlyContinue).UUID;"
       "$b=(Get-CimInstance Win32_BaseBoard -ErrorAction SilentlyContinue).SerialNumber;"
       "$a=@(Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object Status -eq 'Up' | ForEach-Object Name);"
       "$i=@(Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | "
       "Where-Object { $_.IPAddress -notlike '127.*' -and $_.IPAddress -notlike '169.254.*' } | ForEach-Object IPAddress);"
       "@{uuid=\"$u\"; board=\"$b\"; adapters=@($a); ips=@($i)} | ConvertTo-Json -Compress")

def _windows_facts():
    facts = {'uuid': '', 'board': '', 'adapters': [], 'ips': []}
    if sys.platform != 'win32':
        return facts
    try:
        out = subprocess.run(['powershell', '-NoProfile', '-NonInteractive', '-Command', _PS],
                             capture_output=True, text=True, timeout=40,
                             creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0)).stdout.strip()
        d = json.loads(out) if out else {}
        u = str(d.get('uuid') or '').strip().upper()
        if u in ('', 'FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF', '00000000-0000-0000-0000-000000000000'):
            u = ''
        facts['uuid']     = u
        facts['board']    = str(d.get('board') or '').strip()
        facts['adapters'] = [str(x) for x in (d.get('adapters') or [])]
        facts['ips']      = [str(x) for x in (d.get('ips') or [])]
    except Exception:
        pass
    return facts

def _volume_serial():
    try:
        drive = os.environ.get('SystemDrive', 'C:') + '\\'
        serial = ctypes.c_uint32(0)
        ctypes.windll.kernel32.GetVolumeInformationW(
            ctypes.c_wchar_p(drive), None, 0, ctypes.byref(serial), None, None, None, 0)
        return f'{serial.value:08X}'
    except Exception:
        return ''

# Computed once per process (the WMI call is slow; never repeat it in request threads)
_FACTS    = _windows_facts()
_HOSTNAME = socket.gethostname()
_DIGEST   = hashlib.sha256('|'.join((_FACTS['uuid'], _FACTS['board'], _volume_serial(), _HOSTNAME))
                           .encode('utf-8', 'replace')).hexdigest().upper()
_MACHINE_ID = f'{_DIGEST[:4]}-{_DIGEST[4:8]}-{_DIGEST[8:12]}'

def machine_id():   return _MACHINE_ID
def hostname():     return _HOSTNAME
def identity_ok():  return bool(_FACTS['uuid'])
def network_state():
    return {'adapters': list(_FACTS['adapters']), 'ips': list(_FACTS['ips']),
            'connected': bool(_FACTS['adapters']) or bool(_FACTS['ips'])}


# ── licence files ────────────────────────────────────────────────────────────
def canonical(payload):
    return json.dumps(payload, sort_keys=True, separators=(',', ':')).encode('utf-8')

def _public_key():
    from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey
    return Ed25519PublicKey.from_public_bytes(bytes.fromhex(PUBLIC_KEY_HEX))

def parse_licence_bytes(raw):
    """Parse and verify one licence document. Raises ValueError when invalid."""
    try:
        doc = json.loads(raw.decode('utf-8-sig'))
        payload, sig = doc['payload'], bytes.fromhex(doc['sig'])
    except Exception:
        raise ValueError('Not a licence file.')
    try:
        _public_key().verify(sig, canonical(payload))
    except Exception:
        raise ValueError('Licence signature is invalid — the file was edited or issued for another key.')
    for k in ('licence_id', 'site', 'issued', 'expires', 'machines', 'seats'):
        if k not in payload:
            raise ValueError(f'Licence is missing "{k}".')
    if len(payload['machines']) > int(payload['seats']):
        raise ValueError('Licence lists more machines than its seat cap.')
    return payload

def _licence_files():
    return sorted(glob.glob(os.path.join(ROOT, '*.lic')))

_reported_bad = set()

def licences():
    """Every licence file at the root, verified, with per-machine last-seen."""
    seen  = _read_json(SEEN, {})
    today = _dt.date.today()
    out = []
    for path in _licence_files():
        item = {'file': os.path.basename(path), 'valid': False, 'error': ''}
        try:
            with open(path, 'rb') as fh:
                pl = parse_licence_bytes(fh.read())
            item.update(pl); item['valid'] = True
            exp = _dt.date.fromisoformat(pl['expires'])
            item['days_left'] = (exp - today).days
            item['state'] = _term_state(exp, today)[0]
            item['covers_this_pc'] = any(m.get('id') == machine_id() for m in pl['machines'])
            item['version_ok'] = _vtuple(version()) <= _vtuple(pl.get('max_version', '9999'))
            for m in item['machines']:
                m['last_seen'] = seen.get(m.get('id'), {})
        except ValueError as exc:
            item['error'] = str(exc)
            if path not in _reported_bad:
                _reported_bad.add(path)
                _audit('licence_file_invalid', file=item['file'], reason=str(exc))
        out.append(item)
    return out

def install_licence(raw, filename='licence.lic'):
    """Verify an uploaded licence and save it at the root (named by its id)."""
    pl = parse_licence_bytes(raw)
    name = re.sub(r'[^A-Za-z0-9_.-]', '_', str(pl['licence_id']))[:60] or 'licence'
    path = os.path.join(ROOT, f'{name}.lic')
    with open(path, 'wb') as fh:
        fh.write(raw)
    _audit('licence_installed', file=os.path.basename(path), site=pl['site'],
           expires=pl['expires'], machines=len(pl['machines']),
           covers_this_pc=any(m.get('id') == machine_id() for m in pl['machines']))
    return pl


# ── status ───────────────────────────────────────────────────────────────────
def _term_state(exp, today):
    left = (exp - today).days
    if left > NOTICE_DAYS:            return 'LICENSED', True
    if left >= 0:                     return 'EXPIRING', True
    if -left <= GRACE_DAYS:           return 'GRACE', True
    return 'EXPIRED', False

def _read_json(path, default):
    try:
        with open(path, encoding='utf-8') as fh:
            return json.load(fh)
    except Exception:
        return default

def _write_json(path, obj):
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as fh:
        json.dump(obj, fh, indent=2)
    os.replace(tmp, path)

def status(today=None):
    today = today or _dt.date.today()
    st = {'machine_id': machine_id(), 'hostname': hostname(), 'version': version(),
          'identity_ok': identity_ok(), 'network': network_state(),
          'allowed': False, 'state': 'LOCKED', 'site': None, 'licence_id': None,
          'expires': None, 'days_left': None, 'max_version': None, 'message': ''}
    seen = _read_json(SEEN, {}).get(machine_id(), {})
    last = seen.get('last_seen', '')
    if last and today.isoformat() < last[:10]:
        st['state'] = 'CLOCK'
        st['message'] = (f'This PC\'s clock ({today.isoformat()}) is earlier than its last recorded run '
                         f'({last[:10]}). Correct the clock to continue.')
        return st
    mine = [l for l in licences() if l['valid'] and l['covers_this_pc']]
    if not mine:
        n = len(_licence_files())
        st['message'] = ('No licence on this drive covers this PC. Send the machine ID to the PoC lead '
                         'to obtain a licence file.' if n else
                         'No licence file on this drive. Send the machine ID to the PoC lead to obtain one.')
        return st
    usable = [l for l in mine if l['version_ok']]
    if not usable:
        cap = max(mine, key=lambda l: _vtuple(l.get('max_version', '0')))
        st['state'] = 'VERSION'
        st['message'] = (f'The licence covers bundle versions up to {cap.get("max_version")}; this bundle is '
                         f'{version()}. A licence issued for this version is needed.')
        return st
    best = max(usable, key=lambda l: l['expires'])
    exp  = _dt.date.fromisoformat(best['expires'])
    st.update({'site': best['site'], 'licence_id': best['licence_id'], 'expires': best['expires'],
               'days_left': (exp - today).days, 'max_version': best.get('max_version')})
    st['state'], st['allowed'] = _term_state(exp, today)
    when = exp.strftime('%d %b %Y')
    st['message'] = {
        'LICENSED': f'Licensed to {best["site"]} until {when}.',
        'EXPIRING': f'Licence for {best["site"]} expires on {when} ({st["days_left"]} days) — ask the PoC lead for the renewed licence.',
        'GRACE':    f'Licence for {best["site"]} expired on {when}. It keeps working for {GRACE_DAYS + st["days_left"]} more days — install the renewed licence now.',
        'EXPIRED':  f'Licence for {best["site"]} expired on {when} and the grace period has ended. Install the renewed licence file.',
    }[st['state']]
    return st

def record_launch(component):
    st = status()
    seen = _read_json(SEEN, {})
    now = _dt.datetime.now().strftime('%Y-%m-%d %H:%M')
    if st['state'] != 'CLOCK':
        seen[machine_id()] = {'last_seen': now, 'hostname': hostname(), 'user': _user(),
                              'version': version(), 'component': component}
        try:
            _write_json(SEEN, seen)
        except Exception:
            pass
    _audit('launch', component=component, state=st['state'], site=st['site'] or '-',
           network=st['network']['connected'])
    return st
