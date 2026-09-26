"""pq_enrolment.py — PC enrolment for the PQ Operator Report Generator bundle.

The bundle runs only on PCs that have been enrolled with the PoC team's
enrolment password. Enrolment is bound to the machine itself (motherboard /
system UUID, baseboard serial, system-drive serial and hostname, hashed to a
short ID), so copying the drive to another PC yields a different ID and the
copy stays locked until that PC is enrolled. Every enrolment carries a term
(TERM_MONTHS) so the team knows which PCs run which version and can bring
them back for updates; an expired PC is locked until it is re-enrolled.

Files (all beside this module, at the bundle root):
  VERSION               bundle version string, recorded with every enrolment
  enrolment.keymeta     salt + Ed25519 public key + private key encrypted with
                        the enrolment password (PBKDF2-SHA256, 600 000 rounds)
  enrolled_pcs.json     one signed entry per enrolment (Ed25519, verified with
                        the public key on every launch — an edited entry fails)
  enrolment_seen.json   informational: last launch per machine
  enrolment_audit.log   every launch check, enrolment, renewal, revocation and
                        failed password, with Windows user and machine ID

Stated limit: the check, the public key and the registry all live on the
drive. Someone holding the BitLocker password and able to edit Python could
remove it. It is an administrative control with an audit trail, not DRM.
"""
import os
import re
import sys
import json
import base64
import ctypes
import getpass
import hashlib
import socket
import subprocess
import datetime as _dt

ROOT         = os.path.dirname(os.path.abspath(__file__))
VERSION_FILE = os.path.join(ROOT, 'VERSION')
KEYMETA      = os.path.join(ROOT, 'enrolment.keymeta')
REGISTRY     = os.path.join(ROOT, 'enrolled_pcs.json')
SEEN         = os.path.join(ROOT, 'enrolment_seen.json')
AUDIT        = os.path.join(ROOT, 'enrolment_audit.log')

TERM_MONTHS   = 12      # an enrolment lasts this long, then the PC must be re-enrolled
NOTICE_DAYS   = 30      # renewal notice this many days before the end
PBKDF2_ROUNDS = 600_000


# ── version ──────────────────────────────────────────────────────────────────
def version():
    try:
        with open(VERSION_FILE, encoding='utf-8') as fh:
            return fh.read().strip() or 'unknown'
    except Exception:
        return 'unknown'


# ── audit ────────────────────────────────────────────────────────────────────
def _audit(action, **fields):
    try:
        user = getpass.getuser()
    except Exception:
        user = '?'
    line = ' '.join([_dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), action,
                     f'user={user}', f'machine={machine_id()}', f'version={version()}'] +
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
    """System UUID, baseboard serial and live network state via one PowerShell
    call (no extra packages, no sockets opened by this process)."""
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

# Computed once per process — the WMI call is slow and must not be repeated
# inside request threads.
_FACTS    = _windows_facts()
_HOSTNAME = socket.gethostname()
_BASIS    = '|'.join((_FACTS['uuid'], _FACTS['board'], _volume_serial(), _HOSTNAME))
_DIGEST   = hashlib.sha256(_BASIS.encode('utf-8', 'replace')).hexdigest().upper()
_MACHINE_ID = f'{_DIGEST[:4]}-{_DIGEST[4:8]}-{_DIGEST[8:12]}'

def machine_id():
    return _MACHINE_ID

def hostname():
    return _HOSTNAME

def identity_ok():
    """False when the motherboard identifier could not be read — no enrolment then."""
    return bool(_FACTS['uuid'])

def network_state():
    return {'adapters': list(_FACTS['adapters']), 'ips': list(_FACTS['ips']),
            'connected': bool(_FACTS['adapters']) or bool(_FACTS['ips'])}


# ── crypto helpers ───────────────────────────────────────────────────────────
def _crypto():
    from cryptography.fernet import Fernet, InvalidToken
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives.asymmetric.ed25519 import (
        Ed25519PrivateKey, Ed25519PublicKey)
    from cryptography.exceptions import InvalidSignature
    return (Fernet, InvalidToken, hashes, serialization, PBKDF2HMAC,
            Ed25519PrivateKey, Ed25519PublicKey, InvalidSignature)

def _derive(password, salt):
    Fernet, _, hashes, _, PBKDF2HMAC, *_ = _crypto()
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=PBKDF2_ROUNDS)
    return base64.urlsafe_b64encode(kdf.derive(password.encode('utf-8')))

def _b64(b):  return base64.b64encode(b).decode('ascii')
def _unb64(s): return base64.b64decode(s.encode('ascii'))

def _canonical(entry):
    return json.dumps({k: v for k, v in entry.items() if k != 'sig'},
                      sort_keys=True, separators=(',', ':')).encode('utf-8')

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


# ── set-up and password ──────────────────────────────────────────────────────
def is_initialised():
    return os.path.exists(KEYMETA)

def setup(password):
    """First use: choose the enrolment password and create the signing key pair."""
    if is_initialised():
        raise ValueError('Enrolment is already set up on this drive.')
    if len(password or '') < 8:
        raise ValueError('Use at least 8 characters.')
    Fernet, _, _, serialization, _, Ed25519PrivateKey, *_ = _crypto()
    salt = os.urandom(16)
    key  = _derive(password, salt)
    priv = Ed25519PrivateKey.generate()
    raw  = priv.private_bytes(serialization.Encoding.Raw,
                              serialization.PrivateFormat.Raw,
                              serialization.NoEncryption())
    pub  = priv.public_key().public_bytes(serialization.Encoding.Raw,
                                          serialization.PublicFormat.Raw)
    meta = {'salt': _b64(salt), 'public_key': _b64(pub),
            'private_key_enc': Fernet(key).encrypt(raw).decode('ascii'),
            'created': _dt.date.today().isoformat(), 'rounds': PBKDF2_ROUNDS}
    _write_json(KEYMETA, meta)
    _audit('enrolment_setup')

def _private_key(password):
    """Decrypt the signing key with the password; raises ValueError when wrong."""
    Fernet, InvalidToken, _, serialization, _, Ed25519PrivateKey, *_ = _crypto()
    meta = _read_json(KEYMETA, None)
    if not meta:
        raise ValueError('Enrolment has not been set up on this drive yet.')
    try:
        raw = Fernet(_derive(password, _unb64(meta['salt']))).decrypt(
            meta['private_key_enc'].encode('ascii'))
    except InvalidToken:
        _audit('enrolment_password_failed')
        raise ValueError('Enrolment password incorrect.')
    return Ed25519PrivateKey.from_private_bytes(raw)

def _public_key():
    _, _, _, _, _, _, Ed25519PublicKey, _ = _crypto()
    meta = _read_json(KEYMETA, None)
    if not meta:
        return None
    return Ed25519PublicKey.from_public_bytes(_unb64(meta['public_key']))


# ── registry ─────────────────────────────────────────────────────────────────
def _add_months(d, k):
    m = d.month - 1 + k
    y = d.year + m // 12
    m = m % 12 + 1
    last = [31, 29 if y % 4 == 0 and (y % 100 != 0 or y % 400 == 0) else 28,
            31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1]
    return _dt.date(y, m, min(d.day, last))

def _verified_entries():
    """Entries whose signature checks out under the drive's public key."""
    pub = _public_key()
    reg = _read_json(REGISTRY, {})
    out = []
    if pub is None:
        return out
    for e in reg.get('entries', []):
        try:
            pub.verify(_unb64(e.get('sig', '')), _canonical(e))
            out.append(dict(e))
        except Exception:
            tag = (e.get('machine_id', '?'), e.get('sig', '')[:12])
            if tag not in _reported_invalid:          # one audit line per process
                _reported_invalid.add(tag)
                _audit('enrolment_entry_invalid', entry=e.get('machine_id', '?'),
                       reason='signature does not match — entry edited or from another drive')
    return out

_reported_invalid = set()

def enrol(password, note='', confirm_network=False):
    """Enrol (or renew) THIS PC. Returns the new entry."""
    if not is_initialised():
        raise ValueError('Set the enrolment password first.')
    if not identity_ok():
        raise ValueError('The motherboard identifier could not be read on this PC — '
                         'enrolment is refused so the lock cannot be bound to a weak ID.')
    priv = _private_key(password)              # wrong password is reported (and audited) first
    net = network_state()
    if net['connected'] and not confirm_network:
        raise ValueError('NETWORK_CONFIRM')
    today = _dt.date.today()
    try:
        user = getpass.getuser()
    except Exception:
        user = '?'
    entry = {'machine_id': machine_id(), 'hostname': hostname(), 'user': user,
             'enrolled_at': today.isoformat(),
             'expires_at': _add_months(today, TERM_MONTHS).isoformat(),
             'version': version(),
             'network_at_enrol': bool(net['connected']),
             'adapters': net['adapters'], 'ips': net['ips'],
             'note': re.sub(r'[\r\n]+', ' ', str(note or ''))[:120]}
    entry['sig'] = _b64(priv.sign(_canonical(entry)))
    reg = _read_json(REGISTRY, {})
    reg.setdefault('entries', []).append(entry)
    _write_json(REGISTRY, reg)
    _audit('enrolled', expires=entry['expires_at'], network=entry['network_at_enrol'],
           note=entry['note'] or '-')
    return entry

def revoke(machine_id_, password):
    """Remove every entry for a machine (needs the password; audited)."""
    _private_key(password)
    reg = _read_json(REGISTRY, {})
    before = len(reg.get('entries', []))
    reg['entries'] = [e for e in reg.get('entries', []) if e.get('machine_id') != machine_id_]
    _write_json(REGISTRY, reg)
    _audit('revoked', target=machine_id_, removed=before - len(reg['entries']))
    return before - len(reg['entries'])

def entries():
    """Verified entries, newest first, with last-seen information."""
    seen = _read_json(SEEN, {})
    out = []
    today = _dt.date.today()
    for e in _verified_entries():
        exp = _dt.date.fromisoformat(e['expires_at'])
        e['days_left'] = (exp - today).days
        e['state'] = 'EXPIRED' if exp < today else ('EXPIRING' if (exp - today).days <= NOTICE_DAYS else 'ENROLLED')
        e['last_seen'] = seen.get(e['machine_id'], {})
        out.append(e)
    out.sort(key=lambda e: (e['enrolled_at'], e['machine_id']), reverse=True)
    return out


# ── launch-time status ───────────────────────────────────────────────────────
def status(today=None):
    today = today or _dt.date.today()
    st = {'machine_id': machine_id(), 'hostname': hostname(), 'version': version(),
          'identity_ok': identity_ok(), 'network': network_state(),
          'initialised': is_initialised(), 'allowed': False,
          'state': 'NOT_SET_UP', 'expires_at': None, 'days_left': None, 'message': ''}
    if not st['initialised']:
        st['message'] = 'Enrolment has not been set up on this drive. Choose the enrolment password to enrol this PC.'
        return st
    mine = [e for e in _verified_entries() if e.get('machine_id') == machine_id()]
    if not mine:
        st['state'] = 'LOCKED'
        st['message'] = 'This PC is not enrolled. Enter the enrolment password to enrol it.'
        return st
    exp = max(_dt.date.fromisoformat(e['expires_at']) for e in mine)
    st['expires_at'] = exp.isoformat()
    st['days_left'] = (exp - today).days
    if exp < today:
        st['state'] = 'EXPIRED'
        st['message'] = f'This PC\'s enrolment expired on {exp.strftime("%d %b %Y")}. Re-enrol it to continue.'
    elif st['days_left'] <= NOTICE_DAYS:
        st['state'] = 'EXPIRING'; st['allowed'] = True
        st['message'] = f'Enrolment expires on {exp.strftime("%d %b %Y")} ({st["days_left"]} days) — renew it soon.'
    else:
        st['state'] = 'ENROLLED'; st['allowed'] = True
        st['message'] = f'Enrolled until {exp.strftime("%d %b %Y")}.'
    return st

def record_launch(component):
    """Called once by each process at start: last-seen + audit line."""
    st = status()
    seen = _read_json(SEEN, {})
    try:
        user = getpass.getuser()
    except Exception:
        user = '?'
    prev = seen.get(machine_id(), {}).get('last_seen', '')
    now  = _dt.datetime.now().strftime('%Y-%m-%d %H:%M')
    if prev and now[:10] < prev[:10]:
        _audit('clock_earlier_than_last_run', last_seen=prev)
    seen[machine_id()] = {'last_seen': now, 'hostname': hostname(), 'user': user,
                          'version': version(), 'component': component}
    try:
        _write_json(SEEN, seen)
    except Exception:
        pass
    _audit('launch', component=component, state=st['state'],
           network=st['network']['connected'])
    return st
