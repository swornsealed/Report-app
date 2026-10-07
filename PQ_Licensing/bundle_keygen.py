"""bundle_keygen.py — PRIVATE licence issuer for the PQ Operator Report Generator.

Keep this folder OFF the drive and OUT of git. private_key.hex is the only
thing that can sign a licence; the bundle carries just the public key.

  python bundle_keygen.py init                       create the key pair (once)
  python bundle_keygen.py id                         print THIS PC's machine ID
  python bundle_keygen.py issue --site "Townsville PoC" ^
        --machine 2861-7BE8-C880:"PoC office PC" [--machine ...] ^
        [--months 12] [--seats 3] [--max-version 2026.09.27] [--note "QIS review 2026"]
                                                 write issued\\<id>.lic + ledger row
  python bundle_keygen.py show path\\to\\file.lic     verify and print a licence
  python bundle_keygen.py list                       print the ledger of issued licences

A licence is JSON {"payload": {...}, "sig": <hex Ed25519 signature>} over the
canonical payload. Copy the .lic to the drive root (or install it from the
selection page); the bundle verifies it against PUBLIC_KEY_HEX in bundle_licence.py.
"""
import os
import re
import sys
import csv
import json
import argparse
import datetime as _dt
import importlib.util

HERE      = os.path.dirname(os.path.abspath(__file__))
PRIV_PATH = os.path.join(HERE, 'private_key.hex')
ISSUED    = os.path.join(HERE, 'issued')
LEDGER    = os.path.join(HERE, 'issued_licences.csv')
BUNDLE    = os.path.normpath(os.path.join(HERE, '..', 'PQ_Portable'))


def _mod():
    p = os.path.join(BUNDLE, 'bundle_licence.py')
    spec = importlib.util.spec_from_file_location('bundle_licence', p)
    m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
    return m


def _priv():
    from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
    if not os.path.exists(PRIV_PATH):
        sys.exit('No private key — run: python bundle_keygen.py init')
    return Ed25519PrivateKey.from_private_bytes(bytes.fromhex(open(PRIV_PATH).read().strip()))


def cmd_init(_):
    from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
    from cryptography.hazmat.primitives import serialization
    if os.path.exists(PRIV_PATH):
        sys.exit(f'{PRIV_PATH} already exists — delete it deliberately if you really want a new key pair '
                 '(every issued licence would stop verifying).')
    priv = Ed25519PrivateKey.generate()
    raw  = priv.private_bytes(serialization.Encoding.Raw, serialization.PrivateFormat.Raw,
                              serialization.NoEncryption())
    pub  = priv.public_key().public_bytes(serialization.Encoding.Raw, serialization.PublicFormat.Raw)
    with open(PRIV_PATH, 'w') as fh:
        fh.write(raw.hex() + '\n')
    print('private key written to', PRIV_PATH)
    print('PUBLIC_KEY_HEX =', repr(pub.hex()))
    print('paste that value into PQ_Portable\\bundle_licence.py and redeploy the bundle.')


def cmd_id(_):
    m = _mod()
    print('machine id :', m.machine_id())
    print('hostname   :', m.hostname())
    print('identity ok:', m.identity_ok())


def _add_months(d, k):
    mth = d.month - 1 + k
    y, mth = d.year + mth // 12, mth % 12 + 1
    last = [31, 29 if y % 4 == 0 and (y % 100 != 0 or y % 400 == 0) else 28,
            31, 30, 31, 30, 31, 31, 30, 31, 30, 31][mth - 1]
    return _dt.date(y, mth, min(d.day, last))


def cmd_issue(a):
    m = _mod()
    machines = []
    for spec in a.machine:
        mid, _, label = spec.partition(':')
        mid = mid.strip().upper()
        if not re.fullmatch(r'[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}', mid):
            sys.exit(f'bad machine id: {mid!r} (expected XXXX-XXXX-XXXX)')
        machines.append({'id': mid, 'label': label.strip()})
    seats = a.seats or len(machines)
    if len(machines) > seats:
        sys.exit('more machines than seats')
    today = _dt.date.today()
    version = a.max_version or m.version()
    lic_id = f'PQ-{today.strftime("%Y%m%d")}-{re.sub(r"[^A-Za-z0-9]+", "", a.site)[:12].upper()}-{os.urandom(2).hex().upper()}'
    payload = {'licence_id': lic_id, 'product': 'PQ Operator Report Generator',
               'site': a.site, 'issued': today.isoformat(),
               'expires': _add_months(today, a.months).isoformat(),
               'max_version': version, 'seats': seats, 'machines': machines,
               'note': a.note or '', 'issuer': 'PoC lead — Pathology Queensland'}
    sig = _priv().sign(m.canonical(payload)).hex()
    os.makedirs(ISSUED, exist_ok=True)
    out = a.out or os.path.join(ISSUED, f'{lic_id}.lic')
    with open(out, 'w', encoding='utf-8') as fh:
        json.dump({'payload': payload, 'sig': sig}, fh, indent=2)
    new = not os.path.exists(LEDGER)
    with open(LEDGER, 'a', newline='', encoding='utf-8') as fh:
        w = csv.writer(fh)
        if new:
            w.writerow(['licence_id', 'site', 'issued', 'expires', 'max_version', 'seats', 'machines', 'note', 'file'])
        w.writerow([lic_id, a.site, payload['issued'], payload['expires'], version, seats,
                    '; '.join(f"{x['id']} {x['label']}".strip() for x in machines), a.note or '', out])
    print('issued', lic_id)
    print('  site       :', a.site)
    print('  expires    :', payload['expires'], f'({a.months} months)')
    print('  max version:', version)
    print('  machines   :', f'{len(machines)} of {seats} seats')
    for x in machines:
        print('               ', x['id'], x['label'])
    print('  file       :', out)
    print('copy the file to the drive root, or install it from the selection page.')


def cmd_show(a):
    m = _mod()
    with open(a.path, 'rb') as fh:
        pl = m.parse_licence_bytes(fh.read())
    print(json.dumps(pl, indent=2))
    print('signature: VALID under the current public key')


def cmd_list(_):
    if not os.path.exists(LEDGER):
        print('nothing issued yet'); return
    with open(LEDGER, encoding='utf-8') as fh:
        for row in csv.DictReader(fh):
            print(f"{row['licence_id']:34} {row['site']:24} exp {row['expires']}  v<={row['max_version']:12} "
                  f"{row['machines']}")


if __name__ == '__main__':
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest='cmd', required=True)
    sub.add_parser('init').set_defaults(f=cmd_init)
    sub.add_parser('id').set_defaults(f=cmd_id)
    p = sub.add_parser('issue'); p.set_defaults(f=cmd_issue)
    p.add_argument('--site', required=True)
    p.add_argument('--machine', action='append', required=True, help='ID or ID:label (repeatable)')
    p.add_argument('--months', type=int, default=12)
    p.add_argument('--seats', type=int)
    p.add_argument('--max-version', dest='max_version')
    p.add_argument('--note')
    p.add_argument('--out')
    p = sub.add_parser('show'); p.set_defaults(f=cmd_show); p.add_argument('path')
    sub.add_parser('list').set_defaults(f=cmd_list)
    a = ap.parse_args(); a.f(a)
