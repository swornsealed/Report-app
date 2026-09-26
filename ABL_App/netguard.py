"""
Active network guard — restricts this process to loopback-only networking.

Installed at app startup (see app.py).  Any attempt by any component of this
process to open an outbound connection to a non-loopback address is refused
and recorded in abl_network_audit.log.  This is a runtime control, not just a
promise: if any library ever tried to "phone home", the connection would fail
and leave an audit line.

It does NOT affect inbound connections (the local browser talking to the app
on 127.0.0.1) or other processes on the machine.
"""
import os
import socket
import datetime

_AUDIT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           'abl_network_audit.log')
_installed = False


def _log(message):
    try:
        with open(_AUDIT_PATH, 'a', encoding='utf-8') as f:
            stamp = datetime.datetime.now().isoformat(timespec='seconds')
            f.write(f'{stamp}  {message}\n')
    except OSError:
        pass


def _is_allowed(address):
    """Allow only loopback destinations (and non-INET addresses like pipes)."""
    if not isinstance(address, tuple) or not address:
        return True                     # AF_UNIX paths etc. — local by nature
    host = str(address[0]).strip().lower()
    if host in ('localhost', '', '0.0.0.0', '::'):
        return True
    if host.startswith('127.'):
        return True
    if host in ('::1', '::ffff:127.0.0.1'):
        return True
    return False


def install():
    """Monkey-patch socket so outbound non-loopback connections are refused."""
    global _installed
    if _installed:
        return
    _installed = True

    orig_connect    = socket.socket.connect
    orig_connect_ex = socket.socket.connect_ex
    orig_sendto     = socket.socket.sendto

    def guarded_connect(self, address):
        if not _is_allowed(address):
            _log(f'BLOCKED outbound connection attempt to {address}')
            raise ConnectionRefusedError(
                f'netguard: outbound connection blocked ({address})')
        return orig_connect(self, address)

    def guarded_connect_ex(self, address):
        if not _is_allowed(address):
            _log(f'BLOCKED outbound connection attempt to {address}')
            return 111                  # ECONNREFUSED
        return orig_connect_ex(self, address)

    def guarded_sendto(self, data, *args):
        # sendto(data, address) or sendto(data, flags, address)
        address = args[-1] if args else None
        if address is not None and not _is_allowed(address):
            _log(f'BLOCKED outbound datagram to {address}')
            raise ConnectionRefusedError(
                f'netguard: outbound datagram blocked ({address})')
        return orig_sendto(self, data, *args)

    socket.socket.connect    = guarded_connect
    socket.socket.connect_ex = guarded_connect_ex
    socket.socket.sendto     = guarded_sendto

    _log('Network guard ACTIVE - outbound connections restricted to loopback only.')
