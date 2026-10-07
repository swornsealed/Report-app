"""Wait until the PQ report portal is accepting connections, then open the browser.

Launched in the background by 'Start PQ Reports.bat'. The first run on a new
machine can take ~30s while Python cold-compiles, so a fixed delay is not
reliable - this polls the port instead and gives up quietly after 3 minutes.
"""
import socket
import time
import webbrowser

PORT = 5750
DEADLINE = time.time() + 180

while time.time() < DEADLINE:
    try:
        with socket.create_connection(("127.0.0.1", PORT), timeout=1):
            pass
        webbrowser.open(f"http://localhost:{PORT}")
        break
    except OSError:
        time.sleep(0.5)
