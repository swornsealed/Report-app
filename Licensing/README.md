# Licensing — private licence issuer

**Keep this folder off the USB drive and out of git.** `private_key.hex` is the only thing
that can sign a licence for the PQ Operator Report Generator; the bundle carries just the
public key (in `Bundle\bundle_licence.py`).

## Routine

| When | Command (from this folder) |
|---|---|
| Once | `python bundle_keygen.py init` — creates the key pair and prints the public key to paste into `bundle_licence.py` |
| New PC | On that PC, read the machine ID from the selection page (or `python bundle_keygen.py id` if the bundle is present) |
| Issue | `python bundle_keygen.py issue --site "Townsville PoC" --machine XXXX-XXXX-XXXX:"office PC" --months 12` |
| Install | copy `issued\<id>.lic` to the drive root, or use **Install licence file** on the selection page |
| QIS review | bump `Bundle\VERSION`, redeploy, then re-issue every site's licence with `--max-version <new>`; old licences do not cover the new bundle |
| Check | `python bundle_keygen.py show issued\<id>.lic` · `python bundle_keygen.py list` (ledger of everything issued) |

## What a licence controls

- **Which PCs** — the machine IDs listed (motherboard-bound), up to `--seats`
- **How long** — 12 months by default, then 30 days' notice and 30 days' grace, then locked
- **Which bundle** — `max_version`; a newer bundle needs a new licence, so every review re-issues

`issued_licences.csv` is the central register of copies in circulation. Back this folder up
somewhere the PoC lead controls (it is small). Losing `private_key.hex` means a new key pair
and re-issuing every licence.
