# Pathology Queensland — Operator Report Generator
## Security Overview for IT / Information Security

**Audience:** Queensland Health IT and information-security staff, and the implementation team
**Conceived, designed and built by:** Craig MacKenzie — Chemistry Department, Townsville Group Laboratory, Pathology Queensland
**System:** Combined portable reporting bundle — i-STAT (Point of Care) and ABL (Blood Gas) operator report generators behind a single selection page
**Form factor:** Self-contained folder on a removable USB drive (bundled Python 3.13 runtime + three loopback-only local web servers). No installation on the host PC; no admin rights required. Runs only on PCs holding a licence issued by the PoC lead (§6).
**Document status:** Prepared September 2026; amended September 2026 — whole-drive encryption (BitLocker To Go) is now **enabled** on the drive; central PC licensing (§6) added 27 September 2026.
**Companion:** `PQ_Report_Generator_Security_Briefing.pptx` — slide version of this document for briefings.

---

## 1. What the program is

A local business tool. Each month a Point of Care staff member loads the middleware export
(`.xlsx`) for the relevant analyser family and the tool produces per-hospital Word reports on
Pathology Queensland letterhead. One launcher (`Start PQ Reports.bat`) starts a selection
page on `localhost:5750`, which supervises two report engines:

| Engine | Port | Data source |
|---|---|---|
| i-STAT — Point of Care | `localhost:5757` | Abbott i-STAT middleware export |
| ABL — Blood Gas | `localhost:5758` | Radiometer ABL middleware export |

Processing is deterministic Python (openpyxl, pandas, python-docx, matplotlib). **There is no
AI/LLM component, no chat interface, no cloud service, and no external dependency at
runtime.** Closing the console window stops all three servers.

## 2. Data on the drive

Classified per QGEA as **SENSITIVE (personal information of QH staff)**:

| Data | Location | Protection |
|---|---|---|
| i-STAT rolling staff/device history (24 months) | `iSTAT_App\*_history.json` | **Encrypted twice**: application-level AES (§4) inside the BitLocker-encrypted drive |
| ABL rolling operator/analyser histories | `ABL_App\abl_*history.json` | **Encrypted twice**: application-level AES (§4) inside the BitLocker-encrypted drive |
| Monthly middleware exports | `Monthly reports\` | **Encrypted (BitLocker To Go)** |
| Generated Word reports | `Reports\iSTAT\<YYYY-MM Month>\`, `Reports\ABL\<YYYY-MM Month>\` | **Encrypted (BitLocker To Go)** |
| Audit trails | audit / network logs in each app folder | Plaintext by design (accountability), inside the encrypted drive |
| Licence files, last-seen record, licence audit | `*.lic`, `licence_seen.json`, `licence_audit.log` at the drive root | Licences are signed and verified against the public key in the bundle; all inside the encrypted drive |

## 3. Network posture — offline by design *and* by enforcement

- **No internet at runtime.** All libraries ship on the drive; there are no API keys, no
  telemetry, no update checks, and no external endpoints in the code.
- **Loopback-only.** All three web servers bind explicitly to `127.0.0.1` — unreachable from
  the LAN.
- **Active runtime network guard** (`netguard.py`, installed at process start in both
  engines): permits outbound sockets **only to loopback**, blocks-and-logs everything else.
  An enforced control — any "phone home" attempt would be refused and recorded in the
  network audit log as `BLOCKED outbound connection attempt to ...`.
- **Auditor's self-test:** disconnect the host from the network and use the tool end-to-end;
  behaviour is identical. Then inspect the network audit logs.

## 4. Encryption of stored history — both engines (implemented)

The i-STAT history files and the ABL operator/analyser history files contain named staff
members' error performance and are encrypted at rest at the application level (each engine
has its own password prompt and key file; the team may choose the same password for both):

- **Cipher:** Fernet (AES-128-CBC + HMAC-SHA256 authenticated encryption) from the maintained
  `cryptography` library — no home-made crypto. Tampering is detected, not just prevented.
- **Key derivation:** PBKDF2-HMAC-SHA256, 600,000 iterations, 16-byte random salt.
- **Password handling:** chosen by the operating staff on first launch (min. 8 characters);
  **never stored in any form**. `history.keymeta` holds only the salt and an encrypted
  verifier — possession of the drive yields ciphertext only.
- **Enforcement:** each engine refuses to generate reports until unlocked; every unlock
  attempt, **including failures**, is written to the audit log. Pre-existing plaintext backup
  copies of the history were swept into the same encrypted format at setup.
- **No recovery path** by design: a forgotten password requires a history reset (rebuildable
  from retained monthly exports).

## 5. Whole-drive encryption — BitLocker To Go (**enabled**)

Application-level encryption covers the histories of both engines. The monthly exports
and generated reports carry equivalent staff data, so the drive itself is encrypted with
**BitLocker To Go** (AES), enabled from a QH machine in
September 2026. A lost or stolen drive now exposes no readable data of any kind; daily
use is unchanged apart from the password prompt when the drive is inserted. The
combination — password-gated application AES for the most sensitive records, inside a
password-unlocked encrypted drive — provides encryption-at-rest defence in depth.

## 6. Where it can run — central licensing (implemented)

The bundle is not signed off by QH IT for general use, so it must not be runnable on any PC it
is copied to, and the number of copies in circulation must be controlled centrally. Every
launch checks for a **licence file** covering the PC it is running on:

- **Issued centrally, never on the drive.** Licences are created only by the PoC lead's offline
  keygen (`PQ_Licensing\pq_keygen.py`, kept off the drive and out of the repository), which
  holds the Ed25519 private key. The bundle carries only the public key, so a licence cannot
  be forged from the drive. The keygen's ledger (`issued_licences.csv`) is the central register
  of every copy licensed.
- **Machine binding.** A machine ID is derived (SHA-256) from the motherboard/system UUID,
  baseboard serial, system-drive serial and hostname. A licence names the machine IDs it covers
  (with a seat cap); a copy of the drive on another PC yields a different ID and stays locked.
- **Term = QIS review cycle.** A licence lasts 12 months. A renewal notice appears in the last
  30 days, a 30-day grace period follows, then the PC is locked until the renewed licence is
  installed.
- **Version gate.** Each licence names the newest bundle version it covers. Updating the bundle
  at a review therefore requires re-issuing licences — a controlled roll-out, and a check that
  every copy in use is a reviewed one.
- **Integrity.** A licence is a signed JSON document (site, machines, seat cap, issued/expires,
  max version). Any edit breaks the signature and the file is refused and logged. A system
  clock earlier than the last recorded run also locks the tool until corrected.
- **Enforcement.** The selection page shows only a lock card (with the machine ID and an
  **Install licence file** control) until a covering licence is present; both report engines
  answer every request with a lock page until then — there is no route around the portal.
- **Accountability.** `licence_audit.log` at the drive root records every launch (PC, Windows
  user, bundle version, network state), licence install, refusal and invalid file. The
  **Licences** page (selection page footer) shows every licence on the drive and the PCs it
  covers, with last use.
- **Stated limit.** The check and the public key live on the drive. Someone who holds the
  BitLocker password and can edit Python could remove the check. It is an administrative
  control with an audit trail and a central register, not tamper-proof DRM.

## 7. Application hardening (both engines)

- **Same-origin (CSRF) checks** on all state-changing routes.
- **Path-traversal confinement** on the "open reports folder" function.
- **No data remanence from uploads:** each uploaded export is a randomly named temporary
  file, deleted as soon as processing finishes — including on error paths.
- **Fail-closed validation:** unrecognisable export layouts are refused with an explicit
  error rather than producing silently empty or wrong reports; metadata/footer lines in
  exports are filtered out.
- **Accountability:** each engine's `audit.log` records every report generation (timestamp,
  Windows username, period, scope); both logs also record encryption/unlock events.

## 8. Alignment with QH / QGEA security expectations

This is a local business tool, not an accredited ICT system; formal assessment remains a QH
decision. The controls map to the obligations QH commonly applies:

| Framework / obligation | How the tool aligns |
|---|---|
| **Information Privacy Act 2009 (Qld) — IPP 4** | Staff data never leaves the device; i-STAT history encrypted at rest; whole-drive encryption pending; no third-party disclosure path (no network). |
| **QGEA IS18:2018 Information Security Policy** | Confidentiality: encryption + offline enforcement. Integrity: authenticated encryption detects tampering; deterministic processing. Availability: self-contained runtime; histories rebuildable from retained exports. |
| **QGEA information security classification** | Data handled as SENSITIVE; storage/handling controls chosen accordingly. |
| **ACSC guidance — portable media encryption** | In force: application-layer AES plus whole-drive AES (BitLocker To Go). |
| **QH SOE endpoint controls** | No installation, no admin rights, no services, no firewall exceptions, no system changes — host controls untouched. |

**Known limitations (stated for transparency):** locally managed tool, not centrally
patched; the history password and the BitLocker password are shared secrets within the PoC
team; the licence signing key is held by the PoC lead alone (its loss means re-issuing every
licence, never a bypass); licence enforcement is code on the drive (evidence, not DRM); audit logs are plaintext by design (inside the encrypted drive); and some managed SOE devices may block
executables on removable media (AppLocker) — which blocks the tool entirely rather than
degrading its security.

## 9. Verifying these claims

1. **Code is inspectable:** `iSTAT_App\app.py`, `ABL_App\app.py`, `portal\portal.py`,
   `pq_licence.py` and both `netguard.py` files are plain Python source. The `python\` folder is the unmodified
   python.org embeddable distribution plus PyPI wheels.
2. **Offline test:** air-gap a machine, run end-to-end, review the network audit logs.
3. **Drive encryption test:** insert the drive on any machine — Windows demands the
   BitLocker password before any file is accessible.
4. **History encryption test:** after unlocking the drive, open any
   `iSTAT_App\*_history.json` (prefix `ISTATENC1`) or `ABL_App\abl_*history.json`
   (prefix `ABLENC1`) in a hex editor — still ciphertext;
   nothing recoverable without the application password.
5. **Audit trail:** review each engine's `audit.log` for generation and unlock records tied
   to Windows usernames.
6. **Licence test:** launch the bundle on a PC with no licence — the portal and both engines
   show only the lock page. Edit any value in a `.lic` file — it is refused and the refusal
   is written to `licence_audit.log`.

*Questions or review requests: contact Craig MacKenzie (Chemistry Department, Townsville
Group Laboratory) or the Pathology Queensland Point of Care team.*
