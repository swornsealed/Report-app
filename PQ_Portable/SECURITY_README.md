# Pathology Queensland — Operator Report Generator
## Security Overview for IT / Information Security

**Audience:** Queensland Health IT and information-security staff, and the implementation team
**Designed by:** Craig MacKenzie — Chemistry Department, Townsville Group Laboratory
**System:** Combined portable reporting bundle — i-STAT (Point of Care) and ABL (Blood Gas) operator report generators behind a single selection page
**Form factor:** Self-contained folder on a removable USB drive (bundled Python 3.13 runtime + three loopback-only local web servers). No installation on the host PC; no admin rights required.
**Document status:** Prepared September 2026; amended September 2026 — whole-drive encryption (BitLocker To Go) is now **enabled** on the drive.
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
| ABL rolling histories | `ABL_App\abl_*.json` | **Encrypted (BitLocker To Go)** |
| Monthly middleware exports | `Monthly reports\` | **Encrypted (BitLocker To Go)** |
| Generated Word reports | `Reports\iSTAT\<YYYY-MM Month>\`, `Reports\ABL\<YYYY-MM Month>\` | **Encrypted (BitLocker To Go)** |
| Audit trails | audit / network logs in each app folder | Plaintext by design (accountability), inside the encrypted drive |

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

## 4. Encryption of stored i-STAT history (implemented)

The i-STAT history files contain named staff members' error performance and are encrypted at
rest at the application level:

- **Cipher:** Fernet (AES-128-CBC + HMAC-SHA256 authenticated encryption) from the maintained
  `cryptography` library — no home-made crypto. Tampering is detected, not just prevented.
- **Key derivation:** PBKDF2-HMAC-SHA256, 600,000 iterations, 16-byte random salt.
- **Password handling:** chosen by the operating staff on first launch (min. 8 characters);
  **never stored in any form**. `history.keymeta` holds only the salt and an encrypted
  verifier — possession of the drive yields ciphertext only.
- **Enforcement:** the i-STAT engine refuses to generate reports until unlocked; every unlock
  attempt, **including failures**, is written to the audit log. Pre-existing plaintext backup
  copies of the history were swept into the same encrypted format at setup.
- **No recovery path** by design: a forgotten password requires a history reset (rebuildable
  from retained monthly exports).

## 5. Whole-drive encryption — BitLocker To Go (**enabled**)

Application-level encryption deliberately covers the i-STAT history only. The monthly
exports, generated reports, and ABL histories carry equivalent staff data, so the drive
itself is encrypted with **BitLocker To Go** (AES), enabled from a QH machine in
September 2026. A lost or stolen drive now exposes no readable data of any kind; daily
use is unchanged apart from the password prompt when the drive is inserted. The
combination — password-gated application AES for the most sensitive records, inside a
password-unlocked encrypted drive — provides encryption-at-rest defence in depth.

## 6. Application hardening (both engines)

- **Same-origin (CSRF) checks** on all state-changing routes.
- **Path-traversal confinement** on the "open reports folder" function.
- **No data remanence from uploads:** each uploaded export is a randomly named temporary
  file, deleted as soon as processing finishes — including on error paths.
- **Fail-closed validation:** unrecognisable export layouts are refused with an explicit
  error rather than producing silently empty or wrong reports; metadata/footer lines in
  exports are filtered out.
- **Accountability:** each engine's `audit.log` records every report generation (timestamp,
  Windows username, period, scope); the i-STAT log also records encryption/unlock events.

## 7. Alignment with QH / QGEA security expectations

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
team; audit logs are plaintext by design (inside the encrypted drive); ABL histories rely on
BitLocker rather than app-level encryption; and some managed SOE devices may block
executables on removable media (AppLocker) — which blocks the tool entirely rather than
degrading its security.

## 8. Verifying these claims

1. **Code is inspectable:** `iSTAT_App\app.py`, `ABL_App\app.py`, `portal\portal.py`, and
   both `netguard.py` files are plain Python source. The `python\` folder is the unmodified
   python.org embeddable distribution plus PyPI wheels.
2. **Offline test:** air-gap a machine, run end-to-end, review the network audit logs.
3. **Drive encryption test:** insert the drive on any machine — Windows demands the
   BitLocker password before any file is accessible.
4. **History encryption test:** after unlocking the drive, open any
   `iSTAT_App\*_history.json` in a hex editor — still ciphertext prefixed `ISTATENC1`;
   nothing recoverable without the application password.
5. **Audit trail:** review each engine's `audit.log` for generation and unlock records tied
   to Windows usernames.

*Questions or review requests: contact Craig MacKenzie (Chemistry Department, Townsville
Group Laboratory) or the Pathology Queensland Point of Care team.*
