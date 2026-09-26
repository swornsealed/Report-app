# i-STAT Report Generator

A small **offline** desktop tool for Pathology Queensland Point of Care. It
reads the monthly i-STAT data export (`.xlsx`) and produces a formatted
**Word report per hospital** — split **per device** for Townsville, Mt Isa
and Mackay main — on PQ letterhead, ready to send to each site.

No internet connection is used at runtime and there is **no AI/chat
interface** — it is plain, deterministic spreadsheet reading plus Word
document writing.

---

## What it does

1. You drop the monthly export onto the page. The app reads three sheets from
   it automatically (matched by name, tolerant of naming variations):
   cartridge **usage**, **simulator** runs, and operator **events/errors**.
2. It lists every hospital found, grouped by HHS, with tick boxes.
3. It generates one `.docx` per hospital (or per device for the per-device
   sites) containing:

| Section | Content |
|---------|---------|
| Monthly Snapshot | Error rate + band, simulator status, ceramic cleaning status, top error code, staff alerts |
| Historical Performance | Error-rate trend chart per device (builds up month by month) |
| Performance Summary | Cartridges, error count and colour-banded error rate per device |
| Simulators run | Runs per device, with the daily-simulator reminder note |
| Monthly Ceramic Cleaning | Completed ✓ / Not Completed (3 cleaning cartridges required) |
| Cartridge Error Details | Per-staff error table — staff with **≥3 errors highlighted red** |
| Error Type Summary & Analysis | Top error codes with plain-English cause/explanation from the built-in i-STAT code reference |
| Error Type Trend | Month-to-month top error, repeats flagged ⚠ |
| Staff Requiring Follow-up | Staff with ≥3 errors in **consecutive months** |

### Performance bands (fixed)

**Champions** <4% · **Excellent** 4–6% · **Acceptable** 6–10% ·
**Needs attention** >10%

---

## First-time setup

1. **Install Python 3.10+** (already present on this machine).
2. **Install the Python requirements** (one time, internet required once):
   ```
   pip install -r requirements.txt
   ```
   The launcher does this automatically the first time and never again.

## Running it

Double-click **`start_windows.bat`** (or `start_mac.command`), or run:
```
python app.py
```
Your browser opens at `http://127.0.0.1:5757`.

## Using it

1. Drop the monthly i-STAT export onto the page.
2. Set the report month/year (auto-detected from the filename where possible).
3. Click **Load Hospitals**, tick the hospitals to generate.
4. Click **Generate Reports** — the finished documents land in
   `..\Reports_<Month>_<Year>\<HHS>\i-STAT_<Hospital>_<Month>.docx`
   (per-device sites get one file per device).

## Files & folders

```
iSTAT_App\
  app.py                      the local web app + report engine
  netguard.py                 runtime network guard (see below)
  template.docx               PQ letterhead (logo header/footer)
  istat_error_codes.json      built-in i-STAT error-code reference (offline)
  templates\index.html        the browser interface
  start_windows.bat           double-click launcher
  error_rate_history.json     per-device error-rate history (trend charts)
  staff_error_history.json    ≥3-error staff history (follow-up section)
  error_type_history.json     monthly top-error history (trend table)
  audit.log                   who generated what, when
  istat_network_audit.log     network guard log
```

The three history files build the month-to-month trend charts, the error-type
trend and the "Staff Requiring Follow-up" section. They keep a rolling 24
months. **Note:** `staff_error_history.json` contains identifiable staff
names with error counts — treat the app folder as confidential.

---

## Data privacy & offline assurance (for IT / compliance)

This tool is designed so that **confidential staff-performance information
never leaves the computer it runs on.**

- **No internet, no cloud, no AI service.** There are no API keys, no external
  requests, and no large-language-model / chatbot integration anywhere in the
  code. Report generation is deterministic Python (openpyxl, pandas,
  python-docx, matplotlib) running entirely on this PC. Even the i-STAT
  error-code explanations come from a bundled local reference file, not a
  web lookup.
- **Loopback-only web interface.** The interface is a small local web server
  explicitly bound to `127.0.0.1` (this machine only). It is not reachable
  from other computers on the network.
- **Active network guard (`netguard.py`).** At startup the app installs a
  guard that permits outbound connections **only** to loopback and **blocks
  and logs** any attempt to reach an outside address. This is a runtime
  control, not just a promise: if any component ever tried to "phone home",
  the connection would be refused and recorded.
- **Audit trails.** `istat_network_audit.log` records the guard's activity —
  during normal use it contains only:
  ```
  2026-07-14T12:38:11  Network guard ACTIVE - outbound connections restricted to loopback only.
  ```
  A line beginning `BLOCKED outbound connection attempt to ...` would indicate
  something tried to send data out — under normal operation there are none.
  Separately, `audit.log` records every report generation (timestamp, Windows
  user, period, scope) for accountability.
- **No data remanence.** Uploaded spreadsheets are written to a
  randomly-named temporary file and **deleted automatically** as soon as
  processing finishes — including on errors. No copy of the source data is
  left behind by the app.
- **Request hardening.** Cross-origin requests are rejected (CSRF check on
  all upload routes), and the "open folder" function is confined to the
  reports directory (path-traversal check), so a malicious web page open in
  the same browser cannot drive the app or browse the disk.
- **Where data lives.** Generated reports go to `..\Reports_<Month>\`;
  rolling history stays in the three local JSON files listed above. All on
  local disk — deleting those files removes the data.
- **Host prerequisite.** Reports and history are stored unencrypted by the
  app itself; the machine should run **BitLocker** full-disk encryption (the
  standard QH SOE control) so everything is encrypted at rest.

**Simple proof for yourself or an auditor:** after first-time setup,
disconnect the network (unplug Ethernet / turn off Wi-Fi) and use the app
end-to-end. It works identically, demonstrating no external dependency. Then
open `istat_network_audit.log` to confirm no outbound attempts were made.

---

## Notes / limits

- Built for the standard monthly i-STAT export layout (usage / simulator /
  events sheets). Column headings are matched tolerantly (e.g. "Surname" /
  "Last name"), but a fundamentally different export layout would need the
  loader adjusting.
- Trend charts require `matplotlib`; if it is missing the reports still
  generate, just without the charts (the launcher installs it on first run).
- The report includes a link to the QHEPS point-of-care training page — it
  is printed text in the document for the reader, not something the app
  connects to.
