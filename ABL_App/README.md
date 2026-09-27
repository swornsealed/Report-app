# Radiometer ABL — Operator Error Report Generator

A small **offline** desktop tool for Pathology Queensland Point of Care. It
reads the monthly *Operator error report* Excel export from the Radiometer ABL
blood-gas middleware and produces a formatted **Word report per department**,
on PQ letterhead, ready to send to each site.

No internet connection is used at runtime and there is **no AI/chat
interface** — it is plain, deterministic spreadsheet reading plus Word
document writing.

---

## What it does

> **Since 27 Sep 2026:** the operator and analyser histories are **AES-encrypted** with a
> password chosen at first launch (same design as the i-STAT app; no recovery if forgotten);
> every run also writes one **Hospital Summary** per hospital
> (`ABL_<Hospital>_Hospital_Summary_<Month>.docx`: departments, analysers, error rates,
> flagged and low-volume operator counts, and analysers with recent history that did not
> appear this month); and the trend charts show the last 12 months in 12 fixed slots, never
> beyond the report month, with every error segment labelled.


1. You drop the monthly export (`.xlsx`) onto the page.
2. The app lists every **Hospital → Department** found, with the analysers
   belonging to each department shown underneath. Tick the reports you want.
3. It generates one `.docx` per department containing:

| Section | Content |
|---------|---------|
| Monthly Snapshot | Department error rate, analysers, top error code, operators with errors, flags |
| Historical Performance | Per-analyser trend charts: error rate, plus tests-run line with a colour-coded monthly error-type breakdown beneath it (build up month by month) |
| Analyzer Performance | Total tests on analyser, total errors, error % per analyser |
| Operator Error Details by Analyzer | Per-analyser operator table: tests, error codes 328–791, error % |
| PICU Operators (ICU reports only) | Staff on the PICU operator list who recorded errors under ICU are listed in their own section, cross-referenced by name |
| Error Type Summary | Department-wide count of each ABL error code |
| Operators Requiring Review | 12-month lookback: operators flagged in 2+ months, with error rate, error types and month of each flag |

### Flag rules (fixed)

- An operator is **flagged** (red, ⚠) when they have **more than 10 tests**
  (11+) and an error rate **over 10%** in the month.
- Operators with **10 tests or fewer** are never flagged (low volume, marked
  `*`, "not assessed").
- The Monthly Snapshot shows intervention rows (flagged / recurring) **only
  when operators are flagged in the current month**.
- **Analyser error %** = sum of the error-code columns ÷ the *Total tests on
  analyzer* value for that analyser.

### Input quirks handled automatically

The raw export is messy; the loader corrects for all of it: a broken
`#NAME?` hospital header, the hospital name appearing only on the first row,
blank operator/department cells on continuation rows, percentages stored as
text (`"16.67%"`), two-line column headings, and trailing spaces in analyser
names. If a file's error columns cannot be recognised, the app refuses with a
clear message rather than silently reporting zero errors.

---

## First-time setup

1. **Install Python 3.10+** (already present on this machine).
2. **Install the Python requirements** (one time, internet required once):
   ```
   pip install flask python-docx openpyxl pandas matplotlib
   ```
   The launcher does this automatically the first time and never again.

## Running it

Double-click **`start_windows.bat`** (or `start_mac.command`), or run:
```
python app.py
```
Your browser opens at `http://127.0.0.1:5758`.

## Using it

1. Drop the monthly ABL export onto the page.
2. Set the report month/year (auto-detected from the filename where possible).
3. Click **Load Analyzers**, tick the departments to generate.
4. Click **Generate Reports** — the finished documents land in
   `..\ABL_Reports_<Month>_<Year>\<Hospital>\ABL_<Department>_<Month>.docx`.

## Files & folders

```
ABL_App\
  app.py                       the local web app + report engine
  netguard.py                  runtime network guard (see below)
  template.docx                PQ letterhead (logo header/footer)
  picu_operator_list.xlsx      PICU staff list (replace file to update; re-read each run)
  templates\index.html         the browser interface
  start_windows.bat            double-click launcher
  abl_analyzer_history.json    per-analyser error-rate history (trend charts)
  abl_operator_history.json    flagged-operator history (follow-up section)
  abl_audit.log                who generated what, when
  abl_network_audit.log        network guard log
```

The two history files build the month-to-month trend charts and the
"Operators Requiring Follow-up" section. They keep a rolling 24 months.

---

## Data privacy & offline assurance (for IT / compliance)

This tool is designed so that **confidential staff-performance information
never leaves the computer it runs on.**

- **No internet, no cloud, no AI service.** There are no API keys, no external
  requests, and no large-language-model / chatbot integration anywhere in the
  code. Report generation is deterministic Python (openpyxl, pandas,
  python-docx, matplotlib) running entirely on this PC.
- **Loopback-only web interface.** The interface is a small local web server
  explicitly bound to `127.0.0.1` (this machine only). It is not reachable
  from other computers on the network.
- **Active network guard (`netguard.py`).** At startup the app installs a
  guard that permits outbound connections **only** to loopback and **blocks
  and logs** any attempt to reach an outside address. This is a runtime
  control, not just a promise: if any component ever tried to "phone home",
  the connection would be refused and recorded.
- **Audit trails.** `abl_network_audit.log` records the guard's activity —
  during normal use it contains only:
  ```
  2026-07-14T12:33:37  Network guard ACTIVE - outbound connections restricted to loopback only.
  ```
  A line beginning `BLOCKED outbound connection attempt to ...` would indicate
  something tried to send data out — under normal operation there are none.
  Separately, `abl_audit.log` records every report generation (timestamp,
  Windows user, period, scope) for accountability.
- **No data remanence.** Uploaded spreadsheets are written to a
  randomly-named temporary file and **deleted automatically** as soon as
  processing finishes — including on errors. No copy of the source data is
  left behind by the app.
- **Request hardening.** Cross-origin requests are rejected (CSRF check on
  all upload routes), and the "open folder" function is confined to the
  reports directory (path-traversal check), so a malicious web page open in
  the same browser cannot drive the app or browse the disk.
- **Where data lives.** Generated reports go to `..\ABL_Reports_<Month>\`;
  rolling history stays in the two local JSON files listed above. All on
  local disk — deleting those files removes the data.
- **Host prerequisite.** Reports and history are stored unencrypted by the
  app itself; the machine should run **BitLocker** full-disk encryption (the
  standard QH SOE control) so everything is encrypted at rest.

**Simple proof for yourself or an auditor:** after first-time setup,
disconnect the network (unplug Ethernet / turn off Wi-Fi) and use the app
end-to-end. It works identically, demonstrating no external dependency. Then
open `abl_network_audit.log` to confirm no outbound attempts were made.

---

## Notes / limits

- Built for the Radiometer ABL *Operator error report* export layout (as
  sampled, May 2026). A very different middleware export layout would need
  the loader adjusting — it will say so explicitly rather than produce empty
  reports.
- The export's own `Pct. error` column is treated as authoritative for
  operator error rates (it de-duplicates repeated errors on a single
  measurement); raw error-code counts are shown alongside.
- Only the first worksheet ("Operator error ABL") is read; the "errors
  greater than 10%" and "operators for follow up" tabs are filter views of
  the same data and are derived independently by the app.
