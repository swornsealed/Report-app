Pathology Queensland - Operator Report Generator (PORTABLE, USB)
=================================================================

WHAT THIS IS
  Both PQ operator report tools in one bundle, with their own private
  copy of Python. Nothing is installed on the computer you plug it
  into: no Python, no packages, no admin rights, no internet ever.
    - i-STAT (Point of Care)   http://localhost:5757
    - ABL (Blood Gas)          http://localhost:5758

TO USE
  1. Plug the USB stick into a Windows PC.
  2. Double-click "Start PQ Reports.bat".
  3. Your browser opens the selection page (http://localhost:5750).
     NEW PC? The page shows a lock card with this PC's machine ID.
     A PoC team member enters the enrolment password to enrol it
     (first ever use: choose the password). Enrolment is tied to
     that PC's motherboard and lasts 12 months.
  4. Choose the report type from the drop-down and click Open.
  5. i-STAT asks for the history password before it will run.
  6. When finished, close the black console window (stops everything),
     then safely eject.

WHERE THINGS GO
  Generated reports ... Reports\iSTAT\<YYYY-MM Month>\  (i-STAT) and
                        Reports\ABL\<YYYY-MM Month>\    (ABL) on this stick
  Source exports ...... keep them in Monthly reports\ at the drive root
  Histories ........... inside iSTAT_App\ (ENCRYPTED - password needed)
                        and ABL_App\ (covered by BitLocker To Go)
  Audit logs .......... audit/network logs inside each app folder;
                        enrolment_audit.log at the drive root
  Enrolled PCs ........ enrolled_pcs.json + enrolment.keymeta at the
                        drive root (see 'Enrolled PCs' on the page)

FOLDER LAYOUT (keep together)
  Start PQ Reports.bat   the launcher - double-click this
  portal\                the selection page
  iSTAT_App\             i-STAT report engine
  ABL_App\               ABL report engine
  python\                private Python + packages (do not modify)
  pq_enrolment.py        PC enrolment check (used by all three)
  VERSION                bundle version, recorded with each enrolment

SECURITY
  - See SECURITY_README.md for the full security overview
    for IT staff.
  - i-STAT history files are encrypted; the password is set by the
    PoC team and is NOT recoverable if forgotten.
  - The whole drive is protected with BitLocker To Go.
  - The bundle runs only on enrolled PCs (not signed off by QH IT
    for general use). Enrol only approved, non-networked PCs.

NOTES
  - The FIRST launch on a new computer can take up to 30 seconds.
  - Some managed (SOE) computers block programs on USB drives; if the
    window closes instantly, that PC's policy is blocking it.

