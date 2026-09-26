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
  Audit logs .......... audit/network logs inside each app folder

FOLDER LAYOUT (keep together)
  Start PQ Reports.bat   the launcher - double-click this
  portal\                the selection page
  iSTAT_App\             i-STAT report engine
  ABL_App\               ABL report engine
  python\                private Python + packages (do not modify)

SECURITY
  - See SECURITY_README.md for the full security overview
    for IT staff.
  - i-STAT history files are encrypted; the password is set by the
    PoC team and is NOT recoverable if forgotten.
  - The whole drive is protected with BitLocker To Go.

NOTES
  - The FIRST launch on a new computer can take up to 30 seconds.
  - Some managed (SOE) computers block programs on USB drives; if the
    window closes instantly, that PC's policy is blocking it.

