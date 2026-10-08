CCTA INTAKE -> FILL BLANK CELLS AA:AF (Known CAD, Prior MI, Stents, Kidney, Smoker, AFib)
=========================================================================================

WHAT IT DOES
  For every patient in the MASTER sheet with a blank in columns AA-AF, it finds
  their folder on S:\ (folder names: patientID_LASTNAME_FIRSTNAME_DOB), opens the
  ccta_intake file, reads the Yes/No boxes in the "Patient History" table, and
  fills ONLY the blank cells.
  - Your original Excel file is never changed. A new copy is saved in the
    "output" folder, with the newly filled cells highlighted light yellow.
  - Cells that already have a value are never touched.
  - If a mark is unclear (nothing checked, both checked, scribbles), the cell
    stays blank and is listed in the review log so you can check it by hand.

ONE-TIME SETUP
  1. Install Python (if you don't have it): https://www.python.org/downloads/
     Run the installer and TICK "Add python.exe to PATH" on the first screen.
  2. Copy this whole "ccta_intake_fill" folder to your computer (e.g. Desktop).
  3. Double-click 1_INSTALL.bat and wait until it says "Done!".

TEST RUN (first 10 patients)
  4. Double-click 2_TEST_10_PATIENTS.bat
     - A window asks you to pick the Tampa Insight Excel file, then the folder
       that holds the patient folders (S:\ or the patients folder inside it).
       Your choices are remembered.
     - The first run scans S:\ once to list the ccta_intake files. This can take
       a while on a big drive; later runs are fast.
  5. When it finishes, the "output" folder opens. Check:
     - review_log_TEST_....csv     (open in Excel - one line per patient)
     - check_images\               (a picture of each form's table: GREEN boxes
                                    = read as Yes, BLUE = No, RED = unclear)
     - ..._FILLED_TEST_....xlsx    (yellow cells are the new values)

FULL RUN
  6. If the test looks right, double-click 3_RUN_ALL_PATIENTS.bat
     (forms already read during the test are not read again).

OTHER
  - 4_CHOOSE_DIFFERENT_FILES.bat: pick a different Excel file / patient folder,
    and rescan for intake files (use this if new patient folders were added).
  - Statuses in the review log:
      OK                     all blank cells filled
      PARTIAL: AD unclear    that cell left blank - check the form by hand
      NO INTAKE FILE FOUND   no ccta_intake file matched that patient
      "Matched by ... please check"  the folder matched on name only or ID only
