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


=========================================================================================
CLINIC NOTES -> FILL BLANK CELLS AG:AN, AU:BA, CI, CJ
=========================================================================================

WHAT IT DOES
  For every patient with a blank in AG-AN (Polysubstance Abuse ... Sleep Apnea),
  AU-BA (Pre GLP-1 ... Pre Statin Dose), CI (SBP) or CJ (BMI), it finds the
  files in their patient folder on S:\ (\\STC-NAC) whose name contains
  "clinical note" or "progress note", reads them with OCR, finds the visit
  dated right BEFORE the CCTA exam date, and fills ONLY the blank cells.
  - A cell holding only a space counts as blank.
  - Values are always one of the column's dropdown choices from the
    Data Dictionary sheet (e.g. "None", "Rheumatoid arthritis (RA)",
    "Not initiated", "20 mg").
  - The "None"/"No"/"Not initiated" answers are used only when the note has the
    section that would list it (PMH / A/P for conditions, Meds for drugs).
  - Left blank for a person to check (listed in the review log):
      * AW Pre GLP-1 Dose (low fixed vs titratable can't be read from a note)
      * AV / AY durations when the patient's earlier notes don't show it
        (on the drug at a visit 1+ year before the CCTA -> Long;
         not on it at a visit within that year -> Short)
      * a statin/GLP-1 that is not a dropdown choice (fluvastatin, liraglutide)
        or a drug listed as stopped / intolerant
      * anything only mentioned as "possible" / "rule out" / "evaluate for"
  - The review log warns when the visit is more than 2 months (SBP) or 1 year
    (BMI) before the CCTA, per the Data Dictionary.

HOW TO RUN (after the one-time setup above)
  5_NOTES_TEST_10_PATIENTS.bat      first 10 patients - check these first
  6_NOTES_RUN_ALL_PATIENTS.bat      everyone
  7_NOTES_CHOOSE_DIFFERENT_FILES.bat  pick another Excel file / folder and rescan
  It uses the same Excel file and patient folder you chose before. If asked,
  pick S:\ (or \\STC-NAC). The first run scans the drive once for note files.
  To redo just some rows:  py fill_clinical_notes.py --rows 4249,4250

  Output folder:
    ..._NOTES_FILLED_....xlsx   new copy, yellow = newly filled cells
    notes_review_log_....csv    one line per patient: visit used, every value
                                and the note text it came from, or why blank
    visit_text\                 the text the script read from each visit
