CLINIC NOTES -> FILL BLANK CELLS AG:AN, AU:BA, CI, CJ
=====================================================

WHAT IT DOES
  For every patient with a blank in AG-AN (Polysubstance Abuse ... Sleep Apnea),
  AU-BA (Pre GLP-1 ... Pre Statin Dose), CI (SBP) or CJ (BMI), it finds the
  files in their patient folder on S:\ (\\STC-NAC) whose name contains
  "clinical note" or "progress note", reads them with OCR, finds the visit
  dated right BEFORE the CCTA exam date, and fills ONLY the blank cells.
  - Your original Excel file is never changed. A new copy is saved in the
    "output" folder, with the newly filled cells highlighted light yellow.
  - Cells that already have a value are never touched. A cell holding only a
    space counts as blank.
  - Values are always one of the column's dropdown choices from the
    Data Dictionary sheet (e.g. "None", "Rheumatoid arthritis (RA)",
    "Not initiated", "20 mg").
  - "None"/"No"/"Not initiated" is used only when the note has the section
    that would list it (PMH / A/P for conditions, Meds for drugs).
  - Left blank for a person to check (reason is in the review log):
      * AW Pre GLP-1 Dose (low fixed vs titratable can't be read from a note)
      * AV / AY durations when the patient's earlier notes don't show it
        (on the drug at a visit 1+ year before the CCTA -> Long;
         not on it at a visit within that year -> Short)
      * a statin/GLP-1 that is not a dropdown choice (fluvastatin, liraglutide)
        or a drug listed as stopped / intolerant
      * anything only mentioned as "possible" / "rule out" / "evaluate for"
  - The review log warns when the visit is more than 2 months (SBP) or 1 year
    (BMI) before the CCTA, per the Data Dictionary.

ONE-TIME SETUP
  1. Install Python (if you don't have it): https://www.python.org/downloads/
     Run the installer and TICK "Add python.exe to PATH" on the first screen.
  2. Copy this whole "clinic_notes_fill" folder to your computer (e.g. Desktop).
  3. Double-click 1_INSTALL.bat and wait until it says "Done!".
     (Not needed again if you already ran 1_INSTALL.bat for ccta_intake_fill.)

TEST RUN (first 10 patients)
  4. Double-click 2_TEST_10_PATIENTS.bat
     - Pick the Tampa Insight Excel file, then S:\ (or \\STC-NAC).
       Your choices are remembered.
     - The first run scans S:\ once for note files; later runs are fast.
  5. When it finishes, the "output" folder opens. Check:
     - ..._NOTES_FILLED_TEST_....xlsx  (yellow cells are the new values)
     - notes_review_log_TEST_....csv   (one line per patient: visit used,
                                        every value and the note text it came
                                        from, or why it was left blank)
     - visit_text\                     (the text the script read from each visit)

FULL RUN
  6. If the test looks right, double-click 3_RUN_ALL_PATIENTS.bat

OTHER
  - 4_CHOOSE_DIFFERENT_FILES.bat: pick a different Excel file / patient folder
    and rescan for note files (use this if new patient folders were added).
  - Only some rows:  py fill_clinical_notes.py --rows 4249,4250
