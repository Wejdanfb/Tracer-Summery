MCH TRACER SUMMARY  -  local web tool
=====================================

Reads the Nursing Tracer Excel files (one per unit) and fills the uploaded
Tracer Summary template to produce ONE file:  TRACER SUMMARY.xlsx
It runs on your own computer and opens in Google Chrome. Files never leave
this machine. The uploaded tracers and template are only read, never changed.


1. ONE-TIME SETUP (Windows)
---------------------------
 1. Install Python 3 (3.9 or newer) from https://www.python.org/downloads/
    IMPORTANT: on the first installer screen tick "Add python.exe to PATH".
 2. Extract this ZIP to a folder, e.g.  C:\MCH Tools\tracer_summary_tool


2. START THE TOOL
-----------------
 Double-click  run.bat
   - The first time only, it installs 3 packages (Flask, openpyxl, Pillow).
     This needs internet once.
   - A small black window opens and Chrome opens at  http://127.0.0.1:5051
     (if Chrome is not installed, your default browser opens).
   - Keep the black window open while you use the tool. Close it to stop.
 The Chart Review tool uses port 5050, so both tools can run at the same time.


3. USING IT
-----------
 1. Upload Tracer Excel Files - drag the files in, or click the button and
    select many at once. Each file is checked immediately. The list shows the
    unit, final score and number of findings. Use arrows to reorder and x to
    remove a file.
 2. Upload Tracer Summary Template - the summary .xlsx used as the master layout.
 3. Optional settings:
      MONTH\YEAR  - leave empty to keep the template's month
      File name   - default "TRACER SUMMARY"
 4. Click GENERATE TRACER SUMMARY and watch the progress
    ("Processing 3 of 17 - NICU.xlsx ...").
 5. Click DOWNLOAD TRACER SUMMARY.xlsx.


4. RULES APPLIED
----------------
 Rows ................ ONE ROW = ONE UNIT/WARD (one row per uploaded tracer file)
 Unit/Ward ........... cell C6 of each tracer (merged cell handled)
 Tracer Score (%) .... the tracer's own Final Score (e.g. "Final score: 92.37%"),
                       not recalculated
 Significant Findings  every item scored 0 or 1 of that unit, ALL in the unit's one
                       cell, one finding per line (Wrap Text on, row height grown to fit),
                       each written as
                           STD No. - Notes                 (Notes filled)
                           STD No. - Checklist item        (Notes empty)
                       e.g.  NR.3.6.7 - The staff need reeducation on the policy
                       Notes and descriptions are copied exactly as written.
                       When the STD cell is split over several lines in the tracer
                       (e.g. "QM.14/QM" + ".15") it is rejoined as QM.14/QM.15.
                       Items scored 2 or NA are excluded.
                       Scores are read whether stored as numbers or text (" 1 ", "na").
                       A unit with no 0/1 items still gets one row with "/".
 Action Plan / Evidence  always BLANK (the Yes/No drop-down is kept)
 TOTAL SCORE ......... moved under the last row; average of one score per unit
 Row height .......... Excel allows at most 409 pt (about 27 lines). A unit with more
                       lines than that shows a warning; all text is still in the cell.
 Formatting .......... logo, header, fonts, borders, colours, merged cells,
                       column widths and page setup come from the template.


5. MESSAGES
-----------
 Duplicate   - the same file was uploaded twice. Remove it with x, or keep it.
 Same unit   - two files have the same unit in C6 (for example, two months).
               Both are included unless you remove one.
 "Final Score formula has no saved result"
             - the tracer was not last saved in Excel. The score is calculated
               with the tracer's own formula. Opening and saving it in Excel fixes this.
 "Final score in file differs from the checklist scores"
             - usually a score typed as TEXT that Excel's own formula ignores.
               The file's value is used. Please check that tracer.
 Error       - the file cannot be read or has no checklist. It is skipped and
               listed under "Files not included". All other files still process.

 After generating, the tool re-opens the result and runs 26 checks (one row per unit, units,
 scores, every 0/1 finding present with its STD No., 2/NA excluded, blank action columns,
 template formatting unchanged). See "Output checks" on the result panel.


6. FILES
--------
 run.bat              start the tool
 app.py               local web server (Flask)
 templates/, static/  the web page
 tracer_processor.py  reads each tracer
 summary_writer.py    fills the template
 validator.py         checks the output
 excel_utils.py, core.py   helpers
 cli.py               optional command-line mode (not needed for normal use)
 uploads/, outputs/   temporary working folders, emptied every time the tool starts
