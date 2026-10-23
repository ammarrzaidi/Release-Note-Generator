# **Release Note Generator (Python + Tkinter)**

A desktop GUI that turns a excel file ticket export (CSV or XLSX) into a
formatted `.docx` release note, matching the structure of your reference
template:

- Title page (company / client / version) + Revision History table
- Table of Contents (a real Word TOC field — right-click it after opening
the doc and choose "Update Field" to populate it)
- Introduction (Project Information + Objectives)
-Purpose (standard boilerplate)
- Summary – Refinement / Technical Improvement (table of all Enhancement
tickets) + one detailed sub-section per ticket
- Summary – Defect (table of all Bug tickets) + one detailed sub-section
per ticket, each ending in a `\[ Paste screenshot here ]` placeholder
- This version has been run end-to-end and its output verified (converted
to PDF and visually checked) — no Visual Studio, .NET, or any compiler
needed. Just Python.

## Setup (one-time)
1. Install Python 3.9+ from https://python.org if you don't have it
(tick "Add to PATH" during install).
2. Install the two required libraries:
```
   pip install python-docx pandas openpyxl
   ```
3. Run
```
python release\_note\_gui.py
```
4. A window opens:
- Browse to your excel file export (.csv or .xlsx).
- Fill in the header fields (Project ID, Client Name, Version No, etc.) — these populate the title page and Introduction section.
- Tick "Include QC Rejected tickets" only if you want rejected tickets
included (excluded by default).
- Click Generate Release Note and choose where to save the `.docx`.
- Files
`release\_note\_core.py` — all the logic (reading the export, building the .docx). No GUI code here, so it's easy to reuse from a script if you ever want to batch-generate without the GUI.
`release\_note\_gui.py` — the Tkinter window.

## Notes
- Grouping: `Category == "Internal CR"` → Enhancements, `Category == "Bug"`
→ Defects. Any other Category value is currently excluded from both sections — let me know if there are other values that should appear.
- Table styling is plain (Word's built-in "Table Grid" style, grey header row) rather than pixel-matching your template's exact fonts/shading — easy to adjust once you compare against a real generated file.
