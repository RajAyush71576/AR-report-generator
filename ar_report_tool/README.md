# AR Outstanding Report Builder

Automates `AR_Report_Logic.pdf`. Validated against the real 30-09-2026 report:
Public Ins and TPA & CPR totals (old-month and current-month) reproduce the
real file to the rupee.

## Monthly use (no Python needed)

Run `dist\AR_Report_Builder.exe`. Step 1 ("Add reports") collects:

1. **Total Bill Wise Report (OP)** -- required, this month's download.
2. **Discharge Wise IP Revenue Report** -- required, this month's download.
3. **Previous AR Report** -- optional. When given (any file with an
   `OUTSTANDING STATUS` sheet, not just ones this tool generated), it supplies
   Table 2 (Decrease). Without it, the output has just Table 1 -- the current
   outstanding totals.

Either source file can be `.xlsx`, `.xls`, or `.csv`. Then click **Generate AR
Report**. Everything else is automatic:

- **Report date** is taken from the data itself (the latest Bill Date /
  Discharge Date in the two files) -- this is exactly the date the downloads
  were made through, so there's nothing to type.
- **Output location** defaults to a `Reports` folder next to the .exe
  (changeable in Step 1), saved as `AR as on DD-MM-YYYY.xlsx`.
- **Table 2 (Decrease)**'s previous-report column is read straight from the
  previous AR report's own Table 1 -- specifically its second column (the
  pre-current-month "old" bucket) for Public/TPA, the same date range as this
  run's own old bucket. The report date shown for it is parsed from that
  file's "Outstanding status as on `<date>`" title.

Two kinds of "new data" open a window in Step 3 ("Review results") instead of
just failing or being buried in an Excel sheet:

- **NEEDS REVIEW** -- a bill landed in TPA & CPR under a category this tool
  doesn't recognise. Grouped by Category + Company Name (two bills can share
  the same category but belong to different payers); pick where each group
  actually belongs and Apply & Regenerate, or Skip to keep the report as
  generated (those bills stay flagged on the NEEDS REVIEW sheet in the
  output). Decisions are remembered in `category_overrides.json`, next to the
  exe, so the same category/company pair is never flagged again.
- **New column layout** -- a report's columns don't match this tool's
  built-in header list (a template change). Map each missing field to the
  file's actual column and Apply & Retry. Mappings are remembered in
  `column_overrides.json`, next to the exe.

## What it does differently from a literal reading of the PDF

One rule was corrected against real data on 2026-10-12: a `Category = ONGC`
bill is only sent to **OP GNL** when its Company Name is also blank. If a
Company Name is filled in (as it is for all 12 historical ONGC bills checked),
it's treated like any other payer and stays in **OP&IP TPA & CPR** -- matching
what every past report has actually done. (Confirmed with the user; this is
intentional, not an oversight.)

## Rebuilding the .exe after a code change

```
cd ar_report_tool
python -m pip install -r requirements.txt
python -m PyInstaller --onefile --noconsole --name "AR_Report_Builder" --collect-data=customtkinter app.py
```

(`--collect-data=customtkinter` is required -- without it the exe is missing
the GUI's theme file and fails to start.)

The exe lands in `dist\AR_Report_Builder.exe`. Test with `cli.py` first --
it's much faster to iterate on than rebuilding the exe each time:

```
python cli.py "<OP file>" "<IP file>" ["<previous AR report>" | -] ["<output folder>"]
```

## Files

- `logic.py` -- all the report-building logic (no UI code).
- `app.py` -- the Tkinter GUI, packaged into the exe.
- `cli.py` -- command-line runner for quick testing.
