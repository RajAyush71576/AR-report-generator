# AR Outstanding Report Builder

Automates `AR_Report_Logic.pdf`. Validated against the real 30-09-2026 report:
Public Ins and TPA & CPR totals (old-month and current-month) reproduce the
real file to the rupee.

## Monthly use (no Python needed)

Run `dist\AR_Report_Builder.exe`. Three inputs are required:

1. **Total Bill Wise Report (OP)** -- browse to this month's download.
2. **Discharge Wise IP Revenue Report** -- browse to this month's download.
3. **Previous AR Report** -- the AR report output from the last time this
   report was generated (any file with an `OUTSTANDING STATUS` sheet, not
   just ones this tool generated).

Then click **Generate AR Report**. Everything else is automatic:

- **Report date** is taken from the data itself (the latest Bill Date /
  Discharge Date in the two files) -- this is exactly the date the downloads
  were made through, so there's nothing to type.
- **Output location** is always a `Reports` folder next to the .exe (created
  automatically the first time), saved as `AR as on DD-MM-YYYY.xlsx` --
  regardless of where the OP/IP/previous files live, so every generated
  report ends up in one place.
- **Table 2 (Decrease)**'s previous-report column is read straight from the
  previous AR report's own Table 1 -- specifically its second column (the
  pre-current-month "old" bucket) for Public/TPA, the same date range as this
  run's own old bucket. The report date shown for it is parsed from that
  file's "Outstanding status as on `<date>`" title.

Check the **NEEDS REVIEW** sheet in the output (only appears if something was
flagged) -- it lists any bill whose category text wasn't recognised, so a new
insurer/category name doesn't silently get miscategorised.

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
python -m PyInstaller --onefile --noconsole --name "AR_Report_Builder" app.py
```

The exe lands in `dist\AR_Report_Builder.exe`. Test with `cli.py` first --
it's much faster to iterate on than rebuilding the exe each time:

```
python cli.py "<OP file>" "<IP file>" "<previous AR report>" ["<output folder>"]
```

## Files

- `logic.py` -- all the report-building logic (no UI code).
- `app.py` -- the Tkinter GUI, packaged into the exe.
- `cli.py` -- command-line runner for quick testing.
