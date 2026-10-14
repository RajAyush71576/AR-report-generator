# AR Outstanding Report -- Project Context

A running record of what this project is, what was done, and why -- for
picking this back up later (a new session, a teammate, or future-you).

## 1. Starting point

The user had a PDF, `AR_Report_Logic.pdf` (on the Desktop), written by @Team 2
on 2026-10-06, documenting how the monthly "AR Outstanding Report" is
currently built **by hand** in Excel, so that it could be automated. The goal
stated by the user: review the doc and figure out the best way to (a) turn
this into an `.exe` and (b) present the report better.

The doc's logic, in short:
1. Take OP (outpatient) bills from "Total Bill Wise Report" and IP (inpatient)
   final bills from "Discharge Wise IP Revenue Report", confirmed only.
2. Keep bills dated 01-04-2021 through the report date.
3. Route every bill into one of three buckets: **OP GNL** (no company / general
   categories), **Public Ins** (ESI/ECHS/CHS), **TPA & CPR** (everything else).
4. Amount outstanding per bill = Payment Due (Net Amt - Paid Amt).
5. Build a 4-sheet output workbook: `OUTSTANDING STATUS` (two summary tables),
   `OP&IP PUBLIC INS`, `OP&IP TPA & CPR`, `OP GNL`.

## 2. Validating the doc against the real data (before writing any code)

Rather than trust the PDF blindly, I installed pandas/openpyxl and read the
user's actual files in `C:\Users\LIVE\Desktop\AR_reports\`:
- `Total Bill Wise Report (8).xlsx` (17,193 OP bills)
- `Discharge Wise IP Revenue Report (9).xlsx` (7,295 raw IP rows)
- `AR as on  30-092026 (2).xlsx` -- a real, already-generated output file for
  report date 30-09-2026, used as ground truth.

I reconstructed the whole pipeline in a throwaway script and compared my
computed totals against this real output file. Findings:

- **Public Ins totals matched exactly** on the first try (old-month and
  current-month), confirming most of the doc is accurate.
- **TPA & CPR was off by exactly ₹2,25,594** in the old-month bucket. Root
  cause, found by diffing bill-by-bill: 12 IP bills with `Category = ONGC`
  **and a real Company Name filled in** ("Oil & Natural Gas Corporation LTD.")
  sit in **TPA & CPR** in the real file, not **OP GNL** as the doc's literal
  rule ("Category is ONGC -> OP GNL") would require.
- This was a genuine ambiguity only the user could resolve, so I asked via
  `AskUserQuestion`. **Decision: match current actual behavior** -- ONGC only
  routes to OP GNL when Company Name is also blank; otherwise it's treated
  like any other payer and stays in TPA & CPR. This is now hard-coded into
  `route_sheet()` in `logic.py` (ONGC is simply absent from `GNL_PREFIXES`,
  and present in `KNOWN_TPA_PREFIXES` so it isn't flagged for review either).
- Also confirmed the fuller "GNL" category list from page 4 of the doc
  (`Study-Faculty`, `DEP*`, `FDM*`, `P.Card.Fund`, `PM Fund`, not just
  "blank company or ONGC") -- these needed explicit prefix matching, not just
  the blank-company rule.
- Confirmed two real, recurring data-quality issues already described (loosely)
  in the doc:
  - **Shifted rows** in the IP file: some rows have every column from
    `Bill Status` onward pushed one place right, so `CONF`/`OPN` lands under
    `Confirmed Date` instead. 10 such rows existed in the live download.
    Deterministic fix: shift the affected columns back left.
  - **Junk rows** with no Bill No.: both the documented "counter heading" rows
    (e.g. "IP Bill Counter - 1") and two undocumented grand-total footer rows
    ("Payment Due Bills Total = Rs.", "Net Total = Rs."). Both are removed by
    one simple rule: drop any row where Bill No. is blank.
  - Noticed the real output file's detail-sheet titles were stale (said
    "period 1/4/2021 - 19/09/2026", left over from the previous month's copy
    -- a manual-process bug). The automated tool always regenerates this
    correctly from the actual report date.

Once the ONGC rule was fixed per the user's decision, **all four numbers
(Public Ins old/current, TPA & CPR old/current) matched the real report to
the rupee**, giving high confidence to proceed.

## 3. Building the tool (`ar_report_tool/`)

- **`logic.py`** -- all the report logic, no UI code. Key design choice:
  source columns are read **by position**, not by header text, because the
  two source reports are fixed-layout system exports whose header text has
  inconsistent whitespace/line-breaks between monthly downloads. A column-count
  check (`_require_cols`) guards against an actual layout change upstream.
  Includes a `NEEDS REVIEW` safety net: any bill routed into TPA & CPR whose
  category text doesn't match a known prefix gets flagged on its own sheet
  instead of silently blending in -- so a brand-new insurer/category name
  doesn't go unnoticed.
- **`app.py`** -- small Tkinter GUI, the file PyInstaller packages into the exe.
- **`cli.py`** -- command-line runner, used throughout for fast iteration
  (much quicker than rebuilding the exe every time logic.py changes).
- **`requirements.txt`**, **`README.md`** -- setup/usage/rebuild instructions.
- Packaged with PyInstaller: `python -m PyInstaller --onefile --noconsole
  --name "AR_Report_Builder" app.py` -> `dist\AR_Report_Builder.exe`
  (~53 MB, no Python install needed to run it).

Every change was tested against the real source files via `cli.py` before
rebuilding the exe, to keep the numbers pinned to the validated ground truth.

## 4. Simplification: down to 2 inputs

Original version asked for 4 inputs (OP file, IP file, optional previous
report, report date, output folder). The user asked for **just the 2 source
files as input**. Changed `build_report()` so everything else is inferred:

- **Report date** = the latest Bill Date / Discharge Date found across both
  input files (this is, by construction, exactly the date the downloads were
  taken through).
- **Output folder** = the OP file's own folder (matches how the user already
  keeps OP/IP downloads and the AR output together in one folder).
- **Previous report (for the "Decrease in Outstanding status" table)** =
  auto-detected by scanning the output folder for files matching this tool's
  own filename pattern (`AR as on DD-MM-YYYY.xlsx`) and picking the latest one
  dated before the current report date. If none is found, Table 2 is skipped
  with a note in the status message, rather than failing. (Only recognises
  files this tool generated -- the pre-existing manually-named file in the
  folder, `AR as on  30-092026 (2).xlsx`, won't be picked up automatically.)

The GUI (`app.py`) was correspondingly stripped down to exactly two file
pickers and a Generate button.

## 5. Bugs found and fixed after the first working version

1. **End-of-day cutoff bug** in `compute_table1`: the "current month" filter
   compared `Date <= report_date` where `report_date` has no time component
   (midnight). IP discharge dates carry real timestamps (e.g. `21:00`), so
   any bill discharged on the report date itself with a non-midnight time was
   silently excluded from both the old and current buckets -- 30 bills,
   ₹31,70,621 in TPA & CPR, vanished from the totals. Fixed by comparing
   against end-of-day (`report_date + 23:59:59`), matching the cutoff already
   used when building the combined dataset.
2. **`####` (hash marks) in the Date column** of the generated Excel file:
   openpyxl was writing raw Python `datetime` values with no explicit
   `number_format`, so Excel fell back to a wide default datetime format that
   didn't fit the column width. Fixed by explicitly setting the Date column's
   format to `dd-mm-yyyy` in both the detail sheets and the NEEDS REVIEW
   sheet.
3. Non-portable `%-d`/`%-m` strftime codes (Linux-only) were used in several
   date-label strings; replaced with plain `%d`/`%m` for Windows compatibility
   (this is a Windows-only tool).
4. `EMP-TPA` was initially flagged on the NEEDS REVIEW sheet as an unknown
   category even though it's a legitimate, recurring TPA category seen in both
   source files -- added to `KNOWN_TPA_PREFIXES` to stop the false positive.

## 6. "Decrease" table (Table 2) missing after a real run

After the user actually ran the exe on their real files (now living directly
in `C:\Users\LIVE\Desktop\`, not the `AR_reports\` subfolder), the
`OUTSTANDING STATUS` sheet only showed Table 1 -- the "Decrease in
Outstanding status" table was missing.

Cause: `find_previous_report()` only recognised files matching this tool's
own output filename pattern (`AR as on DD-MM-YYYY.xlsx`). The only AR report
actually present was the original manually-built file,
`AR as on  30-092026 (2).xlsx` (irregular spacing, no dash before the year,
a "(2)" suffix) -- it didn't match the pattern, so it was invisible to the
auto-detect logic and Table 2 was silently skipped.

Fix: `find_previous_report()` now ignores filenames entirely. It scans every
`.xlsx` in the output folder (skipping Excel's `~$` lock files), opens each,
and checks for an `OUTSTANDING STATUS` sheet with an "Outstanding status as
on `<date>`" title -- reading that title text directly, the same way
`read_previous_totals()` already did. It still prefers the latest file dated
strictly before the current report date; if the only match shares the exact
same date (as happened here -- the pre-existing file was also generated for
30-09-2026), it falls back to using that file rather than finding nothing,
and the label in Table 2 reads "as on previous report (same date)" instead of
repeating an identical date in both columns.

Verified by copying the user's real OP file, IP file, and the pre-existing
`AR as on  30-092026 (2).xlsx` into a scratch folder and re-running: the tool
found the old file via its same-date fallback and both tables rendered
correctly in the output.

## 7. Current state

- `dist\AR_Report_Builder.exe` is built and working: takes the OP file and IP
  file, infers everything else, and produces `AR as on DD-MM-YYYY.xlsx` with
  all four sheets (plus the conditional NEEDS REVIEW sheet).
- Validated end-to-end against the real 30-09-2026 data: Public Ins
  (₹49,14,65,738.81) and TPA & CPR (₹56,94,17,914.00) match the real report
  exactly; OP GNL total (₹1,26,80,461.51) also confirmed to match.
- Previous-report auto-detection (for the Decrease table) now works whether
  or not the prior file was generated by this tool, as long as it has an
  `OUTSTANDING STATUS` sheet with the standard title text.
- The user's real source files currently live directly in
  `C:\Users\LIVE\Desktop\` (not `AR_reports\`) alongside the original manual
  `AR as on  30-092026 (2).xlsx` -- that's fine, the tool doesn't care where
  the files sit as long as the two inputs are pointed to directly.
- Nothing known to be broken or pending as of this writing.
