---
name: xlsx
description: Use when the task reads, builds, edits or converts a spreadsheet — an Excel workbook (.xlsx), a .csv or .xls file — especially when it needs formulas, totals or a report. Covers writing formulas that stay formulas, calculating them, and checking for error values before the file is handed over.
---

# Excel workbooks

Files you were given are under `inputs/<id>/`. Write every file you deliver to `outputs/`.
Python with `openpyxl` and `pandas` is installed; nothing can be installed.

The helper `/skills/xlsx/scripts/xlsx_tool.py` has three commands and prints JSON:

```
python3 /skills/xlsx/scripts/xlsx_tool.py inspect FILE.xlsx          # sheets, size, header row, formula count
python3 /skills/xlsx/scripts/xlsx_tool.py recalc IN OUT.xlsx         # calculate every formula, save as .xlsx
python3 /skills/xlsx/scripts/xlsx_tool.py check FILE.xlsx            # error values and uncalculated formulas
```

Run it with `--help` for the options.

## Look before you change

1. `inspect` an existing workbook first: sheet names, where the data ends, what the header row says,
   how many formulas it holds.
2. To read data for analysis, use `pandas.read_excel(path, sheet_name=None)` (all sheets) or
   `pandas.read_csv`. Values come back calculated, formulas do not.
3. To change a workbook, open it with `openpyxl.load_workbook(path)` (not `data_only=True`: saving that
   would replace every formula with its last value) and save under a new name in `outputs/`.
   Formats, widths and formulas you did not touch are kept.

## Build it so it stays alive

- **A formula is a string that starts with `=`.** Write `ws["D2"] = "=B2*C2"`, never `ws["D2"] = b * c`.
  Whatever a formula can compute, the sheet computes: totals, differences, percentages, lookups. A typed-in
  number cannot be audited and goes stale the moment an input changes.
- **Put inputs in cells and refer to them.** A rate, a tax percentage or a date lives in one labelled cell;
  formulas point at it (`=B2*$F$1`).
- **One header row**, one record per row below it, no blank rows or merged cells inside the data. A plain
  table like this is what pandas, pivot tables and filters expect. Freeze the header
  (`ws.freeze_panes = "A2"`) and bold it.
- **Number formats carry the meaning**: `cell.number_format = "#,##0.00"` for amounts, `"0.0%"` for a share
  (store `0.256`, not `25.6`), `"yyyy-mm-dd"` for dates (store a real `datetime.date`, not text).
- **Widths**: set `ws.column_dimensions["A"].width` from the longest value, with a floor of 10 and a cap of 60.
- Several sheets are fine, for data, assumptions and summary. Refer across them with `='Sheet name'!B2`.
- English function names and commas between arguments are what the file format stores, whatever the
  user's language: `=SUMIF(A:A,"North",B:B)`, not `SUMME.WENN`.
- A `.csv` or `.xls` source becomes `.xlsx` through `recalc` (it reads those formats). A `.csv` has no
  formulas, so only convert it when the user wants a workbook.

## Calculate and check before you hand over

`openpyxl` writes formulas without results: opened in a viewer that does not calculate, the cells look
empty. So, every time you wrote or changed a formula:

1. `recalc` the file into `outputs/`: `python3 /skills/xlsx/scripts/xlsx_tool.py recalc work.xlsx outputs/report.xlsx`
   (give it a different path than the input, or the same one to replace it).
2. `check` the result. It exits 0 with `"ok": true` only when no cell holds `#REF!`, `#DIV/0!`, `#VALUE!`,
   `#NAME?`, `#N/A`, `#NUM!`, `#NULL!` (or `#SPILL!`, `#CALC!`, `#FIELD!`) and every formula has a value.
   A text cell that only reads like an error is data, not a failure. Otherwise it lists each
   cell (sheet, address, the formula) and exits 1.
3. Fix the cause in the formula or its inputs, then `recalc` and `check` again. Do not deliver a file that
   fails the check, and do not hide an error by typing a value over it. A `#DIV/0!` from an empty input is
   handled in the formula: `=IF(B2=0,"",A2/B2)`.
4. Spot-check one or two totals against the data yourself (for example with pandas) before you tell the
   user the numbers are right.

`recalc` rewrites the file through a spreadsheet application, so styles are kept but exotic features
(macros, pivot tables, some charts) may not survive: say so, or leave such a workbook alone and work on a copy.
If `recalc` answers with an error, read it: a missing or damaged input, or a workbook too large for the time
limit (`--timeout 300` allows longer).

## Hand-over

Tell the user which file is the result, which sheets it holds, which cells are inputs they may change,
and what you checked. If you could not make something work, say what and why instead of delivering it
silently.
