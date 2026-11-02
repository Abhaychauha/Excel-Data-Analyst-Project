# Excel Practical Assignment - Superstore Solution

A complete solution to the Excel Practical Assignment (12 sections, 43 topics), built on the **Sample - Superstore** dataset (Orders 9,994 rows, Returns 296, People 4).

## Files

| File | Purpose |
|---|---|
| `Excel_Assignment_Solution.xlsx` | The main workbook: one sheet per section, all formulas live on the Orders data |
| `Macros.bas` | VBA code for Section 12 (import into the VBA editor) |
| `README.md` | This guide |

## Workbook map

| Section | Sheet(s) | Topics covered |
|---|---|---|
| Start | `Index` | Clickable table of contents and reading guide |
| 1 Basics | `S1_Basics` | SUM/MAX/MIN/SMALL/LARGE/MEDIAN, Transpose vs TRANSPOSE, COUNT/COUNTA/COUNTBLANK, status bar, YoY growth and CAGR with POWER |
| 2 Formatting | `S2_Formatting` | Currency, conditional formats, number formats, wrap text, merge, rotated headers, Format Painter |
| 3 Sort | `S3_Sort` | Sorting numbers, dates, text, colours, two-level and custom sorts |
| 4 Dates | `S4_Dates` | Durations with timestamps, DATEDIF, TODAY/DATE/YEAR/MONTH/DAY, EDATE/EOMONTH |
| 5 Logical | `S5_Logical` | IF, AND, OR, nested IF, IFERROR |
| 6 Cleaning | `S6_Cleaning`, `S6_TextTools` | Duplicates and errors, UPPER/PROPER/LOWER/TRIM/VALUE/LEN/LEFT/RIGHT/MID, SEARCH vs FIND, Go To Special, Text to Columns, CONCATENATE and `&`, Find and Replace |
| 7 Pivot | `S7_Pivot` | Pivot-style tables with filters, date grouping, blanks, custom grouping, conditional formatting |
| 8 Lookup | `S8_Lookup` | VLOOKUP exact and approximate, named ranges, VLOOKUP vs HLOOKUP |
| 9 Conditional aggregation | `S9_CondAgg` | COUNTIF/SUMIF/AVERAGEIF, COUNTIFS/SUMIFS |
| 10 Conditional formatting | `S10_CondFormat` | Highlight rules, colour scales, data bars, icon sets |
| 11 Consolidation | `Jan_2017`, `Feb_2017`, `Mar_2017`, `Consol_Q1`, `Reg_*`, `Consol_Regions`, `S11_Subtotal` | 3-D reference consolidation, regional consolidation, SUBTOTAL |
| 12 Macros | `S12_Macros` | Written answers for all six topics plus the VBA code |
| Data | `Orders`, `Returns`, `People` | Source data used by every formula |

## How to read the sheets

- **Blue text on yellow** is an input or sample value. Change it and the formulas update.
- **Black text** is a formula.
- **Green boxes** are written answers to the "explain" tasks.
- Tasks are labelled **T1-T5** to match the assignment PDF.

## Using the macros

1. Save a copy of the workbook as **Excel Macro-Enabled Workbook (.xlsm)**.
2. Enable the Developer tab: File > Options > Customize Ribbon > tick **Developer**.
3. Press **Alt+F11**, then File > Import File and choose `Macros.bas`.
4. Move the `Workbook_Open` block into **ThisWorkbook** (it will not run automatically from a module), and delete it from the module.

| Macro | Purpose | Run it by |
|---|---|---|
| `FormatSalesTable` | Formats the selected table | Developer > Macros (Alt+F8) |
| `HighlightHighSales` | Highlights values above a threshold you type | Shortcut Ctrl+Shift+H |
| `TotalSelection` | Shows the sum of the selection | A button on the sheet |
| `Workbook_Open` | Welcome message on open | Reopening the file |

Test macros on a copy: they cannot be undone with Ctrl+Z.

## Known limitations

- **4 years, not 5.** The Superstore data only covers 2014-2017, so the growth and CAGR exercise uses 4 years.
- **Assumed times.** The data has dates but no times, so the timestamps in `S4_Dates` (10:30 AM order, 3:45 PM ship) are assumed and labelled.
- **No native PivotTables.** Section 7 uses formula-based summaries with drop-down filters. Steps for building real PivotTables are at the bottom of `S7_Pivot`.
- **UI-only features** (Sort dialog, Text to Columns, Format Painter, Go To Special, the Consolidate dialog, macro recording) are shown as finished results plus exact steps. Consolidation uses 3-D references, which give the same result as the dialog.
- **Two intentional error values** sit in the dirty-data table on `S6_Cleaning` (Sales column, rows 14 and 22) so the "remove errors" task has something to find. Every formula in the workbook is error-free.
- `DATEDIF`, `MINIFS`, `MAXIFS` and `TEXTJOIN` need Excel 2019 or Microsoft 365 for the last three; `DATEDIF` works in all versions.

## Verification

Key results were cross-checked against the raw data: 2017 total sales $733,215.26, 2014-2017 CAGR 14.83%, West Technology sales $251,991.83 (599 lines), Q1 2017 consolidated total equals the Orders total, and the quarter, month and region totals reconcile to the grand total of $2,297,200.86.

## Mentor contact

Questions on the assignment itself can go to the mentor named in the assignment PDF (LinkedIn link on its first page).
