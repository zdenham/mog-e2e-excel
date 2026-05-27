# Agent 10: Workbook and Sheet Protection

Date: 2026-05-27

Scope: Explore whether workbook protection, worksheet protection, protected ranges, hidden/outline metadata on protected sheets, or protected sheets with filters/tables produce corrupt Mog XLSX exports when validated by actual Microsoft Excel on macOS.

## Summary

No confirmed protection-only corruption repro was found.

The clearest protection-specific behaviors were:

- Mog rejected kernel edits to locked cells on protected sheets before export.
- Mog exported allowed edits to unlocked cells on protected sheets cleanly in the confirmed cases.
- Workbook structure protection survived an allowed edit and opened cleanly in Excel in the confirmed case.
- A protected table with unlocked header cells did reproduce Excel's repair dialog, but the XML signature is the already-known table-header mismatch: visible worksheet header cells changed while `xl/tables/tableN.xml` retained the original `tableColumn/@name` values. Protection is not required for that failure.

## Tested Candidates

Temporary fixtures were generated under `/tmp/mog-agent10-protection` and `/tmp/mog-agent10-protection-wave2`.

Confirmed clean in actual Excel:

- `protected-unlocked-inputs.xlsx` with `formula-cell`: protected sheet, edited unlocked formula/input cells. Excel opened exported file OK.
- `workbook-structure-protected.xlsx` with `smoke-header-edit`: workbook structure protection plus hidden support sheet. Excel opened exported file OK.
- `protected-merged-unlocked.xlsx` with `merged-anchor`: protected sheet with merged title and explicitly unlocked edited cells. Excel opened exported file OK.
- `protected-hidden-unlocked.xlsx` with `hidden-row-cell`: protected sheet with hidden row, hidden column, and outline metadata; edited explicitly unlocked cells. Excel opened exported file OK.

Blocked before export:

- `protected-locked-cells.xlsx` with `smoke-header-edit`: Mog rejected editing locked `A1`.
- `protected-hidden-outline.xlsx` with `hidden-row-cell`: Mog rejected editing locked `A3`.
- `protected-merged-cells.xlsx` with `merged-anchor`: Mog rejected editing locked `A1`.
- `protected-autofilter.xlsx` with `autofilter-header-row`: Mog rejected editing locked `A1`.
- `protected-table-header.xlsx` with `table-header-row-values`: Mog rejected editing locked `A1`.

Confirmed corrupt but not protection-specific:

- `protected-table-header-unlocked.xlsx` with `table-header-row-values`: Excel showed the repair dialog after export.
- This is the same table metadata failure as the existing table-header repro. The edited cells are in the table header row, and the exported worksheet values no longer match `xl/tables/tableN.xml` table-column names.

Inconclusive because concurrent Excel automation from other agents was active:

- `protected-header-unlocked.xlsx`
- `protected-autofilter-header-unlocked.xlsx`
- `protected-ranges-unlocked.xlsx`
- `workbook-and-sheet-protected-unlocked.xlsx`

The checker became unreliable while several independent `osascript` Excel validators were open at the same time. Some checks timed out or returned empty output. This matches the race noted by Agent 05: the current checker scans global Excel windows and can be affected by dialogs from other agent runs.

## XML Risk Notes

Protection metadata itself did not appear to be a corruption trigger in the confirmed clean cases. The important XML nodes to watch in future regression tests are:

- `xl/worksheets/sheetN.xml` `<sheetProtection>`
- cell-level `<xf applyProtection="1">` style records and locked/unlocked protection styles
- `xl/workbook.xml` `<workbookProtection>`
- optional `<protectedRanges>`
- protected sheets combined with tables, especially `xl/tables/tableN.xml` `table/@ref`, `autoFilter/@ref`, and `tableColumn/@name`

## Recommendation

Do not add a new protection corpus repro yet. Add protection coverage as expected-OK cases once the Excel checker can serialize validation or match dialogs to the specific workbook under test. The only confirmed repair-popup case from this slice should be tracked under the existing table-header/table-metadata corruption family, not as a separate protection failure.
