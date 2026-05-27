# Agent 07/20: Conditional Formatting Findings

Scope: conditional-formatting corruption candidates, especially rules whose ranges intersect the fixed Mog edit of `A1`, `B1`, and `C1`, plus table interactions.

## Result

No confirmed conditional-formatting-only corrupt output was found.

I tested generated XLSX fixtures through the real Mog import/edit/export path and real Microsoft Excel validation on macOS. After clearing stale Excel dialogs and rerunning targeted checks, the valid conditional-formatting exports opened cleanly in Excel.

## Cases Tested

| Candidate | Edit | Excel result after Mog export | Notes |
| --- | --- | --- | --- |
| `cf-expression-intersects-fixed-edit.xlsx` | `smoke-header-edit` | OK | Expression rule over `A1:C5`; edited cells intersected the CF range. |
| `cf-cell-rules-intersects-fixed-edit.xlsx` | `smoke-header-edit` | OK | `cellIs` and `top10` rules over `A1:C4`. |
| `cf-visual-rules-intersects-fixed-edit.xlsx` | `smoke-header-edit` | Not a valid repro | Initial repair popup traced to a data-bar source rule missing a legacy `<color>` element; that source can itself be invalid depending on Excel's handling. With a valid source color, Mog preserved the color and Excel opened the export. |
| `cf-contains-text-time-intersects-fixed-edit.xlsx` | `smoke-header-edit` | OK on clean rerun | First pass looked corrupt, but a clean Excel session opened it OK. Export added `operator="containsText" text=""`, but this alone did not trigger repair in the reduced fixture. |
| `cf-table-data-intersects-data-edit.xlsx` | `table-data-row` | OK | Conditional formatting inside table body range did not corrupt when editing table data cells. |

## Important Harness Finding

Concurrent Excel validation can create false corrupt results. The current checker scans all Excel windows and dialogs, so a repair popup from another agent's workbook can be attributed to the file currently under test. I saw this repeatedly with stale dialogs from unrelated workbooks.

For future multi-agent runs, Excel validation should either be serialized or the checker should only treat a repair dialog as a failure when the dialog text includes the workbook basename being checked.

## XML Observations

Valid expression, `cellIs`, `top10`, and table-body CF rules were preserved well enough for Excel to open the Mog exports.

Two suspicious-but-not-confirmed patterns are worth tracking:

- Data bars authored without a legacy `<color>` can produce `<color rgb=""/>` after export. Patch-testing `rgb=""` to `rgb="FF638EC6"` made the tested file open cleanly, but the corresponding source workbook was not a clean valid input, so I did not add it as a corpus repro.
- `containsText` rules may export with `text=""` while retaining a formula such as `NOT(ISERROR(SEARCH("Open",A1)))`. A reduced fixture with that export shape opened cleanly in Excel after stale dialogs were cleared.

## Recommended Follow-Up

Do not add a conditional-formatting failure to the corpus from this pass.

Before more concurrent corruption hunting, harden `scripts/check-excel.mjs` to match repair dialogs against the target workbook name. Otherwise agents will keep finding false positives whenever another workbook's repair dialog is still open.
