# Agent 11: Hidden rows/columns/sheets, outline/grouping, freeze panes

Scope: explore whether hidden rows, hidden columns, hidden sheets, outline/grouping metadata, split panes, or freeze panes can produce Mog XLSX exports that actual Microsoft Excel flags as corrupt.

## Result

No confirmed corrupt-output repro was found for this feature family in this pass.

I did not add a corpus repro or expected-corrupt Playwright scenario because every apparent corruption dialog observed during the run named a different workbook from another concurrent agent. Those were treated as false positives rather than Agent 11 findings.

## Fixtures exercised

The temporary runner generated these XLSX variants and sent them through the real app flow: upload into Mog, apply a programmatic edit, export through Mog, then validate with actual Excel on macOS.

| Candidate | Edit | Result |
| --- | --- | --- |
| `hidden-row-and-column.xlsx` | `hidden-row-cell` | One clean Excel open observed; later checks were contaminated by foreign dialogs. |
| `hidden-sheet-formula-reference.xlsx` | `formula-cell` | Clean Excel open observed. |
| `very-hidden-sheet-reference.xlsx` | `formula-cell` | Inconclusive because Excel automation timed out / foreign dialogs were present. |
| `outline-rows-collapsed.xlsx` | `hidden-row-cell` | Inconclusive because foreign dialogs were present. |
| `outline-columns-collapsed.xlsx` | `hidden-row-cell` | Inconclusive because foreign dialogs were present. |
| `freeze-pane-top-row.xlsx` | `smoke-header-edit` | One clean Excel open observed in the first pass; later checks timed out. |
| `freeze-pane-row-column.xlsx` | `hidden-row-cell` | One clean Excel open observed in the first pass; later checks were contaminated by foreign dialogs. |
| `split-pane-hidden-row.xlsx` | `hidden-row-cell` | Did not reach export; Mog threw a WASM trap during the edit. |
| `hidden-first-sheet.xlsx` | `formula-cell` | Inconclusive because Excel automation timed out. |
| `hidden-outline-freeze-combined.xlsx` | `hidden-row-cell` | Inconclusive because Excel automation timed out / foreign dialogs were present. |

## XML observations

For the exported hidden row/column fixture, Mog preserved the expected hidden row and column metadata:

```xml
<cols><col min="2" max="2" width="9" hidden="1" customWidth="1"/></cols>
<row r="3" spans="1:4" hidden="1" x14ac:dyDescent="0.25">
```

For the exported freeze-pane + hidden row/column fixture, Mog preserved the pane metadata and hidden row/column metadata:

```xml
<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>
<selection pane="bottomRight" activeCell="B2" sqref="B2"/>
<col min="2" max="2" width="9" hidden="1" customWidth="1"/>
<row r="3" spans="1:4" hidden="1" x14ac:dyDescent="0.25">
```

For the hidden-sheet formula fixture, Mog preserved workbook sheet visibility:

```xml
<sheet name="Lookup" sheetId="1" state="hidden" r:id="rId4"/>
<sheet name="Visible" sheetId="2" r:id="rId5"/>
```

These observations point away from a simple stale-metadata corruption class for hidden rows, hidden columns, hidden sheets, and frozen panes.

## Separate finding: split pane edit trap

The `split-pane-hidden-row.xlsx` candidate loaded far enough for the harness to attempt the edit, but the edit failed before export:

```text
ModuleTrappedError: [compute_begin_undo_group] WASM module trapped
(originating: [compute_get_viewport_binary] WASM trap during compute_get_viewport_binary: unreachable)
```

This is not an Excel corruption repro because no Mog export was produced. It is worth tracking separately as a Mog import/edit/render stability issue for split-pane workbooks.

## Harness note

Concurrent agents using the same macOS Excel instance can create false corrupt readings. The false dialogs observed in this pass named workbooks such as:

- `cf-expression-intersects-fixed-edit.mog-export.xlsx`
- `cf-table-totals-intersects-totals-edit.mog-export.xlsx`
- `table-no-totals-header-row.mog-export.xlsx`
- `conditional-formatting-databar-empty-color.mog-export.xlsx`

The Excel validator should only classify a corruption dialog as a hit when the dialog text names the workbook currently under test, or the corpus search should serialize all real-Excel validation.

## Recommended next cases

Add these only after the Excel validator is made workbook-name-aware or the real-Excel phase is serialized:

- Split-pane fixture with no edit, to determine whether import/export alone succeeds.
- Split-pane fixture with a visible-cell edit after the viewport is stable.
- Hidden row/column fixture with formulas whose dependencies are hidden.
- Outline row/column fixture with summary rows above and summary columns left.
- Very-hidden sheet fixture with defined names pointing into the very-hidden sheet.
