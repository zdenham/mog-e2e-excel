# Agent 04: Merged Cells, Table Interactions, and Merge Metadata

## Scope

Explored corruption candidates involving merged-cell metadata after Mog import, edit, export, and real Excel validation:

- horizontal, vertical, and 2D merged ranges
- edits to the anchor and non-anchor cells inside merged ranges
- multiple merged ranges on the edited row
- merged formula anchors
- autofilter ranges adjacent to merged headings
- Excel tables below merged headings
- Excel tables with merged cells inside header, data, and totals rows

The current committed harness exposes only one programmatic edit recipe:

```ts
await sheet.setCell('A1', 'Mog E2E export smoke test');
await sheet.setCell('B1', new Date().toISOString());
await sheet.setCell('C1', '=LEN(A1)');
```

So this pass focused on fixtures where `A1:C1` overlaps or sits near merged ranges.

## Temporary Cases Run

Temporary source files were generated under `/tmp/mog-agent04-merged-cells/inputs`, `/tmp/mog-agent04-overlap-probe`, and `/tmp/mog-agent04-table-merge-totals`. Exports were produced through the real app upload/edit/export flow and checked with `scripts/check-excel.mjs`, which opens the output in actual Microsoft Excel on macOS.

| Case | Merge / XML Focus | Edit? | Result |
| --- | --- | --- | --- |
| `merge-title-a1-d1.xlsx` | `mergeCell ref="A1:D1"`; harness edits anchor `A1` and non-anchors `B1:C1` | no | Opened in Excel |
| `merge-title-a1-d1.xlsx` | same | yes | Opened in Excel on isolated rerun; first matrix pass timed out in automation |
| `merge-block-b1-c2.xlsx` | `mergeCell ref="B1:C2"`; harness edits anchor/non-anchor row cells | no/yes | Opened in Excel |
| `merge-vertical-a1-a4.xlsx` | `mergeCell ref="A1:A4"`; harness edits anchor plus adjacent row cells | no/yes | Opened in Excel |
| `table-below-merged-title.xlsx` | merged title `A1:D1`, table starts at `A3` | no/yes | Opened in Excel |
| `merge-adjacent-autofilter.xlsx` | merged title `A1:C1`, worksheet `autoFilter="A2:C5"` | no/yes | Opened in Excel |
| `merge-formula-anchor.xlsx` | merged `A1:C1` with formula anchor before edit | no/yes | Opened in Excel |
| `two-merged-ranges-top-row.xlsx` | `A1:B1` and `C1:D1` both intersect fixed edit | no/yes | Opened in Excel |
| `table-header-merge-a1-b1.xlsx` | table header row has merge `A1:B1` | no/yes | Opened in Excel |
| `table-data-merge-a2-b2.xlsx` | table data row has merge `A2:B2` | no/yes | Opened in Excel |
| `table-total-merge-a4-b4.xlsx` | table totals row has merge `A4:B4` | no/yes | Opened in Excel |
| `table-header-merge-a1-b1-totals.xlsx` | table header merge with totals functions | no/yes | Opened in Excel |
| `table-data-merge-a2-b2-totals.xlsx` | table data merge with totals functions | no/yes | Opened in Excel |
| `table-header-merge-b1-c1-totals.xlsx` | table header merge over edited `B1:C1` with totals functions | no/yes | Opened in Excel |

No confirmed corrupt output was found for merged-cell-specific cases under the current fixed edit.

## XML Observations

The highest-signal observation is that Mog can export values/formulas in non-anchor cells while preserving the original merge range. For `merge-title-a1-d1.xlsx` after the fixed edit, `xl/worksheets/sheet1.xml` contained:

```xml
<c r="A1" s="1" t="s"><v>1</v></c>
<c r="B1" s="1" t="s"><v>2</v></c>
<c r="C1" s="1"><f>LEN(A1)</f><v>25</v></c>
<mergeCells count="1">
  <mergeCell ref="A1:D1"/>
</mergeCells>
```

This is semantically suspicious because `B1` and `C1` are inside the merged `A1:D1` range, but Excel did not report corruption. It likely ignores or hides those non-anchor values while the merge remains active.

For `table-header-merge-a1-b1.xlsx` after the fixed edit, Mog preserved both the table and merge metadata:

```xml
<table ... ref="A1:C3">
  <autoFilter ref="A1:C3"/>
  <tableColumns count="3">
    <tableColumn id="1" name="A" totalsRowLabel="Total"/>
    <tableColumn id="2" name="B"/>
    <tableColumn id="3" name="C"/>
  </tableColumns>
</table>
```

Worksheet row 1 changed to the harness values while the merge remained:

```xml
<c r="A1" t="s"><v>2</v></c>
<c r="B1" t="s"><v>3</v></c>
<c r="C1"><f>LEN(A1)</f><v>25</v></c>
<mergeCell ref="A1:B1"/>
```

Excel still opened this output. That suggests the already confirmed `table-autofilter.xlsx` corruption is more specific than "table header differs from visible cells"; totals-row formulas, table shape, calculated metadata, autofilter details, or exact table serialization may matter.

## Candidate Fixture / Edit Scenarios

1. `merge-non-anchor-only-edit`
   - Fixture: `A1:D1` merged with stable anchor text.
   - Edit: set only `B1` and/or `C1`, leaving `A1` unchanged.
   - XML risk: non-anchor cells gain values while merge metadata remains. Current harness cannot isolate this because it always edits `A1`.
   - Priority: high as a merge-specific metadata invariant.

2. `merge-anchor-clear-with-non-anchor-values`
   - Fixture: `A1:D1` merged.
   - Edit: clear `A1`, set `B1 = "Replacement"`.
   - XML risk: merge anchor is blank but non-anchor contains visible/intended value. Excel may repair or silently drop data.
   - Priority: high.

3. `merge-formula-non-anchor`
   - Fixture: `A1:D1` merged with text anchor.
   - Edit: set `C1 = =LEN(A1)` only.
   - XML risk: formula cell inside merged range is not the anchor. This is closer to a pure OOXML inconsistency than the current edit, which also overwrites the anchor.
   - Priority: high.

4. `merge-unmerge-after-edit`
   - Fixture: `A1:D1` merged.
   - Edit: set `A1:C1`, then unmerge `A1:D1` through Mog API if supported.
   - XML risk: stale `<mergeCell ref="A1:D1"/>` may remain after worksheet cells are expanded.
   - Priority: high if Mog exposes merge/unmerge APIs.

5. `merge-create-over-existing-values`
   - Fixture: unmerged `A1:D1` with distinct values in each cell.
   - Edit: merge `A1:D1` through Mog API if supported.
   - XML risk: stale non-anchor `<c>` elements may remain after creating the merge.
   - Priority: high if Mog exposes merge APIs.

6. `merge-over-table-header-existing-values`
   - Fixture: Excel table at `A1:D5` with valid headers and totals row.
   - Edit: merge `A1:B1` or `B1:C1` through Mog API.
   - XML risk: Excel table header cells are merged after import; table metadata may stay unchanged.
   - Priority: high, but needs a merge API edit rather than only `setCell`.

7. `table-header-edit-plus-merge`
   - Fixture: table at `A1:D5` plus a separate merged range elsewhere on row 1 or directly in the header if Excel accepts the source.
   - Edit: rename only one table header cell.
   - XML risk: combines the known table-header corruption class with merge metadata. Current table+merge probes opened, but they did not reproduce the exact `table-autofilter.xlsx` table shape.
   - Priority: medium-high.

8. `merge-with-array-or-spill-formula-anchor`
   - Fixture: merged `A1:D1` whose anchor contains a dynamic-array formula or normal formula result.
   - Edit: overwrite a non-anchor cell, or overwrite the anchor with a formula while preserving the merge.
   - XML risk: formula/spill metadata conflicts with merged-cell range.
   - Priority: medium.

9. `merge-across-hidden-columns`
   - Fixture: `A1:D1` merged with `B:C` hidden.
   - Edit: set `B1`/`C1` or unhide/hide columns if supported.
   - XML risk: merge metadata plus hidden column metadata may preserve stale non-anchor cells.
   - Priority: medium.

10. `merge-crossing-freeze-pane-or-print-title`
    - Fixture: merged top-row title crossing a frozen pane and/or print-title row.
    - Edit: set non-anchor cells or unmerge.
    - XML risk: stale merge metadata combines with sheet views and `_xlnm.Print_Titles`.
    - Priority: medium.

11. `merge-style-divergence`
    - Fixture: merged `A1:D1` where non-anchor cells have distinct border/fill styles before merging, or styles are applied after import.
    - Edit: set anchor/non-anchor values and styles if style API exists.
    - XML risk: stale style IDs on hidden non-anchor cells inside a merge; likely visual/data bug more than corruption.
    - Priority: medium-low.

12. `merge-over-comments-hyperlinks`
    - Fixture: merged range with comments or hyperlinks on anchor and non-anchor cells.
    - Edit: overwrite or clear anchor/non-anchor cells.
    - XML risk: comments/hyperlinks may remain attached to non-anchor cells inside a merge and conflict with worksheet relationships.
    - Priority: medium-low.

## Suggested Parent Harness Extension

Add case-specific edit modes to `window.__mogHarness.applyEdit(mode)` so the corpus can target merged-cell invariants directly:

- `nonAnchorOnly`: set `B1` and `C1`, do not touch `A1`
- `clearAnchorSetNonAnchor`: clear `A1`, set `B1`
- `formulaNonAnchor`: set `C1 = =LEN(A1)` only
- `unmergeAfterEdit`: set cells inside `A1:D1`, then unmerge
- `mergeAfterValues`: start unmerged with `A1:D1` populated, then merge
- `mergeTableHeader`: merge cells inside a table header after import

The current fixed edit is useful as a broad smoke test, but it cannot prove whether Mog handles merge operations correctly because it never creates or removes merges and cannot isolate non-anchor-only writes.

## Bottom Line

No merged-cell-specific corrupt export was confirmed in this pass. The strongest corpus additions are still worth adding because they capture suspicious XML that Excel currently tolerates: values/formulas in non-anchor cells inside preserved merged ranges. The highest-risk next step is to extend the harness with targeted merge and non-anchor edit modes, then rerun these fixtures through real Excel validation.
