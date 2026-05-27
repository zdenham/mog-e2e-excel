# Agent 01 - Table Header / Totals / Adjacent Resize Candidates

Scope: Excel table corruption cases where Mog exports worksheet cell edits but may leave `xl/tables/tableN.xml` metadata stale. Parent should add fixtures/parameterized edit modes and run actual Excel validation.

## Existing Confirmed Failure

Fixture: `corpus/table-autofilter.xlsx`

Current harness edit:

- `A1 = "Mog E2E export smoke test"`
- `B1 = <ISO timestamp>`
- `C1 = =LEN(A1)`

Observed source XML:

- `xl/tables/table1.xml` has `ref="A1:D5"`, `totalsRowCount="1"`, `autoFilter ref="A1:D4"`.
- Table columns are `SKU`, `Category`, `Units`, `Amount`.
- `xl/worksheets/sheet1.xml` row 1 initially contains shared strings for those same four headers.

Observed corrupt export XML:

- `xl/tables/table1.xml` remains:
  - `ref="A1:D5"`
  - `autoFilter ref="A1:D4"`
  - `tableColumn id="1" name="SKU"`
  - `tableColumn id="2" name="Category"`
  - `tableColumn id="3" name="Units" totalsRowFunction="sum"`
  - `tableColumn id="4" name="Amount" totalsRowFunction="sum"`
- `xl/worksheets/sheet1.xml` row 1 changes to:
  - `A1` shared string index for `"Mog E2E export smoke test"`
  - `B1` shared string index for timestamp
  - `C1` formula `LEN(A1)` with cached value `25`
  - `D1` still shared string index for `"Amount"`

Likely invariant violated: Excel table header cells must correspond to table column metadata. Mog updates worksheet cells but does not update `tableColumn/@name` values or normalize formula/non-string header cells.

Secondary observations:

- Exported `autoFilter` loses the original empty `<filterColumn>` children. This does not look like the primary failure because the header-restored patched export reportedly opened.
- Exported `sharedStrings.xml` had `count="11"` and `uniqueCount="13"` with 13 `<si>` entries. Also not the primary failure based on the header-restored patched export.

## Candidate Scenarios

1. `table-header-single-string-overwrite`
   - Fixture: table with headers `SKU, Category, Units, Amount`, totals row enabled.
   - Edit: set only `A1 = "Renamed SKU"`.
   - Expected XML risk: worksheet header `A1` differs from `tableColumn id="1" name="SKU"`.
   - Priority: highest; isolates header-name mismatch without formula/timestamp noise.

2. `table-header-all-string-overwrite`
   - Fixture: same table.
   - Edit: set `A1:D1 = Product, Segment, Quantity, Revenue`.
   - Expected XML risk: all worksheet headers differ while table metadata stays old.
   - Priority: high; verifies whether any header edit triggers repair, even with valid unique text headers.

3. `table-header-formula`
   - Fixture: same table.
   - Edit: set `C1 = =LEN(A1)`.
   - Expected XML risk: table header cell becomes a formula cell while `tableColumn id="3" name="Units"` remains static.
   - Priority: high; current repro includes this, but isolate it from string renames.

4. `table-header-number`
   - Fixture: same table.
   - Edit: set `C1 = 123`.
   - Expected XML risk: worksheet header is numeric/non-string; table metadata still says `Units`. Excel normally coerces table headers to text, so stale metadata may repair.
   - Priority: high.

5. `table-header-blank`
   - Fixture: same table.
   - Edit: set `B1 = ""` or clear `B1` if API supports clearing.
   - Expected XML risk: blank worksheet table header with `tableColumn id="2" name="Category"` still present. Excel table headers cannot be blank and may auto-repair to `Column1`.
   - Priority: high.

6. `table-header-duplicate`
   - Fixture: same table.
   - Edit: set `B1 = "SKU"`.
   - Expected XML risk: worksheet table headers contain duplicate visible names while table metadata still contains unique names. Excel may repair duplicate headers to `SKU2` or reject the table.
   - Priority: high.

7. `table-header-special-chars`
   - Fixture: table with plain headers.
   - Edit: set `A1 = "SKU[old]"`, `B1 = "Category#All"`, `C1 = "Units, Net"`.
   - Expected XML risk: structured-reference-sensitive names in worksheet row but stale old names in table metadata.
   - Priority: medium; useful if special escaping is mishandled in formulas/table refs.

8. `table-data-edit-control`
   - Fixture: same table.
   - Edit: set `A2 = "A-101"`, `C2 = 21`, no header/totals edits.
   - Expected XML risk: none; worksheet data changes under stable table metadata.
   - Priority: control. If this fails, problem is broader than header metadata.

9. `table-totals-label-overwrite`
   - Fixture: same table with totals row enabled.
   - Edit: set `A5 = "Grand Total"`.
   - Expected XML risk: worksheet totals label differs from `tableColumn id="1" totalsRowLabel="Total"`.
   - Priority: medium. Excel may tolerate or rewrite, but this mirrors the header-name stale metadata class.

10. `table-totals-formula-overwrite`
    - Fixture: same table with totals row enabled.
    - Edit: set `C5 = =SUM(C2:C4)` and/or `D5 = 13050`.
    - Expected XML risk: worksheet totals-row cell no longer matches `totalsRowFunction="sum"`, which expects `SUBTOTAL(109,SalesTable[...])` semantics.
    - Priority: medium-high.

11. `table-adjacent-right-header-row`
    - Fixture: same table.
    - Edit: set `E1 = "Outside Header"`.
    - Expected XML risk: worksheet dimension expands beyond table (`A1:E5` or similar) while table ref stays `A1:D5`; should be valid if relationship metadata remains stable.
    - Priority: control/edge. If corrupt, boundary/dimension handling is implicated.

12. `table-adjacent-below-totals-row`
    - Fixture: same table.
    - Edit: set `A6:D6` to a new row immediately below totals row.
    - Expected XML risk: worksheet dimension expands below table but `table ref="A1:D5"` and `autoFilter ref="A1:D4"` stay unchanged. Should be valid as outside-table data; corruption would suggest table boundary/export range handling issues.
    - Priority: medium.

13. `table-adjacent-below-no-totals-row`
    - Fixture: similar table with no totals row, table ref `A1:D4`.
    - Edit: set `A5:D5` immediately below the table.
    - Expected XML risk: Excel may interpret adjacent row as candidate table expansion in UI, but OOXML is valid if table ref remains `A1:D4`. Useful contrast with totals-row fixture.
    - Priority: medium.

14. `table-header-with-filtered-state`
    - Fixture: table with active filter criteria in `autoFilter/filterColumn`.
    - Edit: rename a filtered header cell.
    - Expected XML risk: stale `tableColumn/@name` plus stale filter column criteria. Export already drops empty filter column nodes; non-empty criteria may expose additional repair cases.
    - Priority: high if Mog supports/imports filter criteria.

15. `table-structured-formula-after-header-rename`
    - Fixture: table plus a formula outside the table using `=SUM(SalesTable[Units])`.
    - Edit: set `C1 = "Qty"`.
    - Expected XML risk: worksheet header says `Qty`, table metadata and structured formula still reference `Units`. Excel may repair table metadata or formula references.
    - Priority: high; finds downstream formula fallout from stale header metadata.

## Suggested Parent Harness Extension

Add an edit-mode parameter exposed through `window.__mogHarness.applyEdit(mode)` so each scenario can reuse the same fixture with targeted edits. Keep the current all-header edit as one case, then add isolated modes for single-header, formula-header, blank-header, duplicate-header, totals-label, totals-formula, data-control, right-adjacent, and below-adjacent.

Expected first validation order:

1. Single header string overwrite.
2. Header formula only.
3. Header blank and duplicate.
4. Totals formula/label.
5. Adjacent right/below controls.
6. Filtered-state and structured-reference fixtures.
