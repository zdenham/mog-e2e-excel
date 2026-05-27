# Agent 02/20: Table Structured References, Totals Rows, and Calculated Columns

Scope: explore table structured-reference formulas, totals-row formulas, calculated columns, and edits that may leave table metadata inconsistent after Mog export.

Repo context observed:

- Current harness imports a corpus XLSX, runs `applyEdit`, exports via `workbook.exportXlsx()`, then opens the exported file in real Excel through `scripts/check-excel.mjs`.
- Current `applyEdit` writes `A1`, `B1`, and `C1` on the active sheet.
- Existing confirmed failing fixture is `corpus/table-autofilter.xlsx`.
- Its table range is `A1:D5`, so the current edit rewrites table header cells.
- Exported failure keeps `xl/tables/table1.xml` column metadata as `SKU`, `Category`, `Units`, `Amount`, while `xl/worksheets/sheet1.xml` changes header cells `A1:C1`. Excel repairs the workbook.

## Confirmed Risk Pattern

### Header cell values diverge from `tableColumn/@name`

Candidate fixture:

- `table-header-overwrite.xlsx`

Setup:

- One worksheet with an Excel table at `A1:D5`.
- Header row: `SKU`, `Category`, `Units`, `Amount`.
- Table metadata in `xl/tables/table1.xml` has matching `tableColumn name` values.

Edit:

- Write any new value into one or more header cells inside the table range, especially `A1:C1`.
- Include a formula in a header cell as the current harness does with `C1 = LEN(A1)`.

Expected corruption risk:

- `xl/worksheets/sheet1.xml` contains mutated header cells.
- `xl/tables/table1.xml` still contains old `<tableColumn name="...">` values.
- Excel table headers are defined by table metadata and must match the displayed header row. Mismatch can trigger the repair popup.

Existing repro evidence:

- `test-results/exports/table-autofilter.mog-export.xlsx`
- `xl/tables/table1.xml`: `SalesTable` columns remain `SKU`, `Category`, `Units`, `Amount`.
- `xl/worksheets/sheet1.xml`: `A1`, `B1`, `C1` are rewritten by the harness edit.

## High-Value Candidate Scenarios

### 1. Structured references outside the table after header rename

Candidate fixture:

- `table-structured-ref-external-formulas.xlsx`

Setup:

- Table `SalesTable` at `A1:D6`.
- Headers: `SKU`, `Category`, `Units`, `Amount`.
- External formulas outside the table:
  - `F2 = SUM(SalesTable[Units])`
  - `F3 = SUM(SalesTable[Amount])`
  - `F4 = COUNTIF(SalesTable[Category],"Hardware")`
  - `F5 = XLOOKUP("A-100",SalesTable[SKU],SalesTable[Amount])`

Edit variants:

- Overwrite `A1` with `Item ID`.
- Overwrite `C1` with `Qty`.
- Overwrite `D1` with `Revenue`.

XML risks:

- `table1.xml` may retain `name="SKU"`, `name="Units"`, or `name="Amount"`.
- `sheet1.xml` may contain formulas still referencing old structured names.
- If Mog rewrites formulas to the new header text without updating `table1.xml`, formulas and table metadata diverge in the opposite direction.
- If Mog updates only cells and not table metadata, Excel sees header mismatch plus formulas that refer to columns no longer represented by displayed headers.

Likely Excel symptoms:

- Repair popup.
- Repaired formulas may become `#REF!`.
- Table may be downgraded or formulas rewritten during repair.

### 2. Calculated column with `[#This Row]` references

Candidate fixture:

- `table-calculated-column-this-row.xlsx`

Setup:

- Table `SalesTable` at `A1:E6`.
- Headers: `SKU`, `Category`, `Units`, `Unit Price`, `Amount`.
- `Amount` table data cells contain the calculated-column formula:
  - `=[@Units]*[@[Unit Price]]`

Edit variants:

- Overwrite header `C1` from `Units` to `Qty`.
- Overwrite header `D1` from `Unit Price` to `Price`.
- Write a literal value into one calculated-column body cell, such as `E3 = 999`.

XML risks:

- Header rename can leave calculated-column formulas referencing old structured names while the sheet header cells display new names.
- A single body-cell override can break calculated-column consistency. In OOXML this can show up as differing formulas in cells inside the same table column, while `table1.xml` still presents a uniform table column.
- Mog export may serialize formulas as regular cell formulas but preserve stale table metadata.

Likely Excel symptoms:

- Repair popup if header metadata and formula references cannot be reconciled.
- Excel may drop the calculated-column behavior or repair formulas.

### 3. Totals row custom formula references a renamed column

Candidate fixture:

- `table-totals-custom-structured-ref.xlsx`

Setup:

- Table `SalesTable` at `A1:E7` with `totalsRowCount="1"`.
- Headers: `SKU`, `Category`, `Units`, `Amount`, `Margin`.
- Totals row formulas:
  - `C7 = SUBTOTAL(109,SalesTable[Units])`
  - `D7 = SUBTOTAL(109,SalesTable[Amount])`
  - `E7 = SUMPRODUCT(SalesTable[Amount],SalesTable[Margin])/SUM(SalesTable[Amount])`

Edit variants:

- Overwrite `D1` from `Amount` to `Revenue`.
- Overwrite totals cell `D7` with `=SUM(SalesTable[Amount])`.
- Overwrite totals cell `E7` with a formula referencing renamed headers.

XML risks:

- `table1.xml` can store totals metadata using `totalsRowFunction`, `totalsRowFormula`, or extension-backed formula metadata depending on authoring tool.
- `sheet1.xml` stores the visible totals-row formula.
- Mog may preserve one side and rewrite the other side, producing disagreement between totals-row cell formulas and table column totals metadata.
- Header rename makes totals formulas especially fragile because `SalesTable[Amount]` must resolve through `tableColumn/@name`.

Likely Excel symptoms:

- Repair popup.
- Totals row function may be removed, replaced with static value, or changed to `#REF!`.

### 4. Duplicate table header names after edit

Candidate fixture:

- `table-duplicate-header-after-edit.xlsx`

Setup:

- Table `SalesTable` at `A1:D5`.
- Headers: `SKU`, `Category`, `Units`, `Amount`.

Edit variants:

- Set `C1` to `Amount`, creating displayed headers `SKU`, `Category`, `Amount`, `Amount`.
- Set `A1` to blank.
- Set `B1` to a formula.

XML risks:

- Excel tables require unique column names. Excel normally disambiguates duplicates as `Amount`, `Amount2`, etc.
- If `sheet1.xml` displays duplicate or blank headers but `table1.xml` keeps original unique names, metadata and displayed row diverge.
- If Mog updates `table1.xml` naively, it may create duplicate `<tableColumn name="Amount">` entries, which is directly invalid for Excel tables.

Likely Excel symptoms:

- Repair popup.
- Excel may rename headers during recovery.

### 5. Table display name and structured references after table rename-like edits

Candidate fixture:

- `table-name-structured-ref-formulas.xlsx`

Setup:

- Table `SalesTable` at `A1:D6`.
- External formulas reference `SalesTable`.
- Defined name or data validation formula also references `SalesTable[SKU]`.

Edit variants:

- Use UI or kernel APIs, if available, to rename the table to `Sales`.
- If table rename API is not exposed, edit formulas to reference `Sales[Amount]` while table metadata still says `SalesTable`.

XML risks:

- `xl/tables/table1.xml` stores both `name` and `displayName`.
- `sheet1.xml` formulas and `xl/workbook.xml` defined names may refer to the table display name.
- If Mog changes formula text but not table name metadata, Excel cannot resolve structured references.
- If Mog changes only one of `name` or `displayName`, Excel may repair duplicate or invalid table identity state.

Likely Excel symptoms:

- Repair popup or formulas repaired to `#REF!`.

### 6. Row insertion/deletion inside a table with structured refs

Candidate fixture:

- `table-resize-structured-ref.xlsx`

Setup:

- Table `SalesTable` at `A1:E6`.
- External formulas reference full columns and `[#All]`:
  - `=ROWS(SalesTable[#All])`
  - `=SUM(SalesTable[Amount])`
  - `=INDEX(SalesTable[SKU],1)`

Edit variants:

- Insert a data row inside the table.
- Delete the last data row.
- Clear all body values but leave the totals row.

XML risks:

- `table1.xml/@ref`, `<autoFilter ref>`, `worksheet/dimension`, and `sheetData` row extents must agree.
- Formulas using `[#All]`, `[#Data]`, and totals row are sensitive to table extents.
- Mog may change visible rows while preserving stale table range metadata.

Likely Excel symptoms:

- Repair popup if `table1.xml/@ref` points at missing rows/cells or excludes visible table formulas.
- AutoFilter range may be repaired separately from table range.

### 7. Special-character headers in structured references

Candidate fixture:

- `table-special-header-structured-refs.xlsx`

Setup:

- Table headers:
  - `Region`
  - `Unit Price`
  - `Margin %`
  - `Amount [USD]`
  - `Q1/Q2`
- Formulas:
  - `=[@[Unit Price]]*[@[Q1/Q2]]`
  - `=SUM(SalesTable[Amount [USD]])`

Edit variants:

- Rename a special-character header through cell edit.
- Write a formula into a special-character header cell.

XML risks:

- Structured references require escaping for spaces, `%`, brackets, and some punctuation.
- If Mog serializes formula text without Excel-compatible escaping, Excel may repair formulas.
- If header cell text changes but `tableColumn/@name` keeps old escaped/unescaped names, table metadata mismatch remains.

Likely Excel symptoms:

- Repair popup or formulas changed to `#REF!`.

### 8. Hidden totals row toggled or stale totals metadata

Candidate fixture:

- `table-hidden-totals-metadata.xlsx`

Setup:

- Table with `totalsRowCount="1"` initially.
- Totals formulas in the last row.

Edit variants:

- Hide or remove the totals row if UI/kernel supports it.
- Otherwise clear totals row cells with kernel writes.
- Edit body values only, then export.

XML risks:

- `table1.xml/@totalsRowCount`, `table1.xml/@ref`, and actual `sheetData` totals row cells must agree.
- If totals row cells are deleted but `totalsRowCount="1"` remains, Excel may repair the table.
- If totals metadata remains but formulas are absent or shifted, repair may drop totals row configuration.

Likely Excel symptoms:

- Repair popup or silent removal of totals row.

## Recommended Corpus Additions for This Agent

Add these fixtures first because they maximize coverage of distinct table/XML failure modes:

1. `table-structured-ref-external-formulas.xlsx`
2. `table-calculated-column-this-row.xlsx`
3. `table-totals-custom-structured-ref.xlsx`
4. `table-duplicate-header-after-edit.xlsx`
5. `table-special-header-structured-refs.xlsx`

For each fixture, run at least two test modes:

- Import/export with no edit, to isolate pure Mog round-trip corruption.
- Import/edit/export, where the edit targets a table header, calculated-column body cell, or totals-row cell.

The current global `applyEdit` is useful for reproducing header-table mismatch, but it will not test every scenario cleanly because it always edits `A1:C1`. A future harness extension should allow fixture-specific edit scripts so each scenario can mutate the intended table region.

## XML Checks to Automate Before Opening Excel

These checks should not replace the real Excel corruption check, but they can explain failures quickly:

- Compare visible worksheet header row values in each table range against `xl/tables/tableN.xml` `<tableColumn name="...">`.
- Assert table column names are unique and non-empty after resolving shared strings.
- Compare `table/@ref`, `autoFilter/@ref`, `worksheet/dimension/@ref`, and occupied `sheetData` rows.
- Find formulas containing structured references and verify table display names exist in `xl/tables/*.xml`.
- Verify structured-reference column names exist in the target table metadata.
- For totals rows, compare `table/@totalsRowCount`, totals-row cell formulas in `sheetData`, and table column totals metadata.
- For calculated columns, detect formula drift inside one table column after export.

## Hypothesis

The highest-probability corrupt outputs are not generic formula bugs. They are table metadata synchronization bugs: Mog can edit normal worksheet cells inside table-owned regions while leaving the table part unchanged. Excel then receives an XLSX package where worksheet XML, table XML, formulas, and totals metadata describe different versions of the same table.
