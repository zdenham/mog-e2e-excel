# Agent 06/20: Data Validation Findings

Scope: data validations, dropdown lists, formulas, whole-number constraints, and edits that invalidate validation ranges after Mog import/export.

## Result

No confirmed data-validation-only corruption repro was found.

I tested generated XLSX fixtures by:

1. Opening the original in real Microsoft Excel on macOS.
2. Uploading into the Mog harness through Playwright.
3. Applying the current harness edit (`A1`, `B1`, `C1`).
4. Exporting through the UI.
5. Opening the exported XLSX in real Microsoft Excel.
6. Inspecting exported `xl/worksheets/sheet*.xml`, `xl/workbook.xml`, and `xl/tables/table*.xml`.

All completed data-validation cases opened cleanly in Excel after Mog export. One first-pass Excel check timed out for `dv-inline-list-a1.xlsx`, but rerunning the same exported file returned `OK`; I treated that as automation noise, not corruption.

## Candidate Fixtures Tested

These should still be added to the corpus because they exercise distinct data-validation serialization paths, even though they did not corrupt in this pass.

| Fixture | Purpose | Excel after Mog export | Notable XML observation |
| --- | --- | --- | --- |
| `dv-inline-list-a1.xlsx` | Inline dropdown list on edited cell `A1` with stop error message. | OK on rerun | `formula1` exported as text content `"Open,Closed,Hold"`; original used `&quot;...&quot;`. Both are valid XML text forms. |
| `dv-range-list-a1-source-b1-b3.xlsx` | Dropdown on `A1` sourced from `$B$1:$B$3`; edit changes `B1`. | OK | `<formula1>$B$1:$B$3</formula1>` preserved. |
| `dv-overwritten-source-list-a1-from-b1-b3.xlsx` | Same-sheet list source intentionally invalidated by harness edit to `B1`. | OK | Validation formula remains syntactically valid even though source values change. |
| `dv-whole-number-a1.xlsx` | Whole-number validation on edited `A1`; edit writes text. | OK | Mog omitted default `errorStyle="stop"`, but Excel accepts it. |
| `dv-decimal-b1.xlsx` | Decimal range validation on edited `B1`; edit writes timestamp text. | OK | `type="decimal"` with formula bounds preserved. |
| `dv-date-c1.xlsx` | Date validation on edited `C1`; edit writes formula. | OK | Date bounds preserved as serials (`46023`, `46387`). |
| `dv-text-length-a1.xlsx` | Text-length validation on edited `A1`. | OK | `operator="lessThanOrEqual"` preserved. |
| `dv-custom-c1-len.xlsx` | Custom validation formula on edited `C1`. | OK | XML escaping for `<` preserved as `&lt;`. |
| `dv-multi-cells-a1-c1.xlsx` | Single validation rule over `A1:C1`, all edited by harness. | OK | `sqref="A1:C1"` and count preserved. |
| `dv-hidden-sheet-direct-ref.xlsx` | Dropdown source on hidden sheet using direct sheet reference. | OK | `Choices!$A$1:$A$3` preserved. |
| `dv-named-range-hidden-list-a1.xlsx` | Dropdown source through workbook defined name on hidden sheet. | OK | `<definedName name="StatusChoices">Choices!$A$1:$A$3</definedName>` preserved. |
| `dv-quoted-sheet-list-a1.xlsx` | Dropdown source on hidden sheet with spaces in sheet name. | OK | Quoted reference `'Lookup Values'!$A$1:$A$3` preserved. |
| `dv-custom-cross-sheet-countif.xlsx` | Custom validation formula referencing another sheet. | OK | `COUNTIF(... )&gt;0` escaped correctly. |
| `dv-structured-ref-list-a1.xlsx` | Dropdown source using table structured reference. | OK | `ChoiceTable[Choice]` preserved and table XML still present. |
| `dv-operator-notbetween-a1.xlsx` | Non-default whole-number operator. | OK | `operator="notBetween"` preserved. |
| `dv-prompt-error-message-a1.xlsx` | Validation prompt and non-default warning style. | OK | `showInputMessage`, `promptTitle`, `prompt`, `errorStyle="warning"`, `errorTitle`, and `error` preserved. |
| `dv-table-and-validation-overlap.xlsx` | Data validations inside a table; harness edits table headers. | OK | Table column metadata did not match edited header cells, but Excel did not flag this variant. |
| `dv-table-totals-header-overwrite.xlsx` | Data validations inside a totals-row table; harness edits table headers. | OK | Header mismatch was present, but Excel did not flag this variant. This is table-risk, not validation-risk. |

## Highest-Risk Corpus Additions

Add these first if we want focused coverage without bloating the corpus:

1. `data-validation-inline-list.xlsx`
   - Validation on `A1`: `type=list`, `formula1="Open,Closed,Hold"`, `showErrorMessage=true`, `errorStyle=stop`.
   - Harness edit writes an invalid value to `A1`.

2. `data-validation-source-range-edited.xlsx`
   - Validation on `A1`: `type=list`, `formula1=$B$1:$B$3`.
   - Harness edit changes both validated cell `A1` and list source cell `B1`.

3. `data-validation-hidden-named-range.xlsx`
   - Hidden sheet `Choices`, workbook defined name `StatusChoices=Choices!$A$1:$A$3`.
   - Validation on `A1`: `type=list`, `formula1=StatusChoices`.

4. `data-validation-custom-cross-sheet.xlsx`
   - Validation on `A1`: `type=custom`, `formula1=COUNTIF(Choices!$A$1:$A$3,A1)>0`.
   - Exercises formula XML escaping and cross-sheet references.

5. `data-validation-structured-ref-list.xlsx`
   - Validation on `A1`: `type=list`, `formula1=ChoiceTable[Choice]`.
   - Exercises data validation plus table relationship preservation.

6. `data-validation-prompt-error-style.xlsx`
   - Validation on `A1` with prompt/error metadata and non-default `errorStyle=warning`.
   - Catches attribute loss, not necessarily corruption.

## XML Checks To Automate

These checks should run on exported XLSX files before invoking Excel. They should not replace the real Excel open check; they should explain failures faster.

### Data Validation Integrity

- For each worksheet XML, if `<dataValidations count="N">` exists, verify `N` equals the number of child `<dataValidation>` elements.
- Every `<dataValidation>` must have a non-empty `sqref`.
- Every `sqref` range must parse as one or more valid A1 references/ranges.
- `type="list"` and `type="custom"` must have a non-empty `<formula1>`.
- Operators `between` and `notBetween` must have both `<formula1>` and `<formula2>`.
- Formula text must not contain `#REF!`, malformed XML entities, or unescaped raw `<`, `>`, or `&` characters.
- Re-exported inline list formulas should normalize by XML text value, not raw bytes. `&quot;Open,Closed&quot;` and `"Open,Closed"` are equivalent once parsed as element text.

### Reference Resolution

- For data-validation formulas that are bare names, verify the name exists in `xl/workbook.xml` under `<definedNames>`.
- For formulas referencing quoted sheet names, verify the referenced sheet exists in `xl/workbook.xml`.
- For formulas referencing table structured refs, verify the table exists in `xl/tables/table*.xml` and the referenced column name exists in `<tableColumns>`.
- For same-sheet or cross-sheet range references, verify the target sheet exists and the referenced cells/ranges are syntactically valid.

### Attribute Preservation

These are not corruption checks, but they catch fidelity regressions:

- Preserve non-default `operator` values.
- Preserve non-default `errorStyle` values such as `warning` and `information`.
- Preserve `showInputMessage`, `promptTitle`, `prompt`, `showErrorMessage`, `errorTitle`, and `error`.
- Default `errorStyle="stop"` may be omitted without corruption because Excel treats stop as the default.

### Table Interaction Guard

Data-validation tests that also involve tables should run the table header check, because table corruption can look like a validation failure:

- For each table `ref`, compare worksheet header-row cell display strings with `xl/tables/table*.xml` `<tableColumn name="...">`.
- Verify table column count matches the table range width.
- Verify table `ref` and `autoFilter ref` are consistent with totals-row settings.

## Example Observed Export XML

Inline list validation after Mog export:

```xml
<dataValidations count="1">
  <dataValidation sqref="A1" type="list" showErrorMessage="1" errorTitle="Invalid status" error="Pick a listed value.">
    <formula1>"Open,Closed,Hold"</formula1>
  </dataValidation>
</dataValidations>
```

Named hidden-range validation after Mog export:

```xml
<definedNames>
  <definedName name="StatusChoices">Choices!$A$1:$A$3</definedName>
</definedNames>
<dataValidations count="1">
  <dataValidation sqref="A1" type="list" showErrorMessage="1">
    <formula1>StatusChoices</formula1>
  </dataValidation>
</dataValidations>
```

Custom cross-sheet validation after Mog export:

```xml
<dataValidations count="1">
  <dataValidation sqref="A1" type="custom" showErrorMessage="1">
    <formula1>COUNTIF(Choices!$A$1:$A$3,A1)&gt;0</formula1>
  </dataValidation>
</dataValidations>
```

## Recommendation

I would add the six highest-risk fixtures above as passing coverage for this category. For confirmed corruption hunting, the next sub-agents should focus more on tables, totals rows, auto filters, defined names, conditional formatting, merged regions, drawings, charts, comments, and pivot/cache structures. In this pass, data validation looked well-preserved and did not independently trigger Excel repair dialogs.
