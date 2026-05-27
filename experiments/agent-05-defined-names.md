# Agent 05: Defined Names, Print Areas, and Name-Dependent Formulas

## Scope

Explored defined-name scenarios that could leave stale `xl/workbook.xml` after Mog import, programmatic edit, and export:

- workbook-scoped range names
- sheet-scoped duplicate names using `localSheetId`
- formulas that reference defined names
- formula/constant names
- dynamic names such as `OFFSET(...)`
- discontiguous range names
- hidden-sheet references
- quoted sheet names with spaces
- print areas and print titles (`_xlnm.Print_Area`, `_xlnm.Print_Titles`)
- stale `#REF!` names

The current committed harness exposes only one programmatic edit recipe:

```ts
await sheet.setCell('A1', 'Mog E2E export smoke test');
await sheet.setCell('B1', new Date().toISOString());
await sheet.setCell('C1', '=LEN(A1)');
```

So this pass tested defined-name inputs under that mutation. Sheet rename/name-update scenarios are listed below as follow-up corpus candidates because the current UI/harness does not expose a per-case edit recipe.

## Temporary Cases Run

Temporary source files were generated under `/tmp/mog-agent-05-defined-names/inputs` and exported to `/tmp/mog-agent-05-defined-names/outputs` through the real app upload/edit/export flow.

| Case | Input XML Focus | Result |
| --- | --- | --- |
| `dn-workbook-range-overlap-a1.xlsx` | `RevenueName -> Inputs!$A$1:$A$3`, overlapping edited `A1` | Opened in Excel on isolated check |
| `dn-print-area-overlap-a1.xlsx` | `_xlnm.Print_Area`, `_xlnm.Print_Titles`, both overlapping row 1 | Opened in Excel on isolated check |
| `dn-sheet-scoped-duplicates.xlsx` | duplicate `LocalRate` names with `localSheetId=0` and `localSheetId=1` | Opened in Excel |
| `dn-formula-constant-names.xlsx` | constant/formula names: `TaxRate`, `RevenueBase`, `GrowthRate` | Opened in Excel |
| `dn-hidden-sheet-reference.xlsx` | defined name points to a hidden sheet | Opened in Excel |
| `dn-quoted-sheet-space.xlsx` | defined name points to `'Input Sheet'!$B$2` | Opened in Excel |
| `dn-dynamic-offset.xlsx` | `DynamicRange -> OFFSET(Dynamic!$A$2,0,0,3,1)` | Opened in Excel |
| `dn-discontiguous-ranges.xlsx` | `TwoBlocks -> Blocks!$A$2:$A$3,Blocks!$B$2:$B$3` | Opened in Excel on isolated rerun |
| `dn-local-print-area-two-sheets.xlsx` | two local print areas with separate `localSheetId`s | Opened in Excel on isolated rerun |
| `dn-ref-error-stale-name.xlsx` | `StaleName -> #REF!` | Opened in Excel |

No confirmed corrupt output was found for defined names using the current fixed `A1/B1/C1` edit. Two early corrupt readings were false positives: Excel dialog text referenced unrelated concurrent-agent files such as `table-no-totals-header-row.mog-export.xlsx` and `table-offset-header-row.mog-export.xlsx`. Clean reruns of the named files opened successfully.

## XML Observations

For all generated cases, Mog preserved the `<definedNames>` block in `xl/workbook.xml`. Examples:

```xml
<definedName name="LocalRate" localSheetId="0">East!$B$2</definedName>
<definedName name="LocalRate" localSheetId="1">West!$B$2</definedName>
```

```xml
<definedName name="_xlnm.Print_Area" localSheetId="0">One!$A$1:$B$3</definedName>
<definedName name="_xlnm.Print_Area" localSheetId="1">Two!$A$1:$B$3</definedName>
```

```xml
<definedName name="DynamicRange">OFFSET(Dynamic!$A$2,0,0,3,1)</definedName>
```

The only observed rewrite was benign quoting normalization for a simple print area:

```diff
- <definedName name="_xlnm.Print_Area" localSheetId="0">&apos;Report&apos;!$A1:$C4</definedName>
+ <definedName name="_xlnm.Print_Area" localSheetId="0">Report!$A1:$C4</definedName>
```

Excel accepted the exported file after this rewrite.

## Highest-Risk Follow-Up Scenarios

These are the scenarios most likely to expose defined-name corruption, but they require the harness to support case-specific Mog API edits instead of only the fixed cell edit:

| Scenario | Proposed Mog Mutation | XML Risk |
| --- | --- | --- |
| Rename sheet with workbook-scoped names | `workbook.sheets.rename('Inputs', 'Inputs Renamed')` | `definedName` formulas in `xl/workbook.xml` may still reference the old sheet name, while formulas and workbook sheets use the new one |
| Rename sheet with sheet-scoped names | rename a sheet that has `localSheetId`-scoped names | scope may remain by index but reference text may point to old sheet name |
| Delete/reorder sheets with local names | `workbook.sheets.remove(...)` or `workbook.sheets.move(...)` | `localSheetId` is index-based; stale IDs can point names/print areas at the wrong sheet |
| Copy sheet with local names | `workbook.sheets.copy('East', 'East Copy')` | duplicate local names may be missing, incorrectly workbook-scoped, or assigned the wrong `localSheetId` |
| Update a named range reference | `workbook.names.update('RevenueName', { reference: 'Inputs!$A$2:$A$4' })` | formulas may recalc but `xl/workbook.xml` can preserve the old range |
| Remove a name used by formulas | `workbook.names.remove('RevenueName')` | formulas should become `#NAME?`; exporter could leave dangling workbook name metadata |
| Create names from labels after editing headers | `workbook.names.createFromSelection(...)` after changing row 1 | generated names may conflict with existing names or invalid Excel name syntax |
| Print area after sheet rename/delete | rename/delete a sheet with `_xlnm.Print_Area` | Excel is sensitive to stale print-area `localSheetId` and sheet references |

## Harness Note

Parallel sub-agents using the same local Excel instance can contaminate corruption checks. The current AppleScript checker scans all Excel windows, so a dialog opened by another workbook can be attributed to the file currently under test. A robust next step is to make the checker verify that the dialog workbook name matches the expected basename, or to serialize real-Excel validation across agents.
