# Agent 03/20 - Worksheet AutoFilter Scenarios

Scope: standalone worksheet `autoFilter` only, without Excel table parts. This pass explored filtered columns, hidden filter buttons, sort state metadata, and harness edits inside or near filtered ranges.

## Result

No isolated standalone AutoFilter scenario produced an Excel repair dialog.

The strongest finding is that this class does not reproduce the current `table-autofilter.xlsx` failure. Mog preserves worksheet-level `autoFilter` XML almost verbatim after import/edit/export, and Excel accepts the output even when the harness overwrites `A1:C1` inside the AutoFilter header row.

That differs from the known table failure because standalone worksheet filters do not carry table column header names in `xl/tables/table*.xml`. There is no separate table metadata for Excel to compare against the edited header cells.

## Method

Temporary workbooks were generated outside the repo and run through the existing harness:

1. Upload XLSX into the Mog React app.
2. Apply the current programmatic harness edit:
   - `A1 = "Mog E2E export smoke test"`
   - `B1 = <ISO timestamp>`
   - `C1 = =LEN(A1)`
3. Export from Mog.
4. Open exported XLSX in actual Microsoft Excel on macOS via `scripts/check-excel.mjs`.
5. Inspect exported `xl/worksheets/sheet1.xml`.

During the first run, several repair-dialog detections were contaminated by other concurrently running agents; the dialog text named other files. After clearing stale dialogs and rerunning this agent's exported files in isolation, all valid worksheet AutoFilter candidates opened cleanly.

## Candidate Matrix

These scenarios were exercised and opened successfully in Excel after Mog export:

| Scenario | Input AutoFilter shape | Export XML behavior | Excel result |
|---|---|---|---|
| Plain AutoFilter with header overwrite | `<autoFilter ref="A1:D7"/>` | Preserved | OK |
| Single value filter | `filterColumn colId="1"` with `<filter val="Hardware"/>` | Preserved | OK |
| Multi-value filter | `filterColumn colId="1"` with two filter values | Preserved | OK |
| Hidden filter button | `filterColumn colId="2" hiddenButton="1"` | Preserved | OK |
| Numeric custom filter | `customFilter operator="greaterThan" val="8"` | Preserved | OK |
| AND custom filter | two custom filters with `and="1"` | Preserved | OK |
| Top 10 filter | `<top10 val="3"/>` | Preserved | OK |
| Dynamic filter | `<dynamicFilter type="aboveAverage"/>` | Preserved | OK |
| Sort state | `<sortState ref="A2:D7"><sortCondition .../>` | Preserved | OK |
| Date group filter | `<dateGroupItem year="2026" month="1" .../>` | Preserved | OK |
| Two filtered columns | value filter plus custom numeric filter | Preserved | OK |
| Edit near, not inside, filtered range | AutoFilter starts at `A3:D9`; harness edits `A1:C1` | Preserved | OK |

Representative exported XML for the header-overwrite cases:

```xml
<autoFilter ref="A1:D7">
  <filterColumn colId="1">
    <filters>
      <filter val="Hardware"/>
    </filters>
  </filterColumn>
</autoFilter>
```

and the edited header formula cell:

```xml
<c r="C1" s="1"><f>LEN(A1)</f><v>25</v></c>
```

Excel accepted this combination for standalone AutoFilters.

## Candidates To Keep For Corpus Coverage

These are not failing repros today, but they are the best AutoFilter-specific regression fixtures to add if the corpus should cover this surface:

1. `worksheet-autofilter-value-filter-header-overwrite.xlsx`
2. `worksheet-autofilter-hidden-button-header-overwrite.xlsx`
3. `worksheet-autofilter-dynamic-filter-header-overwrite.xlsx`
4. `worksheet-autofilter-sort-state-header-overwrite.xlsx`
5. `worksheet-autofilter-date-group-header-overwrite.xlsx`
6. `worksheet-autofilter-lower-range-edit-nearby.xlsx`

Expected current result for all six: exported file should open in Excel without repair.

## Repair Candidates Not Recommended Yet

These shapes are likely to trigger Excel repair dialogs if they appear in exported XML, but they should not be added as failing corpus files until a clean source workbook can be produced and verified in Excel first:

1. `autoFilter ref` includes a missing/blank header column after export, while filter metadata references that column.
2. `filterColumn colId` points outside the `autoFilter ref` width.
3. `autoFilter ref` is serialized as whole columns, such as `A:D`, instead of a bounded range.
4. `sortState ref` or `sortCondition ref` points outside the `autoFilter ref`.
5. AutoFilter extension metadata references differential formats, icon sets, or dynamic filters that are not present elsewhere in the package.

Those would be useful fuzz targets, but they are not confirmed Mog-generated corruptions from valid standalone AutoFilter input in this pass.

## Conclusion

Standalone worksheet AutoFilter is probably not the root cause of the current Excel repair popup. The known failure mode should stay focused on Excel tables: table header cells can be edited by Mog while `xl/tables/table*.xml` keeps stale `tableColumn name` metadata. Worksheet-level `autoFilter` does not have the same duplicated header-name contract, so the same edit pattern opens cleanly in Excel.
