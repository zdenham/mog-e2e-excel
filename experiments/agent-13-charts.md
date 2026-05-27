# Agent 13: Charts

Scope: embedded Excel chart workbooks exported through the Mog upload/edit/export harness, then validated by actual Microsoft Excel on macOS.

## Result

No confirmed chart-specific corrupt-output repro was found in this pass.

All source workbooks used for the final probes were authored by Microsoft Excel itself and opened cleanly in Excel before Mog import. The Mog-exported files also opened cleanly in Excel after the edits below.

## Probe Matrix

| Workbook shape | Mog edit | Excel result after Mog export | Notes |
| --- | --- | --- | --- |
| Embedded clustered column chart over `ChartData!A1:B4` | Header-row edit over `A1:C1` | OK | Chart parts and drawing relationships survived export. |
| Embedded clustered column chart over `ChartData!A1:B4` | Direct series edit of `B2:B3` | OK | Chart formulas still referenced the source range; cached chart values were not fully recalculated, but Excel accepted the file. |
| Embedded clustered column chart over `ChartData!A1:B4` | Direct category edit of `A2:A3` | OK | Stale category caches did not trigger Excel repair. |
| Embedded line chart with title linked to `ChartData!A1` | Header/title source edit | OK | Linked title formula and chart relationship remained valid. |
| Embedded chart over an Excel table source range | Table body edit via `table-data-row` | OK | This avoided the already known table-header corruption class; table body edits plus chart metadata opened cleanly. |
| Embedded chart with value data labels | Direct series edit of `B2:B3` | OK | Data-label chart XML remained valid after export. |

## XML Findings

For the exported chart workbooks, the relevant OOXML parts were retained:

- `xl/charts/chart1.xml`
- `xl/drawings/drawing1.xml`
- `xl/drawings/_rels/drawing1.xml.rels`
- `xl/worksheets/_rels/sheet1.xml.rels`

The chart formulas such as `ChartData!$A$2:$A$4` and `ChartData!$B$2:$B$4` remained present in `xl/charts/chart1.xml`. Mog did not fully refresh chart caches after direct source-cell edits; for example, exported chart XML still contained `c:numCache` / `c:strCache` blocks from the original chart. Actual Excel tolerated those stale caches and opened the workbooks without repair.

## Discarded Candidates

I first tried hand-built minimal OOXML chart fixtures. Excel reported repair dialogs on those source files before Mog touched them, so they were discarded as invalid fixtures.

I also attempted to automate a chart-sheet fixture through Excel AppleScript, but the chart-sheet creation command failed with a parameter error before a valid source workbook was produced. I did not count that as a Mog result.

## Recommendation

Do not add a chart expected-corrupt corpus case from this pass. Chart fixtures would still be useful as expected-OK regression coverage because Mog preserves drawing/chart relationships while editing cells that charts reference.
