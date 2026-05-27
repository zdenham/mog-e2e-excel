# Agent 14: drawings, images, and object anchors

## Scope

Focused on image/drawing/object-related workbooks only:

- Inserted PNG images with two-cell and one-cell anchors.
- Edits near anchored images using the Mog programmatic edit harness.
- Multiple images on one sheet.
- Images across multiple sheets.
- Images on hidden sheets.
- Images combined with legacy cell notes.
- Images over merged-cell regions.
- Hyperlinked images.
- Excel-created autoshapes.
- Images near Excel tables as an isolation probe.

## Result

No confirmed image/drawing/object-specific Excel corruption repro was found.

The image-only and shape-only exports opened cleanly in actual Microsoft Excel on macOS after Mog import/export. Excel table combinations produced one apparent image-related failure, but isolation showed the root cause was table metadata, not the drawing/image part.

## Confirmed clean cases

All of these exported through the Mog UI path and then opened in actual Excel with `scripts/check-excel.mjs`:

- `image-two-cell.xlsx`, no edit: `ok`.
- `image-two-cell.xlsx`, edit adjacent cells via `outside-table`: `ok`.
- `image-two-cell.xlsx`, edit nearby formula cells via `formula-cell`: `ok`.
- `image-one-cell.xlsx`, no edit: `ok`.
- `image-multiple.xlsx`, no edit: `ok`.
- `image-multiple.xlsx`, edit adjacent cells via `outside-table`: `ok`.
- `image-multi-sheet.xlsx`, no edit: `ok`.
- `image-hidden-sheet.xlsx`, no edit: `ok`.
- `image-with-note.xlsx`, no edit: `ok`.
- `image-merged-anchor.xlsx`, edit merged/image-adjacent cells via `merged-anchor`: `ok`.
- `image-hyperlinked.xlsx`, no edit: `ok`.
- Excel-created autoshape workbook, no edit: `ok`.
- Excel-created autoshape workbook, edit adjacent cells via `outside-table`: `ok`.
- `image-table-with-totals.xlsx`, no edit: `ok`.
- `table-with-totals-control.xlsx`, no edit: `ok`.

## False positive: image plus table

Initial probe:

- Source workbook: table in `A1:C3`, PNG image anchored at `E1:G6`.
- Source opened cleanly in Excel.
- Mog round-trip export produced an Excel repair dialog.

That looked drawing-related until controls were run:

- `image-table-no-totals.xlsx`, no edit: corrupt.
- `table-no-totals-control.xlsx`, no image, no edit: corrupt.
- `image-table-with-totals.xlsx`, no edit: ok.
- `table-with-totals-control.xlsx`, no image, no edit: ok.

So the corrupt condition follows a no-totals table round-trip, not the inserted image.

## XML notes

For the initial image-plus-table false positive, `xl/drawings/drawing1.xml`, `xl/drawings/_rels/drawing1.xml.rels`, and `xl/media/image1.png` remained structurally equivalent after export. The changed part was `xl/tables/table1.xml`.

The exported no-totals table part added/rewrote table totals metadata, including `totalsRowCount="1"` with `totalsRowShown="1"` on a range that did not contain a totals row. The same repair reproduced when the image was removed, confirming this is a table export issue rather than a drawing relationship or image anchor issue.

The browser console repeatedly logged:

```text
[GridRenderer] syncSceneGraph: getAllObjectBounds returned empty but N objects exist.
```

That appears to be a Mog UI rendering/object-bounds issue. It did not correlate with Excel file corruption in the clean drawing/image cases.

## Commands used

Focused Playwright probes were run serially with actual Excel checks, for example:

```bash
npx playwright test tests/agent-14-drawings-probe.spec.ts --workers=1 --grep "image-multiple|image-multi-sheet|image-hidden-sheet|image-with-note|image-merged-anchor|image-hyperlinked"
npx playwright test tests/agent-14-drawings-probe.spec.ts --workers=1 --grep "excel-shape"
```

The temporary probe spec and generated probe workbooks were removed after this report so no non-repro corpus files are added.
