# Agent 08: Comments and Notes

Scope: legacy Excel cell notes/comments on or near the fixed Mog edit cells `A1`, `B1`, and `C1`.

## Summary

I did not find a confirmed comments/notes-specific corrupt export repro.

The tested exports preserve the expected legacy note package pieces after Mog import, fixed `A1/B1/C1` edit, and UI export:

- `xl/comments1.xml`
- `xl/drawings/vmlDrawing1.vml`
- `xl/worksheets/_rels/sheet1.xml.rels`
- `xl/worksheets/sheet1.xml` `<legacyDrawing r:id="..."/>`

The exported XML still points comments to the edited cell refs, and the worksheet relationships still point to both the comments part and VML drawing part. I did not see the stale-reference pattern that caused the confirmed table-header corruption cases.

Actual Excel validation could not be assigned a reliable final status during this pass because the shared macOS Excel instance was actively being driven by other concurrent agent checks. The existing checker scans all Excel dialogs globally, so running Agent 08 checks during that contention would risk attributing another workbook's repair dialog to these note fixtures.

## Candidate Fixtures Exercised

All candidates were generated under `/tmp` and driven through the real mounted Mog app with the same host flow as the main harness:

1. Upload XLSX through the UI.
2. Apply `smoke-header-edit`, which writes:
   - `A1 = "Mog E2E export smoke test"`
   - `B1 = <ISO timestamp>`
   - `C1 = =LEN(A1)`
3. Export via the UI.
4. Inspect exported XLSX package XML.

### `notes-plain-edited-cells`

Plain legacy notes were attached to `A1`, `B1`, and `C1`, then those cells were overwritten by the fixed edit.

Observed export:

- `comments1.xml` retained comments for `A1`, `B1`, and `C1`.
- `vmlDrawing1.vml` retained matching VML shapes.
- `sheet1.xml.rels` retained both comments and VML relationships.
- No obvious XML mismatch found.

### `notes-rich-anchored-edited-cells`

Rich-text notes with custom margins, protection settings, and mixed `editAs` anchors were attached to `A1`, `B1`, and `C1`.

Observed export:

- Rich note text and formatting survived in `comments1.xml`.
- Comments remained aligned to `A1`, `B1`, and `C1` after cell values/formula changed.
- `legacyDrawing` and rel targets remained present.
- No obvious XML mismatch found.

### `notes-formula-and-string-edited-cells`

Notes were attached to two string cells and one formula cell, then `C1` was overwritten with a new formula through Mog.

Observed export:

- Comment refs remained `A1`, `B1`, and `C1`.
- The edited formula cell and its note coexist in the exported XML.
- No obvious XML mismatch found.

### `notes-protected-unlocked-edited-cells`

The sheet was protected, but `A1`, `B1`, and `C1` were unlocked so Mog could apply the fixed edit. Notes used mixed protection settings.

Observed export:

- Mog edit completed.
- Export succeeded.
- No obvious XML mismatch found in the note/comment relationship structure.

### `notes-merged-range-edited-cells`

A note was attached to merged range anchor `A1` for `A1:C1`, then the fixed edit wrote into `A1`, `B1`, and `C1`.

Observed export:

- Export succeeded.
- This overlaps with Agent 04's merged-cell risk area, but I did not confirm a comments-specific corruption signal.

### `notes-comments-away-from-edits`

Notes were attached away from the edited cells, including `D2`, `A3`, and `C3`.

Observed export:

- Export succeeded.
- Comment refs stayed away from edited cells.
- No obvious XML mismatch found.

## Negative / Blocked Cases

### Protected locked comments

A protected sheet with locked `A1/B1/C1` cells blocks the fixed Mog edit:

```text
Cannot edit cell (0, 0): sheet is protected and cell is locked
```

This is an editability failure, not an export-corruption repro.

## XML Risk Notes

The relevant corruption risk for legacy notes would be a mismatch such as:

- `sheet1.xml` keeps `<legacyDrawing r:id="rIdN"/>` but `sheet1.xml.rels` drops the VML relationship.
- `comments1.xml` contains a comment ref but the worksheet rel to comments is missing.
- `vmlDrawing1.vml` contains note shapes whose cell anchors no longer match `comments1.xml`.
- `[Content_Types].xml` drops the comments or VML overrides/defaults while relationships still reference them.

I did not observe those patterns in the exported candidates inspected during this pass.

## Recommendation

Do not add comments/notes fixtures as expected-corrupt repros yet. They are still useful as expected-OK regression coverage once the Excel checker is serialized, because they exercise legacy comments, VML drawings, and edited cell values together.

Before using concurrent subagents for further Excel validation, add a repo-level lock or workbook-specific dialog matching to `scripts/check-excel.mjs`. Without that, any concurrent repair dialog can be misattributed to whichever agent happens to poll the global Excel UI.
