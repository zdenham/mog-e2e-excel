# Agent 09: Hyperlinks and Relationship Corruption

## Scope

Explored hyperlink and relationship corruption candidates around Mog import/edit/export:

- External hyperlink relationships.
- Internal workbook hyperlinks.
- Mixed hyperlink targets: HTTPS query strings, `mailto:`, file URLs, and internal locations.
- Stale/unused hyperlink relationships in `xl/worksheets/_rels/sheet1.xml.rels`.
- Edits that overwrite linked cells, blank linked cells, replace linked cells with `HYPERLINK(...)` formulas, and edit adjacent cells.

All corruption confirmation attempts used actual Microsoft Excel on macOS. I did not count XML inspection or ZIP validity as corruption confirmation.

## Result

No hyperlink-specific corruption repro was confirmed.

The main exported XML pattern was stale but tolerated hyperlink relationships. When a Mog edit overwrote `A2`, Mog removed the corresponding `<hyperlink ref="A2" .../>` from `xl/worksheets/sheet1.xml`, but left the now-unused `rId1` relationship in `xl/worksheets/_rels/sheet1.xml.rels`.

Example external-link export after editing `A2/B2`:

```xml
<hyperlinks>
  <hyperlink ref="A3" r:id="rId2" tooltip="Second external HTTPS link"/>
</hyperlinks>
```

```xml
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://openai.com" TargetMode="External"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://github.com" TargetMode="External"/>
```

That leaves `rId1` unused, but I did not capture a target-matched Excel repair popup for this condition.

## Tested Candidates

### External Hyperlinks

Fixture shape:

- `A2` external HTTPS hyperlink.
- `A3` second external HTTPS hyperlink.
- `B2/B3` plain URL strings.

Edits tested:

- Overwrite `A2/B2` with plain strings.
- Replace `A2` with `=HYPERLINK("https://example.com/edited","Edited formula link")`.
- Blank `A2/B2`.
- Edit adjacent `C2/D2`.

Observed XML:

- Edited linked cells drop their `<hyperlink ref="A2".../>` entry.
- Old `rId1` remains in the relationships part as an unused relationship.
- Remaining links preserve valid `r:id` references.

No confirmed target-matched Excel corruption popup.

### Internal Hyperlinks

Fixture shape:

- `A2` links to `#Targets!B2`.
- `A3` links to `#InternalLinks!A1`.
- Separate `Targets` sheet exists.

Edits tested:

- Overwrite `A2/B2` with plain strings.
- Blank `A2/B2`.

Observed XML:

- Edited `A2` hyperlink is removed from `<hyperlinks>`.
- The unused relationship remains:

```xml
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="#Targets!B2" TargetMode="External"/>
```

- The remaining same-sheet hyperlink is preserved as:

```xml
<hyperlink ref="A3" location="#InternalLinks!A1" tooltip="Internal same-sheet link"/>
```

No confirmed target-matched Excel corruption popup.

### Mixed Targets

Fixture shape:

- HTTPS target with query string.
- `mailto:` target.
- `file:///` target.
- Internal `#MixedLinks!A1` target.

Edit tested:

- Overwrite `A2/B2` with plain strings.

Observed XML:

- Removed the edited `A2` hyperlink from sheet XML.
- Left stale unused `rId1`.
- Remaining `mailto:`, file URL, and internal links stayed referentially valid.

One actual Excel run returned `OK` for this case before later concurrent Excel validator interference.

### Explicit Stale Relationship

Fixture shape:

- Same as external hyperlinks.
- Added an extra unused relationship manually:

```xml
<Relationship Id="rIdAgent09Unused" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/stale-unused-relationship" TargetMode="External"/>
```

Edit tested:

- Overwrite `A2/B2` with plain strings.

Observed XML:

- Both the original unused `rId1` and explicit unused `rIdAgent09Unused` remained.
- No missing `r:id`, duplicate `r:id`, or hyperlink element pointing at a missing relationship was observed.

No confirmed target-matched Excel corruption popup.

## Harness Finding

Concurrent agents were using the same global Microsoft Excel instance. The current checker scans all Excel window text, so it can report corruption for another workbook while validating the current file.

Observed false-positive messages named unrelated workbooks, including:

- `conditional-formatting-databar-empty-color.mog-export.xlsx`
- `conditional-formatting-databar.xlsx`
- `table-autofilter-header-formula.mog-export.xlsx`

Recommendation: serialize actual Excel validation through a repo-level lock, and update the checker so repair dialogs are matched to the target exported workbook name before returning `corrupt`.

## Highest-Risk Follow-Ups

- Confirm hyperlink cases again after Excel validation is serialized.
- Add a target-matched repair-dialog checker before treating hyperlink stale-rel cases as confirmed.
- Explore valid inputs with multi-cell hyperlink refs if we can produce them from Excel itself, then edit only part of the referenced range through Mog.
- Explore hyperlinks inside table header cells separately, because table metadata is already a confirmed corruption class and may combine with hyperlink relationship cleanup.
