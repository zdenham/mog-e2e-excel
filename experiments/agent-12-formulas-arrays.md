# Agent 12/20: Formulas, Arrays, Cached Values, and Calc Chain

Scope: explore corruption candidates around shared formulas, legacy array formulas, dynamic arrays, formula cached values, and `calcChain`.

## Confirmed corrupt repro

### `formula-dynamic-array-spill-block`

Fixture:

- `corpus/dynamic-array-formula.xlsx`

Edit:

- `dynamic-array-spill-cell`
- Sets `A2 = 10`
- Sets `C3 = "blocks spill"` while `C2` contains `=SEQUENCE(4,1,A2,1)`

Result:

- Source workbook opens in real Excel without a repair dialog.
- Mog import/edit/export succeeds.
- Exported workbook triggers Excel repair:
  - `test-results/exports/formula-dynamic-array-spill-block.mog-export.xlsx`
  - Excel dialog: `We found a problem with some content in 'formula-dynamic-array-spill-block.mog-export.xlsx'...`

Exported XML pattern:

```xml
<c r="C2" t="e">
  <f>SEQUENCE(4,1,A2,1)</f>
  <v>#SPILL!</v>
</c>
<c r="C3" t="s"><v>4</v></c>
```

The source has the dynamic-array formula at `C2` and no cell in the spill target `C3`. After the edit, Mog exports a blocked spill range with the dynamic-array formula cached as an error cell. Excel treats this output as repair-worthy.

## Negative coverage added

These scenarios exported and opened cleanly in real Excel after Mog import/edit/export:

- `formula-shared-child-overwrite`: shared formula range with one shared child overwritten.
- `formula-legacy-array-input-edit`: legacy array formula workbook with dependency cells edited. Direct child/anchor array edits are blocked by Mog before export with `PartialArrayWrite`.
- `formula-cache-overwrite`: stale numeric/text/error cached formula values followed by formula edits.
- `formula-calc-chain-overwrite`: workbook with an explicit `xl/calcChain.xml` part followed by formula edits.

## Harness finding

Concurrent Excel checks can produce false readings because Excel is a single shared app instance. I hardened `scripts/check-excel.mjs` so it:

- Forces a clean Excel process before each validation.
- Only attributes repair dialogs to the workbook under test when the dialog text includes that workbook name.
- Polls longer for delayed repair dialogs.

Validation commands run from the Agent 12 worktree:

```bash
npm run corpus:create
REQUIRE_EXCEL=1 E2E_PORT=5174 npx playwright test tests/mog-excel.spec.ts -g "formula-dynamic-array-spill-block"
E2E_PORT=5174 npx playwright test tests/mog-excel.spec.ts -g "formula-"
```

