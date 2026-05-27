# Mog XLSX E2E Excel Harness

This repo mounts the public `@mog-sdk/spreadsheet-app` React embed in a small Vite app and drives the high-level XLSX path:

1. Upload a corpus workbook.
2. Apply a programmatic edit through Mog workbook APIs.
3. Export through the host UI.
4. Open the exported file in the real Microsoft Excel app on macOS and flag corruption dialogs.

## Commands

```bash
npm run corpus:create
npm run dev
npm run test:e2e
npm run test:e2e:excel
npm run test:pipeline -- --parallel 4
npm run test:pipeline -- --parallel 4 --com
```

`npm run test:e2e` runs the browser import/export flow and performs the Excel check when Excel is available. `npm run test:e2e:excel` requires Excel and fails if Excel is missing or automation cannot run.

The macOS Excel check uses `osascript` plus `System Events` to inspect Excel windows for corruption or repair dialogs. macOS may require Accessibility permission for the terminal/Codex app that runs the tests. Before opening a workbook, the checker copies it into the stable repo-local `excel-validation/` folder so Excel does not repeatedly ask for access to random temp directories.

`npm run test:pipeline -- --parallel N` runs Mog import/edit/export in parallel, writes exported artifacts under `test-results/pipeline/exports`, then validates those artifacts with actual Excel. On macOS, Excel validation remains serialized because Excel and AppleScript dialogs are process-global. On Windows, `--com` switches validation to the COM/UI Automation checker and uses `--parallel N` as the COM validator concurrency.

Current confirmed failure classes are table header metadata mismatches, dynamic array spill edits, and a workbook metadata relationship that imports but fails Mog export before any XLSX is produced.
