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
```

`npm run test:e2e` runs the browser import/export flow and performs the Excel check when Excel is available. `npm run test:e2e:excel` requires Excel and fails if Excel is missing or automation cannot run.

The macOS Excel check uses `osascript` plus `System Events` to inspect Excel windows for corruption or repair dialogs. macOS may require Accessibility permission for the terminal/Codex app that runs the tests.

Current local result: Excel accepts `simple-formulas`, `formats-dates-merged`, and `multi-sheet-references` exports, but reports a repair/corruption dialog for `table-autofilter.mog-export.xlsx`.
