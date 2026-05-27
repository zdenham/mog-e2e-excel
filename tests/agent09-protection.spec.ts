import { expect, test, type TestInfo } from '@playwright/test';
import ExcelJS from 'exceljs';
import { mkdirSync } from 'node:fs';
import path from 'node:path';
import {
  checkWithCom,
  exportAfterEdits,
  repoRoot,
  type CellEdit,
  writeWorkbook,
} from './agent-test-utils';

function unlockCells(sheet: ExcelJS.Worksheet, addresses: string[]) {
  for (const address of addresses) {
    sheet.getCell(address).protection = { locked: false };
  }
}

async function protectSheet(sheet: ExcelJS.Worksheet) {
  await sheet.protect('agent09', {
    selectLockedCells: false,
    selectUnlockedCells: true,
    formatCells: false,
    formatColumns: false,
    formatRows: false,
    insertColumns: false,
    insertRows: false,
    insertHyperlinks: false,
    deleteColumns: false,
    deleteRows: false,
    sort: true,
    autoFilter: true,
    pivotTables: false,
  });
}

const edits: CellEdit[] = [
  { address: 'A3', value: 'Edited SKU' },
  { address: 'C3', value: 42 },
  { address: 'D3', value: 4242 },
];

test('agent09 protected table body edits corrupt exported workbook', async ({
  page,
}, testInfo: TestInfo) => {
  const artifactDir = path.join(repoRoot, 'test-results', 'agent09-p09');
  mkdirSync(artifactDir, { recursive: true });
  const sourcePath = path.join(artifactDir, 'source.xlsx');
  const exportedPath = path.join(artifactDir, 'exported.xlsx');

  await writeWorkbook(sourcePath, async (workbook) => {
    const sheet = workbook.addWorksheet('TableBody');
    sheet.addTable({
      name: 'ProtectedBodyTable',
      ref: 'A1',
      headerRow: true,
      totalsRow: false,
      style: { theme: 'TableStyleMedium5', showRowStripes: true },
      columns: [
        { name: 'SKU', filterButton: true },
        { name: 'Category', filterButton: true },
        { name: 'Units', filterButton: true },
        { name: 'Amount', filterButton: true },
      ],
      rows: [
        ['A-100', 'Hardware', 20, 1200],
        ['B-205', 'Software', 5, 3750],
        ['C-310', 'Services', 9, 8100],
      ],
    });
    unlockCells(sheet, ['A3', 'C3', 'D3']);
    await protectSheet(sheet);
  });

  const sourceResult = checkWithCom(sourcePath);
  testInfo.annotations.push({
    type: 'excel-com-source',
    description: `${sourceResult.status}: ${sourceResult.message}`,
  });
  expect(sourceResult.status, sourceResult.message).toBe('ok');

  await exportAfterEdits(page, sourcePath, exportedPath, edits);

  const exportResult = checkWithCom(exportedPath);
  testInfo.annotations.push({
    type: 'excel-com-export',
    description: `${exportResult.status}: ${exportResult.message}`,
  });

  expect(exportResult.status, exportResult.message).toBe('corrupt');
});
