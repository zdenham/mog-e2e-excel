import { expect, test } from '@playwright/test';
import type ExcelJS from 'exceljs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import {
  checkWithCom,
  exportAfterEdits,
  type CellEdit,
  writeWorkbook,
} from './agent-test-utils';

type Candidate = {
  id: string;
  issue: string;
  configure: (workbook: ExcelJS.Workbook) => void;
  edits: CellEdit[];
};

const runDir = path.join(tmpdir(), `mog-e2e-agent08-${process.pid}-${Date.now()}`);

function candidatePaths(id: string) {
  const dir = path.join(runDir, id);
  return {
    sourcePath: path.join(dir, `${id}.xlsx`),
    exportedPath: path.join(dir, `${id}.mog-export.xlsx`),
  };
}

const candidates: Candidate[] = [
  {
    id: 'table-beside-merge-non-anchor-only',
    issue: 'non-anchor edit in merged cells adjacent to a table part',
    configure(workbook) {
      const sheet = workbook.addWorksheet('TableBesideMerge');
      sheet.addTable({
        name: 'SalesTable',
        ref: 'A1',
        headerRow: true,
        totalsRow: false,
        columns: [
          { name: 'SKU' },
          { name: 'Category' },
          { name: 'Units' },
          { name: 'Revenue' },
        ],
        rows: [
          ['A-100', 'Hardware', 4, 400],
          ['B-200', 'Services', 2, 800],
          ['C-300', 'Hardware', 7, 1050],
        ],
      });
      sheet.mergeCells('F1:H2');
      sheet.getCell('F1').value = 'Side note';
    },
    edits: [
      { address: 'G2', value: 'non-anchor note' },
    ],
  },
  {
    id: 'merged-title-above-table-non-anchor-only',
    issue: 'non-anchor edit in a merged title above a table',
    configure(workbook) {
      const sheet = workbook.addWorksheet('TitleAboveTable');
      sheet.mergeCells('A1:D1');
      sheet.getCell('A1').value = 'Sales report';
      sheet.addTable({
        name: 'SalesTable',
        ref: 'A3',
        headerRow: true,
        totalsRow: false,
        columns: [
          { name: 'SKU' },
          { name: 'Category' },
          { name: 'Units' },
          { name: 'Revenue' },
        ],
        rows: [
          ['A-100', 'Hardware', 4, 400],
          ['B-200', 'Services', 2, 800],
          ['C-300', 'Hardware', 7, 1050],
        ],
      });
    },
    edits: [
      { address: 'B1', value: 'non-anchor title edit' },
    ],
  },
];

for (const candidate of candidates) {
  test(`agent08 ${candidate.id}`, async ({ page }, testInfo) => {
    const { sourcePath, exportedPath } = candidatePaths(candidate.id);

    await writeWorkbook(sourcePath, candidate.configure);

    const sourceCheck = checkWithCom(sourcePath);
    expect(sourceCheck.status, `source workbook must be valid: ${sourceCheck.message}`).toBe('ok');

    await exportAfterEdits(page, sourcePath, exportedPath, candidate.edits);

    const exportedCheck = checkWithCom(exportedPath);
    testInfo.annotations.push({
      type: 'excel-check',
      description: `${candidate.issue}: ${exportedCheck.status}: ${exportedCheck.message}`,
    });

    expect(exportedCheck.status, exportedCheck.message).toBe('corrupt');
  });
}
