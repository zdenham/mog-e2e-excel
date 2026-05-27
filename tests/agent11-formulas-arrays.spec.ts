import { expect, test } from '@playwright/test';
import ExcelJS from 'exceljs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import {
  checkWithCom,
  exportAfterEdits,
  type CellEdit,
  writeWorkbook,
} from './agent-test-utils';

type WorkbookConfigurer = (workbook: ExcelJS.Workbook) => Promise<void> | void;

type Candidate = {
  id: string;
  issue: string;
  distinctness: string;
  configure: WorkbookConfigurer;
  edits: CellEdit[];
};

function candidatePaths(id: string) {
  const dir = path.join(tmpdir(), 'mog-agent11-formulas-arrays', id);
  return {
    sourcePath: path.join(dir, `${id}.xlsx`),
    exportedPath: path.join(dir, `${id}.mog-export.xlsx`),
  };
}

function addDynamicSeedRows(sheet: ExcelJS.Worksheet) {
  sheet.addRows([
    ['Seed', 'Label', 'Spill A', 'Spill B', 'Spill C', 'Spill D'],
    [1, 'alpha', null, null, null, null],
    [2, 'beta', null, null, null, null],
    [3, 'alpha', null, null, null, null],
    [4, 'gamma', null, null, null, null],
    [5, 'beta', null, null, null, null],
  ]);
}

function addHorizontalSequenceWorkbook(workbook: ExcelJS.Workbook) {
  const sheet = workbook.addWorksheet('HorizontalSpill');
  addDynamicSeedRows(sheet);
  sheet.getCell('C2').value = { formula: 'SEQUENCE(1,4,A2,1)', result: 1 };
}

function addTwoDimensionalSequenceWorkbook(workbook: ExcelJS.Workbook) {
  const sheet = workbook.addWorksheet('TwoDimSpill');
  addDynamicSeedRows(sheet);
  sheet.getCell('C2').value = { formula: 'SEQUENCE(2,2,A2,1)', result: 1 };
}

function addUniqueWorkbook(workbook: ExcelJS.Workbook) {
  const sheet = workbook.addWorksheet('UniqueSpill');
  addDynamicSeedRows(sheet);
  sheet.getCell('C2').value = { formula: 'UNIQUE(B2:B6)', result: 'alpha' };
}

const candidates: Candidate[] = [
  {
    id: 'formula-dynamic-array-horizontal-spill-block',
    issue: 'horizontal dynamic array spill range blocked after import/export',
    distinctness: 'horizontal spill shape, not the existing vertical SEQUENCE spill-block case',
    configure: addHorizontalSequenceWorkbook,
    edits: [
      { address: 'D2', value: 'blocks horizontal spill' },
      { address: 'A2', value: 10 },
    ],
  },
  {
    id: 'formula-dynamic-array-2d-spill-block',
    issue: 'two-dimensional dynamic array spill range blocked after import/export',
    distinctness: '2x2 spill footprint instead of the existing single-column SEQUENCE spill-block case',
    configure: addTwoDimensionalSequenceWorkbook,
    edits: [
      { address: 'D3', value: 'blocks 2d spill' },
      { address: 'A2', value: 10 },
    ],
  },
  {
    id: 'formula-dynamic-array-unique-text-spill-block',
    issue: 'UNIQUE dynamic array spill range blocked after import/export',
    distinctness: 'uses UNIQUE over repeated text values rather than SEQUENCE over numbers',
    configure: addUniqueWorkbook,
    edits: [
      { address: 'C3', value: 'blocks unique spill' },
      { address: 'B6', value: 'delta' },
    ],
  },
];

for (const candidate of candidates) {
  test(`agent11 ${candidate.id}`, async ({ page }, testInfo) => {
    test.setTimeout(240_000);
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
    testInfo.annotations.push({
      type: 'distinctness',
      description: candidate.distinctness,
    });

    expect(exportedCheck.status, exportedCheck.message).toBe('corrupt');
  });
}
