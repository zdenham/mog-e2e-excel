import { expect, test, type Page, type TestInfo } from '@playwright/test';
import ExcelJS from 'exceljs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import {
  checkWithCom,
  exportAfterEdits,
  type CellEdit,
  type ExcelCheckResult,
  writeWorkbook,
} from './agent-test-utils';

type Candidate = {
  id: string;
  issue: string;
  configure: (workbook: ExcelJS.Workbook) => Promise<void> | void;
  edits: CellEdit[];
};

test.setTimeout(240_000);

const maxExcelTextLength = 32767;
const overlongTextLength = 33000;
const runRoot = path.join(
  tmpdir(),
  `mog-agent16-shared-strings-${process.pid}-${Date.now()}`,
);

function addBaseStringsSheet(workbook: ExcelJS.Workbook) {
  const sheet = workbook.addWorksheet('Strings');
  sheet.addRows([
    ['Key', 'Value', 'Notes'],
    ['plain-marker', 'Plain source string', 'source row'],
    ['edit-target', 'Original edit target', 'edits land here'],
  ]);
  sheet.getColumn(1).width = 24;
  sheet.getColumn(2).width = 34;
  sheet.getColumn(3).width = 28;
}

function addMaxLengthSharedStringSheet(workbook: ExcelJS.Workbook) {
  addBaseStringsSheet(workbook);
  const sheet = workbook.getWorksheet('Strings')!;
  sheet.getCell('B2').value = 'L'.repeat(maxExcelTextLength);
}

const candidates: Candidate[] = [
  {
    id: 'edit-plain-to-overlong-string',
    issue: 'plain shared-string cell is edited beyond Excel cell text length limits',
    configure: addBaseStringsSheet,
    edits: [
      { address: 'B3', value: 'O'.repeat(overlongTextLength) },
    ],
  },
  {
    id: 'edit-max-shared-string-to-overlong',
    issue: 'max-length shared string is edited beyond Excel cell text length limits',
    configure: addMaxLengthSharedStringSheet,
    edits: [
      { address: 'B2', value: 'M'.repeat(overlongTextLength) },
    ],
  },
];

function candidatePaths(id: string) {
  const dir = path.join(runRoot, id);
  return {
    sourcePath: path.join(dir, `${id}.xlsx`),
    exportedPath: path.join(dir, `${id}.mog-export.xlsx`),
  };
}

function checkWithComSignal(filePath: string): ExcelCheckResult {
  let result = checkWithCom(filePath);
  for (let attempt = 2; attempt <= 3 && result.status === 'error'; attempt += 1) {
    result = checkWithCom(filePath);
  }
  return result;
}

async function runCandidate(
  page: Page,
  testInfo: TestInfo,
  candidate: Candidate,
) {
  const { sourcePath, exportedPath } = candidatePaths(candidate.id);

  await writeWorkbook(sourcePath, candidate.configure);

  const sourceCheck = checkWithComSignal(sourcePath);
  testInfo.annotations.push({
    type: 'agent16-source-com',
    description: `${candidate.issue}: ${sourceCheck.status}: ${sourceCheck.message}`,
  });
  expect(sourceCheck.status, `source workbook must be valid: ${sourceCheck.message}`).toBe('ok');

  await exportAfterEdits(page, sourcePath, exportedPath, candidate.edits);

  const exportedCheck = checkWithComSignal(exportedPath);
  testInfo.annotations.push({
    type: 'agent16-export-com',
    description: `${candidate.issue}: ${exportedCheck.status}: ${exportedCheck.message}`,
  });
  expect(exportedCheck.status, exportedCheck.message).toBe('corrupt');
}

for (const candidate of candidates) {
  test(`agent16 ${candidate.id}`, async ({ page }, testInfo) => {
    await runCandidate(page, testInfo, candidate);
  });
}
