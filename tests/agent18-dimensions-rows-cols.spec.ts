import { expect, test } from '@playwright/test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
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
  postprocess?: (buffer: Uint8Array) => Promise<Uint8Array> | Uint8Array;
};

test.setTimeout(240_000);

const runRoot = path.join(tmpdir(), `mog-agent18-dimensions-${process.pid}-${Date.now()}`);

async function rewriteZip(
  buffer: Uint8Array,
  rewriter: (zip: JSZip) => Promise<void> | void,
) {
  const zip = await JSZip.loadAsync(buffer);
  await rewriter(zip);
  return zip.generateAsync({ type: 'uint8array' });
}

async function patchSheetXml(
  buffer: Uint8Array,
  patcher: (xml: string) => string,
  sheetPath = 'xl/worksheets/sheet1.xml',
) {
  return rewriteZip(buffer, async (zip) => {
    const sheet = zip.file(sheetPath);
    if (!sheet) {
      throw new Error(`Missing worksheet part: ${sheetPath}`);
    }
    zip.file(sheetPath, patcher(await sheet.async('string')));
  });
}

function replaceRequired(xml: string, pattern: RegExp, replacement: string) {
  if (!pattern.test(xml)) {
    throw new Error(`Pattern not found: ${pattern}`);
  }
  return xml.replace(pattern, replacement);
}

function setDimension(ref: string) {
  return (buffer: Uint8Array) =>
    patchSheetXml(buffer, (xml) =>
      replaceRequired(xml, /<dimension ref="[^"]+"/, `<dimension ref="${ref}"`),
    );
}

function addBaseSheet(
  workbook: ExcelJS.Workbook,
  name = 'Dimensions',
  rows = 4,
  columns = 4,
) {
  const sheet = workbook.addWorksheet(name);
  sheet.addRow(Array.from({ length: columns }, (_, index) => `Header ${index + 1}`));
  for (let row = 2; row <= rows; row += 1) {
    sheet.addRow(
      Array.from({ length: columns }, (_, columnIndex) => {
        if (columnIndex === 0) return `R${row}`;
        if (columnIndex === 1) return row * 10;
        return `R${row}C${columnIndex + 1}`;
      }),
    );
  }
  return sheet;
}

function addTableWithHiddenColumnsWorkbook(workbook: ExcelJS.Workbook) {
  const sheet = workbook.addWorksheet('HiddenTableCols');
  sheet.addTable({
    name: 'HiddenColumnTable',
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    columns: [
      { name: 'SKU' },
      { name: 'Region' },
      { name: 'Hidden Amount' },
      { name: 'Hidden Units' },
      { name: 'Status' },
    ],
    rows: [
      ['A-100', 'North', 100, 1, 'open'],
      ['B-200', 'South', 200, 2, 'open'],
      ['C-300', 'West', 300, 3, 'closed'],
      ['D-400', 'East', 400, 4, 'open'],
    ],
  });
  for (const columnNumber of [3, 4]) {
    const column = sheet.getColumn(columnNumber);
    column.hidden = true;
    column.outlineLevel = 1;
    column.width = 18;
  }
}

function addTableWithHiddenRowsWorkbook(workbook: ExcelJS.Workbook) {
  const sheet = workbook.addWorksheet('HiddenTableRows');
  sheet.addTable({
    name: 'HiddenRowTable',
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    columns: [
      { name: 'SKU' },
      { name: 'Region' },
      { name: 'Amount' },
      { name: 'Units' },
    ],
    rows: [
      ['A-100', 'North', 100, 1],
      ['B-200', 'South', 200, 2],
      ['C-300', 'West', 300, 3],
      ['D-400', 'East', 400, 4],
    ],
  });
  for (const rowNumber of [3, 4]) {
    const row = sheet.getRow(rowNumber);
    row.hidden = true;
    row.outlineLevel = 1;
    row.height = 22;
  }
}

function addOffsetTableWorkbook(workbook: ExcelJS.Workbook) {
  const sheet = addBaseSheet(workbook, 'OffsetTable');
  sheet.addTable({
    name: 'FarDimensionTable',
    ref: 'J20',
    headerRow: true,
    totalsRow: false,
    columns: [
      { name: 'Item' },
      { name: 'Quantity' },
      { name: 'Amount' },
    ],
    rows: [
      ['Alpha', 1, 10],
      ['Beta', 2, 20],
      ['Gamma', 3, 30],
    ],
  });
  sheet.getColumn('J').width = 20;
  sheet.getColumn('K').hidden = true;
  sheet.getRow(22).height = 24;
}

const candidates: Candidate[] = [
  {
    id: 'table-hidden-column-body-edit',
    issue: 'Excel table body cells are edited in hidden outline columns with custom widths',
    configure: addTableWithHiddenColumnsWorkbook,
    edits: [
      { address: 'C3', value: 2300 },
      { address: 'D4', value: 44 },
    ],
  },
  {
    id: 'table-hidden-row-body-edit',
    issue: 'Excel table body cells are edited in hidden outline rows with custom heights',
    configure: addTableWithHiddenRowsWorkbook,
    edits: [
      { address: 'C3', value: 2300 },
      { address: 'D4', value: 44 },
    ],
  },
  {
    id: 'table-outside-stale-dimension-edit',
    issue: 'offset Excel table lies outside a stale worksheet dimension and table cells are edited',
    configure: addOffsetTableWorkbook,
    postprocess: setDimension('A1:D4'),
    edits: [
      { address: 'K22', value: 2200 },
      { address: 'L23', value: 3300 },
    ],
  },
];

function checkWithComSignal(filePath: string): ExcelCheckResult {
  let result = checkWithCom(filePath);
  for (let attempt = 2; attempt <= 3 && result.status === 'error'; attempt += 1) {
    result = checkWithCom(filePath);
  }
  return result;
}

function candidatePaths(id: string) {
  const dir = path.join(runRoot, id);
  return {
    sourcePath: path.join(dir, `${id}.xlsx`),
    exportedPath: path.join(dir, `${id}.mog-export.xlsx`),
  };
}

for (const candidate of candidates) {
  test(`agent18 ${candidate.id}`, async ({ page }, testInfo) => {
    const { sourcePath, exportedPath } = candidatePaths(candidate.id);

    await writeWorkbook(sourcePath, candidate.configure, candidate.postprocess);

    const sourceCheck = checkWithComSignal(sourcePath);
    testInfo.annotations.push({
      type: 'agent18-source-com',
      description: `${candidate.issue}: ${sourceCheck.status}: ${sourceCheck.message}`,
    });
    expect(sourceCheck.status, `source workbook must be valid: ${sourceCheck.message}`).toBe('ok');

    await exportAfterEdits(page, sourcePath, exportedPath, candidate.edits);

    const exportedCheck = checkWithComSignal(exportedPath);
    testInfo.annotations.push({
      type: 'agent18-export-com',
      description: `${candidate.issue}: ${exportedCheck.status}: ${exportedCheck.message}`,
    });
    expect(exportedCheck.status, exportedCheck.message).toBe('corrupt');
  });
}
