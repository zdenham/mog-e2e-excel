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
  distinctness: string;
  configure: (workbook: ExcelJS.Workbook) => Promise<void> | void;
  edits: CellEdit[];
  postprocess?: (buffer: Uint8Array) => Promise<Uint8Array> | Uint8Array;
};

type DefinedNamePatch = {
  name: string;
  formula: string;
  localSheetId?: number;
  hidden?: boolean;
};

test.setTimeout(360_000);

const runRoot = path.join(tmpdir(), `mog-agent20-tables-names-print-${process.pid}-${Date.now()}`);

function escapeXml(value: string | number) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function insertBeforeClosingTag(xml: string, closingTag: string, insertion: string) {
  if (!xml.includes(closingTag)) {
    throw new Error(`Expected ${closingTag} in XML part.`);
  }
  return xml.replace(closingTag, `${insertion}${closingTag}`);
}

async function rewriteZip(
  buffer: Uint8Array,
  rewriter: (zip: JSZip) => Promise<void> | void,
) {
  const zip = await JSZip.loadAsync(buffer);
  await rewriter(zip);
  return zip.generateAsync({ type: 'uint8array' });
}

function definedNameXml(name: DefinedNamePatch) {
  const localSheetId =
    typeof name.localSheetId === 'number' ? ` localSheetId="${name.localSheetId}"` : '';
  const hidden = name.hidden ? ' hidden="1"' : '';
  return `<definedName name="${escapeXml(name.name)}"${localSheetId}${hidden}>${escapeXml(
    name.formula,
  )}</definedName>`;
}

function addDefinedNames(names: DefinedNamePatch[]) {
  return (buffer: Uint8Array) =>
    rewriteZip(buffer, async (zip) => {
      const workbookPath = 'xl/workbook.xml';
      const workbook = zip.file(workbookPath);
      if (!workbook) {
        throw new Error(`Missing ${workbookPath}.`);
      }
      const insertion = names.map(definedNameXml).join('');
      const xml = await workbook.async('string');
      zip.file(
        workbookPath,
        xml.includes('<definedNames>')
          ? xml.replace('</definedNames>', `${insertion}</definedNames>`)
          : insertBeforeClosingTag(xml, '</workbook>', `<definedNames>${insertion}</definedNames>`),
      );
    });
}

function addPageSetup(sheet: ExcelJS.Worksheet, printArea = 'A1:F8') {
  sheet.pageSetup.paperSize = 9;
  sheet.pageSetup.orientation = 'landscape';
  sheet.pageSetup.fitToPage = true;
  sheet.pageSetup.fitToWidth = 1;
  sheet.pageSetup.fitToHeight = 0;
  sheet.pageSetup.printArea = printArea;
  sheet.pageSetup.printTitlesRow = '1:1';
  sheet.pageSetup.printTitlesColumn = 'A:A';
  sheet.pageSetup.horizontalCentered = true;
  sheet.pageSetup.margins = {
    left: 0.25,
    right: 0.25,
    top: 0.5,
    bottom: 0.5,
    header: 0.2,
    footer: 0.2,
  };
}

function addTotalsTableWorkbook(workbook: ExcelJS.Workbook, sheetName = 'SalesPrint') {
  const sheet = workbook.addWorksheet(sheetName);
  addPageSetup(sheet);
  sheet.addTable({
    name: 'SalesPrintTable',
    displayName: 'SalesPrintTable',
    ref: 'A1',
    headerRow: true,
    totalsRow: true,
    style: { theme: 'TableStyleMedium2', showRowStripes: true },
    columns: [
      { name: 'SKU', filterButton: true, totalsRowLabel: 'Printed Total' },
      { name: 'Region', filterButton: true },
      { name: 'Units', filterButton: true, totalsRowFunction: 'sum' },
      {
        name: 'Amount',
        filterButton: true,
        totalsRowFunction: 'custom',
        totalsRowFormula: 'SUBTOTAL(109,SalesPrintTable[Amount])',
      },
    ],
    rows: [
      ['A-100', 'North', 2, 120],
      ['B-200', 'South', 4, 280],
      ['C-300', 'West', 6, 420],
    ],
  });
  sheet.getColumn(4).numFmt = '$#,##0';
  sheet.getCell('F1').value = 'Structured total';
  sheet.getCell('F2').value = { formula: 'SUM(SalesPrintTable[Amount])', result: 820 };
  sheet.getCell('F3').value = { formula: 'SUM(SalesPrintTable[Units])', result: 12 };
  workbook.definedNames.add(`${sheetName}!$D$2:$D$4`, 'Agent20AmountRange');
}

function addOffsetTableWorkbook(workbook: ExcelJS.Workbook) {
  const sheet = workbook.addWorksheet('OffsetPrint');
  addPageSetup(sheet, 'A1:H10');
  sheet.getCell('A1').value = 'Printable report band';
  sheet.getCell('B2').value = 'Title outside table';
  sheet.addTable({
    name: 'OffsetPrintTable',
    displayName: 'OffsetPrintTable',
    ref: 'B3',
    headerRow: true,
    totalsRow: true,
    style: { theme: 'TableStyleMedium6', showRowStripes: true },
    columns: [
      { name: 'SKU', filterButton: true, totalsRowLabel: 'Printed Total' },
      { name: 'Region', filterButton: true },
      { name: 'Units', filterButton: true, totalsRowFunction: 'sum' },
      {
        name: 'Amount',
        filterButton: true,
        totalsRowFunction: 'custom',
        totalsRowFormula: 'SUBTOTAL(109,OffsetPrintTable[Amount])',
      },
    ],
    rows: [
      ['A-100', 'North', 2, 120],
      ['B-200', 'South', 4, 280],
      ['C-300', 'West', 6, 420],
    ],
  });
  sheet.getCell('H4').value = { formula: 'SUM(OffsetPrintTable[Amount])', result: 820 };
}

function addTwoPrintTablesWorkbook(workbook: ExcelJS.Workbook) {
  const sheet = workbook.addWorksheet('TwoPrintTables');
  addPageSetup(sheet, 'A1:I8');
  sheet.addTable({
    name: 'LeftPrintTable',
    displayName: 'LeftPrintTable',
    ref: 'A1',
    headerRow: true,
    totalsRow: true,
    style: { theme: 'TableStyleMedium3', showRowStripes: true },
    columns: [
      { name: 'SKU', filterButton: true, totalsRowLabel: 'Total' },
      { name: 'Units', filterButton: true, totalsRowFunction: 'sum' },
      {
        name: 'Amount',
        filterButton: true,
        totalsRowFunction: 'custom',
        totalsRowFormula: 'SUBTOTAL(109,LeftPrintTable[Amount])',
      },
    ],
    rows: [
      ['A-100', 2, 120],
      ['B-200', 4, 280],
    ],
  });
  sheet.addTable({
    name: 'RightPrintTable',
    displayName: 'RightPrintTable',
    ref: 'F1',
    headerRow: true,
    totalsRow: true,
    style: { theme: 'TableStyleMedium5', showRowStripes: true },
    columns: [
      { name: 'SKU', filterButton: true, totalsRowLabel: 'Total' },
      { name: 'Units', filterButton: true, totalsRowFunction: 'sum' },
      {
        name: 'Amount',
        filterButton: true,
        totalsRowFunction: 'custom',
        totalsRowFormula: 'SUBTOTAL(109,RightPrintTable[Amount])',
      },
    ],
    rows: [
      ['C-300', 6, 420],
      ['D-400', 8, 640],
    ],
  });
  sheet.getCell('D2').value = { formula: 'SUM(LeftPrintTable[Amount])', result: 400 };
  sheet.getCell('I2').value = { formula: 'SUM(RightPrintTable[Amount])', result: 1060 };
}

const candidates: Candidate[] = [
  {
    id: 'totals-row-label-only-edit',
    issue: 'table totals label cell is edited while the printed table keeps totals-row metadata',
    distinctness: 'isolates the totals label cell without changing table headers or formulas',
    configure: addTotalsTableWorkbook,
    edits: [
      { address: 'A5', value: 'Edited printed total' },
    ],
  },
  {
    id: 'structured-name-with-totals-label-edit',
    issue: 'structured-reference defined name is present while the table totals label is edited under print setup',
    distinctness: 'explicit table-backed defined name plus totals-label corruption, not header-name mismatch',
    configure: addTotalsTableWorkbook,
    postprocess: addDefinedNames([
      { name: 'Agent20NamedPrintedAmounts', formula: 'SalesPrintTable[Amount]' },
    ]),
    edits: [
      { address: 'A5', value: 'Edited printed total' },
      { address: 'F2', value: '=SUM(Agent20NamedPrintedAmounts)' },
    ],
  },
  {
    id: 'offset-table-total-scalar-overwrite-with-print-titles',
    issue: 'offset table totals formula and label are overwritten while print titles span the table',
    distinctness: 'offset totals-row overwrite, not the known offset header corruption',
    configure: addOffsetTableWorkbook,
    edits: [
      { address: 'B7', value: 'Edited printed total' },
      { address: 'E7', value: 999 },
    ],
  },
  {
    id: 'two-tables-second-total-scalar-overwrite-with-print-area',
    issue: 'second table totals formula and label are overwritten inside a print area containing two tables',
    distinctness: 'second-table totals-row overwrite, not the known second-table header corruption',
    configure: addTwoPrintTablesWorkbook,
    edits: [
      { address: 'F4', value: 'Edited total' },
      { address: 'H4', value: 999 },
    ],
  },
];

const activeIds = process.env.AGENT20_ACTIVE
  ? new Set(process.env.AGENT20_ACTIVE.split(',').map((id) => id.trim()).filter(Boolean))
  : null;

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

for (const candidate of candidates.filter((item) => !activeIds || activeIds.has(item.id))) {
  test(`agent20 ${candidate.id}`, async ({ page }, testInfo) => {
    const { sourcePath, exportedPath } = candidatePaths(candidate.id);

    await writeWorkbook(sourcePath, candidate.configure, candidate.postprocess);

    const sourceCheck = checkWithComSignal(sourcePath);
    testInfo.annotations.push({
      type: 'agent20-source-com',
      description: `${candidate.issue}: ${sourceCheck.status}: ${sourceCheck.message}`,
    });
    expect(sourceCheck.status, `source workbook must be valid: ${sourceCheck.message}`).toBe('ok');

    await exportAfterEdits(page, sourcePath, exportedPath, candidate.edits);

    const exportedCheck = checkWithComSignal(exportedPath);
    testInfo.annotations.push({
      type: 'agent20-export-com',
      description: `${candidate.issue}: ${exportedCheck.status}: ${exportedCheck.message}`,
    });
    testInfo.annotations.push({
      type: 'distinctness',
      description: candidate.distinctness,
    });

    expect(exportedCheck.status, exportedCheck.message).toBe('corrupt');
  });
}
