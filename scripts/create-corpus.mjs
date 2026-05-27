import ExcelJS from 'exceljs';
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const corpusDir = path.join(repoRoot, 'corpus');

async function writeWorkbook(name, configure) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'mog-e2e-excel';
  workbook.created = new Date('2026-01-01T00:00:00Z');
  workbook.modified = new Date('2026-01-01T00:00:00Z');
  await configure(workbook);
  const buffer = await workbook.xlsx.writeBuffer();
  await writeFile(path.join(corpusDir, name), Buffer.from(buffer));
}

async function simpleWorkbook(workbook) {
  const sheet = workbook.addWorksheet('Inputs');
  sheet.columns = [
    { header: 'Item', key: 'item', width: 18 },
    { header: 'Units', key: 'units', width: 12 },
    { header: 'Price', key: 'price', width: 12 },
    { header: 'Total', key: 'total', width: 12 },
  ];
  [
    ['Widgets', 12, 4.5],
    ['Adapters', 8, 9.25],
    ['Licenses', 3, 125],
  ].forEach(([item, units, price], index) => {
    const row = index + 2;
    sheet.addRow({ item, units, price, total: { formula: `B${row}*C${row}` } });
  });
  sheet.getCell('D5').value = { formula: 'SUM(D2:D4)' };
  sheet.getCell('D5').numFmt = '$#,##0.00';
}

async function formattedWorkbook(workbook) {
  const sheet = workbook.addWorksheet('Formats');
  sheet.mergeCells('A1:D1');
  sheet.getCell('A1').value = 'Regional revenue snapshot';
  sheet.getCell('A1').font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
  sheet.getCell('A1').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF124559' },
  };
  sheet.addRow(['Region', 'Date', 'Revenue', 'Margin']);
  sheet.getRow(2).font = { bold: true };
  [
    ['North', new Date('2026-01-31T00:00:00Z'), 182000, 0.32],
    ['South', new Date('2026-02-28T00:00:00Z'), 161500, 0.28],
    ['West', new Date('2026-03-31T00:00:00Z'), 203250, 0.36],
  ].forEach((row) => sheet.addRow(row));
  sheet.getColumn(2).numFmt = 'm/d/yyyy';
  sheet.getColumn(3).numFmt = '$#,##0';
  sheet.getColumn(4).numFmt = '0.0%';
  sheet.views = [{ state: 'frozen', ySplit: 2 }];
}

async function multiSheetWorkbook(workbook) {
  const assumptions = workbook.addWorksheet('Assumptions');
  assumptions.addRows([
    ['Metric', 'Value'],
    ['Base revenue', 500000],
    ['Growth', 0.12],
    ['Tax rate', 0.21],
  ]);
  assumptions.getColumn(2).numFmt = '#,##0.00';

  const model = workbook.addWorksheet('Model');
  model.addRows([
    ['Year', 'Revenue', 'Tax'],
    [2026, { formula: 'Assumptions!B2' }, { formula: 'B2*Assumptions!B4' }],
    [2027, { formula: 'B2*(1+Assumptions!B3)' }, { formula: 'B3*Assumptions!B4' }],
    [2028, { formula: 'B3*(1+Assumptions!B3)' }, { formula: 'B4*Assumptions!B4' }],
  ]);
  model.getColumn(2).numFmt = '$#,##0';
  model.getColumn(3).numFmt = '$#,##0';
}

async function tableWorkbook(workbook) {
  const sheet = workbook.addWorksheet('Table');
  sheet.addTable({
    name: 'SalesTable',
    ref: 'A1',
    headerRow: true,
    totalsRow: true,
    style: {
      theme: 'TableStyleMedium2',
      showRowStripes: true,
    },
    columns: [
      { name: 'SKU', filterButton: true },
      { name: 'Category', filterButton: true },
      { name: 'Units', filterButton: true, totalsRowFunction: 'sum' },
      { name: 'Amount', filterButton: true, totalsRowFunction: 'sum' },
    ],
    rows: [
      ['A-100', 'Hardware', 20, 1200],
      ['B-205', 'Software', 5, 3750],
      ['C-310', 'Services', 9, 8100],
    ],
  });
  sheet.getColumn(4).numFmt = '$#,##0';
}

await mkdir(corpusDir, { recursive: true });
await writeWorkbook('simple-formulas.xlsx', simpleWorkbook);
await writeWorkbook('formats-dates-merged.xlsx', formattedWorkbook);
await writeWorkbook('multi-sheet-references.xlsx', multiSheetWorkbook);
await writeWorkbook('table-autofilter.xlsx', tableWorkbook);

console.log(`Wrote XLSX corpus to ${corpusDir}`);
