import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const corpusDir = path.join(repoRoot, 'corpus');

async function writeWorkbook(name, configure, postprocess) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'mog-e2e-excel';
  workbook.created = new Date('2026-01-01T00:00:00Z');
  workbook.modified = new Date('2026-01-01T00:00:00Z');
  await configure(workbook);
  let buffer = Buffer.from(await workbook.xlsx.writeBuffer());
  if (postprocess) {
    buffer = await postprocess(buffer);
  }
  await writeFile(path.join(corpusDir, name), buffer);
}

async function addUnusedHyperlinkRelationship(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const relPath = 'xl/worksheets/_rels/sheet1.xml.rels';
  const relsFile = zip.file(relPath);
  let relsXml = relsFile
    ? await relsFile.async('string')
    : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  if (!relsXml.includes('rIdAgent09Unused')) {
    relsXml = relsXml.replace(
      '</Relationships>',
      '<Relationship Id="rIdAgent09Unused" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/stale-unused-relationship" TargetMode="External"/></Relationships>',
    );
    zip.file(relPath, relsXml);
  }
  return zip.generateAsync({ type: 'nodebuffer' });
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

async function tableNoTotalsWorkbook(workbook) {
  const sheet = workbook.addWorksheet('TableNoTotals');
  sheet.addTable({
    name: 'NoTotalsTable',
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    style: { theme: 'TableStyleMedium4', showRowStripes: true },
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
}

async function tableStructuredFormulaWorkbook(workbook) {
  const sheet = workbook.addWorksheet('Structured');
  sheet.addTable({
    name: 'StructuredSales',
    ref: 'A1',
    headerRow: true,
    totalsRow: true,
    style: { theme: 'TableStyleMedium9', showRowStripes: true },
    columns: [
      { name: 'SKU', filterButton: true },
      { name: 'Units', filterButton: true, totalsRowFunction: 'sum' },
      { name: 'Price', filterButton: true },
      { name: 'Amount', filterButton: true, totalsRowFunction: 'sum' },
    ],
    rows: [
      ['A-100', 20, 60, { formula: 'B2*C2', result: 1200 }],
      ['B-205', 5, 750, { formula: 'B3*C3', result: 3750 }],
      ['C-310', 9, 900, { formula: 'B4*C4', result: 8100 }],
    ],
  });
  sheet.getColumn(4).numFmt = '$#,##0';
}

async function tableSpecialHeadersWorkbook(workbook) {
  const sheet = workbook.addWorksheet('SpecialHeaders');
  sheet.addTable({
    name: 'SpecialHeaderTable',
    ref: 'A1',
    headerRow: true,
    totalsRow: true,
    style: { theme: 'TableStyleLight11', showRowStripes: true },
    columns: [
      { name: 'Item #', filterButton: true },
      { name: 'Region/Team', filterButton: true },
      { name: 'Units Sold', filterButton: true, totalsRowFunction: 'sum' },
      { name: 'Net $', filterButton: true, totalsRowFunction: 'sum' },
    ],
    rows: [
      ['A-100', 'North', 20, 1200],
      ['B-205', 'South', 5, 3750],
      ['C-310', 'West', 9, 8100],
    ],
  });
}

async function tableOffsetWorkbook(workbook) {
  const sheet = workbook.addWorksheet('OffsetTable');
  sheet.getCell('A1').value = 'Preamble outside the table';
  sheet.addTable({
    name: 'OffsetSales',
    ref: 'B2',
    headerRow: true,
    totalsRow: true,
    style: { theme: 'TableStyleMedium6', showRowStripes: true },
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
}

async function twoTablesWorkbook(workbook) {
  const sheet = workbook.addWorksheet('TwoTables');
  sheet.addTable({
    name: 'LeftSales',
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    style: { theme: 'TableStyleMedium2', showRowStripes: true },
    columns: [
      { name: 'Left SKU', filterButton: true },
      { name: 'Left Units', filterButton: true },
      { name: 'Left Amount', filterButton: true },
    ],
    rows: [
      ['L-1', 2, 20],
      ['L-2', 4, 40],
    ],
  });
  sheet.addTable({
    name: 'RightSales',
    ref: 'F1',
    headerRow: true,
    totalsRow: false,
    style: { theme: 'TableStyleMedium3', showRowStripes: true },
    columns: [
      { name: 'Right SKU', filterButton: true },
      { name: 'Right Units', filterButton: true },
      { name: 'Right Amount', filterButton: true },
    ],
    rows: [
      ['R-1', 3, 30],
      ['R-2', 6, 60],
    ],
  });
}

async function autofilterOnlyWorkbook(workbook) {
  const sheet = workbook.addWorksheet('AutoFilterOnly');
  sheet.addRows([
    ['SKU', 'Category', 'Units', 'Amount'],
    ['A-100', 'Hardware', 20, 1200],
    ['B-205', 'Software', 5, 3750],
    ['C-310', 'Services', 9, 8100],
  ]);
  sheet.autoFilter = 'A1:D4';
}

async function dataValidationWorkbook(workbook) {
  const sheet = workbook.addWorksheet('Validation');
  sheet.addRows([
    ['Item', 'Status', 'Score'],
    ['A-100', 'Open', 10],
    ['B-205', 'Closed', 20],
  ]);
  sheet.getCell('B2').dataValidation = {
    type: 'list',
    allowBlank: false,
    formulae: ['"Open,Closed,Blocked"'],
  };
  sheet.getCell('C2').dataValidation = {
    type: 'whole',
    operator: 'between',
    formulae: [1, 100],
  };
}

async function hyperlinkWorkbook(workbook) {
  const sheet = workbook.addWorksheet('Links');
  sheet.addRows([
    ['Label', 'URL'],
    ['Mog', 'https://github.com/fundamental-research-labs/mog'],
  ]);
  sheet.getCell('A2').value = {
    text: 'Mog repository',
    hyperlink: 'https://github.com/fundamental-research-labs/mog',
    tooltip: 'Open Mog on GitHub',
  };
  sheet.getCell('B2').value = {
    text: 'Internal target',
    hyperlink: '#Links!A1',
  };
}

async function hyperlinkExternalWorkbook(workbook) {
  const sheet = workbook.addWorksheet('ExternalLinks');
  sheet.addRows([
    ['Label', 'URL', 'Notes'],
    ['OpenAI', 'https://openai.com', 'External link in A2'],
    ['GitHub', 'https://github.com', 'External link in A3'],
  ]);
  sheet.getCell('A2').value = {
    text: 'OpenAI',
    hyperlink: 'https://openai.com',
    tooltip: 'External HTTPS link',
  };
  sheet.getCell('A3').value = {
    text: 'GitHub',
    hyperlink: 'https://github.com',
    tooltip: 'Second external HTTPS link',
  };
}

async function hyperlinkInternalWorkbook(workbook) {
  const links = workbook.addWorksheet('InternalLinks');
  links.addRows([
    ['Label', 'Destination', 'Notes'],
    ['Jump to target', '#Targets!B2', 'Internal workbook link in A2'],
    ['Jump to local cell', '#InternalLinks!A1', 'Internal same-sheet link in A3'],
  ]);
  links.getCell('A2').value = {
    text: 'Jump to target',
    hyperlink: '#Targets!B2',
    tooltip: 'Internal link to Targets!B2',
  };
  links.getCell('A3').value = {
    text: 'Jump to local cell',
    hyperlink: '#InternalLinks!A1',
    tooltip: 'Internal same-sheet link',
  };

  const targets = workbook.addWorksheet('Targets');
  targets.addRows([
    ['Key', 'Value'],
    ['Target', 'Linked destination'],
  ]);
}

async function hyperlinkMixedWorkbook(workbook) {
  const sheet = workbook.addWorksheet('MixedLinks');
  sheet.addRows([
    ['Label', 'Target', 'Notes'],
    ['External', 'https://example.com/report?x=1&y=2', 'query string'],
    ['Mailto', 'mailto:test@example.com?subject=Mog%20E2E', 'mailto relation'],
    ['File URL', 'file:///tmp/mog-e2e-link-target.txt', 'file URL relation'],
    ['Internal', '#MixedLinks!A1', 'internal location'],
  ]);
  sheet.getCell('A2').value = {
    text: 'External query link',
    hyperlink: 'https://example.com/report?x=1&y=2',
  };
  sheet.getCell('A3').value = {
    text: 'Mail link',
    hyperlink: 'mailto:test@example.com?subject=Mog%20E2E',
  };
  sheet.getCell('A4').value = {
    text: 'File link',
    hyperlink: 'file:///tmp/mog-e2e-link-target.txt',
  };
  sheet.getCell('A5').value = {
    text: 'Internal link',
    hyperlink: '#MixedLinks!A1',
  };
}

async function definedNamesWorkbook(workbook) {
  const sheet = workbook.addWorksheet('Names');
  sheet.addRows([
    ['Metric', 'Value', 'Formula'],
    ['NamedInput', 100, { formula: 'NamedInput*2' }],
  ]);
  workbook.definedNames.add('Names!$B$2', 'NamedInput');
}

async function mergedCellsWorkbook(workbook) {
  const sheet = workbook.addWorksheet('Merged');
  sheet.mergeCells('A1:D1');
  sheet.getCell('A1').value = 'Merged title';
  sheet.addRows([
    [],
    ['Item', 'Units', 'Price', 'Amount'],
    ['A-100', 2, 10, { formula: 'B4*C4' }],
  ]);
}

async function hiddenOutlineWorkbook(workbook) {
  const sheet = workbook.addWorksheet('HiddenOutline');
  sheet.addRows([
    ['Item', 'Units'],
    ['Visible', 1],
    ['Hidden row', 2],
    ['Grouped row', 3],
    ['Visible tail', 4],
  ]);
  sheet.getRow(3).hidden = true;
  sheet.getRow(4).outlineLevel = 1;
  sheet.getColumn(2).hidden = true;
}

async function formulaWorkbook(workbook) {
  const sheet = workbook.addWorksheet('Formulas');
  sheet.addRows([
    ['A', 'B', 'C', 'D'],
    [1, 2, { formula: 'A2+B2', result: 3 }, { formula: 'SUM(A2:C2)', result: 6 }],
  ]);
}

const scenarios = [
  {
    id: 'table-autofilter-header-row',
    file: 'table-autofilter.xlsx',
    edit: 'table-header-row-values',
    expectedExcelStatus: 'corrupt',
    issue: 'table header cells changed but xl/tables/table1.xml keeps old tableColumn names',
  },
  {
    id: 'table-autofilter-header-a1-only',
    file: 'table-autofilter.xlsx',
    edit: 'table-header-a1-only',
    expectedExcelStatus: 'corrupt',
    issue: 'single table header cell changed without tableColumn metadata update',
  },
  {
    id: 'table-autofilter-header-formula',
    file: 'table-autofilter.xlsx',
    edit: 'table-header-formula-cell',
    expectedExcelStatus: 'corrupt',
    issue: 'formula written into table header while tableColumn metadata remains string header',
  },
  {
    id: 'table-autofilter-header-duplicate',
    file: 'table-autofilter.xlsx',
    edit: 'table-header-duplicate-cell',
    expectedExcelStatus: 'corrupt',
    issue: 'duplicate visible table header while tableColumn metadata remains unique',
  },
  {
    id: 'table-autofilter-header-blank',
    file: 'table-autofilter.xlsx',
    edit: 'table-header-blank-cell',
    expectedExcelStatus: 'corrupt',
    issue: 'blank visible table header while tableColumn metadata remains non-empty',
  },
  {
    id: 'table-autofilter-header-number',
    file: 'table-autofilter.xlsx',
    edit: 'table-header-number-cell',
    expectedExcelStatus: 'corrupt',
    issue: 'numeric visible table header while tableColumn metadata remains text',
  },
  {
    id: 'table-autofilter-header-special-chars',
    file: 'table-autofilter.xlsx',
    edit: 'table-header-special-chars',
    expectedExcelStatus: 'corrupt',
    issue: 'structured-reference-sensitive visible headers while tableColumn metadata keeps old names',
  },
  {
    id: 'table-offset-header-row',
    file: 'table-offset.xlsx',
    edit: 'table-header-offset-row',
    expectedExcelStatus: 'corrupt',
    issue: 'non-A1 table header cells changed while tableColumn metadata keeps old names',
  }
];

await mkdir(corpusDir, { recursive: true });
await writeWorkbook('simple-formulas.xlsx', simpleWorkbook);
await writeWorkbook('formats-dates-merged.xlsx', formattedWorkbook);
await writeWorkbook('multi-sheet-references.xlsx', multiSheetWorkbook);
await writeWorkbook('table-autofilter.xlsx', tableWorkbook);
await writeWorkbook('table-no-totals.xlsx', tableNoTotalsWorkbook);
await writeWorkbook('table-structured-formulas.xlsx', tableStructuredFormulaWorkbook);
await writeWorkbook('table-special-headers.xlsx', tableSpecialHeadersWorkbook);
await writeWorkbook('table-offset.xlsx', tableOffsetWorkbook);
await writeWorkbook('two-tables.xlsx', twoTablesWorkbook);
await writeWorkbook('autofilter-only.xlsx', autofilterOnlyWorkbook);
await writeWorkbook('data-validation.xlsx', dataValidationWorkbook);
await writeWorkbook('hyperlinks.xlsx', hyperlinkWorkbook);
await writeWorkbook('defined-names.xlsx', definedNamesWorkbook);
await writeWorkbook('merged-cells.xlsx', mergedCellsWorkbook);
await writeWorkbook('hidden-outline.xlsx', hiddenOutlineWorkbook);
await writeWorkbook('formulas.xlsx', formulaWorkbook);
await writeFile(
  path.join(corpusDir, 'scenarios.json'),
  JSON.stringify({ generatedAt: '2026-01-01T00:00:00.000Z', scenarios }, null, 2) + '\n',
);

console.log(`Wrote XLSX corpus to ${corpusDir}`);
