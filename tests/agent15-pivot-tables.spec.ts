import { expect, test } from '@playwright/test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { tmpdir } from 'node:os';
import path from 'node:path';
import {
  checkWithCom,
  exportAfterEdits,
  type CellEdit,
  writeWorkbook,
} from './agent-test-utils';

type PivotOptions = {
  cacheId: number;
  sourceSheetName: string;
  sourceRef?: string;
  sourceTableName?: string;
  pivotSheetIndex: number;
  pivotLocationRef: string;
  pivotTableName: string;
};

type Candidate = {
  id: string;
  issue: string;
  configure: (workbook: ExcelJS.Workbook) => Promise<void> | void;
  pivot: PivotOptions;
  edits: CellEdit[];
};

const packageRelationshipNamespace =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const officeRelationshipNamespace =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const spreadsheetNamespace =
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const pivotCacheDefinitionContentType =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml';
const pivotCacheRecordsContentType =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml';
const pivotTableContentType =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml';
const workRoot = path.join(tmpdir(), 'mog-agent15-pivot-tables');

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

function addRelationship(xml: string, id: string, type: string, target: string) {
  if (xml.includes(`Id="${id}"`)) {
    return xml;
  }
  return insertBeforeClosingTag(
    xml,
    '</Relationships>',
    `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`,
  );
}

function emptyRelationships() {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<Relationships xmlns="${packageRelationshipNamespace}"></Relationships>`,
  ].join('');
}

function addContentTypeOverride(xml: string, partName: string, contentType: string) {
  if (xml.includes(`PartName="${partName}"`)) {
    return xml;
  }
  return insertBeforeClosingTag(
    xml,
    '</Types>',
    `<Override PartName="${partName}" ContentType="${contentType}"/>`,
  );
}

async function wait(milliseconds: number) {
  await new Promise((resolve) => {
    setTimeout(resolve, milliseconds);
  });
}

async function checkWithComRetry(filePath: string) {
  let result = checkWithCom(filePath);
  for (let attempt = 1; result.status === 'error' && attempt < 6; attempt += 1) {
    await wait(2_000 * attempt);
    result = checkWithCom(filePath);
  }
  return result;
}

function ensureWorksheetRelationshipNamespace(xml: string) {
  if (xml.includes('xmlns:r=')) {
    return xml;
  }
  return xml.replace('<worksheet ', `<worksheet xmlns:r="${officeRelationshipNamespace}" `);
}

function addWorkbookPivotCache(xml: string, cacheId: number, relationshipId: string) {
  const pivotCaches = `<pivotCaches><pivotCache cacheId="${cacheId}" r:id="${relationshipId}"/></pivotCaches>`;
  if (xml.includes('<pivotCaches>')) {
    return xml.replace('</pivotCaches>', `<pivotCache cacheId="${cacheId}" r:id="${relationshipId}"/></pivotCaches>`);
  }
  if (/<calcPr\b[^>]*\/>/.test(xml)) {
    return xml.replace(/(<calcPr\b[^>]*\/>)/, `$1${pivotCaches}`);
  }
  return insertBeforeClosingTag(xml, '</workbook>', pivotCaches);
}

function renderPivotCacheDefinition(options: PivotOptions) {
  const worksheetSource = options.sourceTableName
    ? `<worksheetSource name="${escapeXml(options.sourceTableName)}"/>`
    : `<worksheetSource ref="${escapeXml(options.sourceRef ?? 'A1:D6')}" sheet="${escapeXml(
        options.sourceSheetName,
      )}"/>`;

  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<pivotCacheDefinition xmlns="${spreadsheetNamespace}" xmlns:r="${officeRelationshipNamespace}" r:id="rId1" refreshedBy="Agent15" refreshedDate="46169" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="5">`,
    `<cacheSource type="worksheet">${worksheetSource}</cacheSource>`,
    '<cacheFields count="4">',
    '<cacheField name="Region" numFmtId="0"><sharedItems count="3"><s v="East"/><s v="West"/><s v="North"/></sharedItems></cacheField>',
    '<cacheField name="Product" numFmtId="0"><sharedItems/></cacheField>',
    '<cacheField name="Amount" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="80" maxValue="140"/></cacheField>',
    '<cacheField name="Units" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="1" maxValue="4"/></cacheField>',
    '</cacheFields>',
    '</pivotCacheDefinition>',
  ].join('');
}

function renderPivotCacheRecords() {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<pivotCacheRecords xmlns="${spreadsheetNamespace}" count="5">`,
    '<r><x v="0"/><s v="A"/><n v="100"/><n v="2"/></r>',
    '<r><x v="0"/><s v="B"/><n v="120"/><n v="3"/></r>',
    '<r><x v="1"/><s v="A"/><n v="80"/><n v="1"/></r>',
    '<r><x v="1"/><s v="B"/><n v="140"/><n v="4"/></r>',
    '<r><x v="2"/><s v="A"/><n v="90"/><n v="2"/></r>',
    '</pivotCacheRecords>',
  ].join('');
}

function renderPivotTableDefinition(options: PivotOptions) {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<pivotTableDefinition xmlns="${spreadsheetNamespace}" name="${escapeXml(
      options.pivotTableName,
    )}" cacheId="${options.cacheId}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">`,
    `<location ref="${escapeXml(options.pivotLocationRef)}" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>`,
    '<pivotFields count="4">',
    '<pivotField axis="axisRow" showAll="0"><items count="4"><item x="0"/><item x="2"/><item x="1"/><item t="default"/></items></pivotField>',
    '<pivotField showAll="0"/>',
    '<pivotField dataField="1" showAll="0"/>',
    '<pivotField showAll="0"/>',
    '</pivotFields>',
    '<rowFields count="1"><field x="0"/></rowFields>',
    '<rowItems count="4"><i><x/></i><i><x v="1"/></i><i><x v="2"/></i><i t="grand"><x/></i></rowItems>',
    '<colItems count="1"><i/></colItems>',
    '<dataFields count="1"><dataField name="Sum of Amount" fld="2" baseField="0" baseItem="0"/></dataFields>',
    '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>',
    '</pivotTableDefinition>',
  ].join('');
}

function addSalesRows(sheet: ExcelJS.Worksheet) {
  sheet.addRows([
    ['Region', 'Product', 'Amount', 'Units'],
    ['East', 'A', 100, 2],
    ['East', 'B', 120, 3],
    ['West', 'A', 80, 1],
    ['West', 'B', 140, 4],
    ['North', 'A', 90, 2],
  ]);
  sheet.getColumn(1).width = 14;
  sheet.getColumn(2).width = 12;
  sheet.getColumn(3).width = 12;
  sheet.getColumn(4).width = 12;
}

function addPivotOutput(sheet: ExcelJS.Worksheet, startColumn = 1) {
  const rows = [
    ['Region', 'Sum of Amount'],
    ['East', 220],
    ['North', 90],
    ['West', 220],
    ['Grand Total', 530],
  ];
  rows.forEach((row, index) => {
    row.forEach((value, columnOffset) => {
      sheet.getCell(3 + index, startColumn + columnOffset).value = value;
    });
  });
}

async function addPivotPackage(buffer: Uint8Array, options: PivotOptions) {
  const zip = await JSZip.loadAsync(buffer);
  const workbookXmlPath = 'xl/workbook.xml';
  const workbookRelsPath = 'xl/_rels/workbook.xml.rels';
  const worksheetPath = `xl/worksheets/sheet${options.pivotSheetIndex}.xml`;
  const worksheetRelsPath = `xl/worksheets/_rels/sheet${options.pivotSheetIndex}.xml.rels`;
  const pivotCacheDefinitionPath = 'xl/pivotCache/pivotCacheDefinition1.xml';
  const pivotCacheRecordsPath = 'xl/pivotCache/pivotCacheRecords1.xml';
  const pivotCacheDefinitionRelsPath =
    'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels';
  const pivotTablePath = 'xl/pivotTables/pivotTable1.xml';
  const pivotTableRelsPath = 'xl/pivotTables/_rels/pivotTable1.xml.rels';
  const workbookPivotCacheRelId = 'rIdAgent15PivotCache1';
  const worksheetPivotRelId = 'rIdAgent15PivotTable1';

  const workbookXml = await zip.file(workbookXmlPath)?.async('string');
  if (!workbookXml) {
    throw new Error(`Missing ${workbookXmlPath}.`);
  }
  zip.file(
    workbookXmlPath,
    addWorkbookPivotCache(workbookXml, options.cacheId, workbookPivotCacheRelId),
  );

  const workbookRels = await zip.file(workbookRelsPath)?.async('string');
  if (!workbookRels) {
    throw new Error(`Missing ${workbookRelsPath}.`);
  }
  zip.file(
    workbookRelsPath,
    addRelationship(
      workbookRels,
      workbookPivotCacheRelId,
      `${officeRelationshipNamespace}/pivotCacheDefinition`,
      'pivotCache/pivotCacheDefinition1.xml',
    ),
  );

  const worksheetXml = await zip.file(worksheetPath)?.async('string');
  if (!worksheetXml) {
    throw new Error(`Missing ${worksheetPath}.`);
  }
  zip.file(
    worksheetPath,
    insertBeforeClosingTag(
      ensureWorksheetRelationshipNamespace(worksheetXml),
      '</worksheet>',
      `<pivotTableDefinition r:id="${worksheetPivotRelId}"/>`,
    ),
  );

  const worksheetRels = (await zip.file(worksheetRelsPath)?.async('string')) ?? emptyRelationships();
  zip.file(
    worksheetRelsPath,
    addRelationship(
      worksheetRels,
      worksheetPivotRelId,
      `${officeRelationshipNamespace}/pivotTable`,
      '../pivotTables/pivotTable1.xml',
    ),
  );

  zip.file(pivotCacheDefinitionPath, renderPivotCacheDefinition(options));
  zip.file(pivotCacheRecordsPath, renderPivotCacheRecords());
  zip.file(
    pivotCacheDefinitionRelsPath,
    addRelationship(
      emptyRelationships(),
      'rId1',
      `${officeRelationshipNamespace}/pivotCacheRecords`,
      'pivotCacheRecords1.xml',
    ),
  );
  zip.file(pivotTablePath, renderPivotTableDefinition(options));
  zip.file(
    pivotTableRelsPath,
    addRelationship(
      emptyRelationships(),
      'rId1',
      `${officeRelationshipNamespace}/pivotCacheDefinition`,
      '../pivotCache/pivotCacheDefinition1.xml',
    ),
  );

  const contentTypesPath = '[Content_Types].xml';
  const contentTypes = await zip.file(contentTypesPath)?.async('string');
  if (!contentTypes) {
    throw new Error(`Missing ${contentTypesPath}.`);
  }
  zip.file(
    contentTypesPath,
    addContentTypeOverride(
      addContentTypeOverride(
        addContentTypeOverride(
          contentTypes,
          '/xl/pivotCache/pivotCacheDefinition1.xml',
          pivotCacheDefinitionContentType,
        ),
        '/xl/pivotCache/pivotCacheRecords1.xml',
        pivotCacheRecordsContentType,
      ),
      '/xl/pivotTables/pivotTable1.xml',
      pivotTableContentType,
    ),
  );

  return zip.generateAsync({ type: 'uint8array' });
}

function addTableBackedPivotWorkbook(workbook: ExcelJS.Workbook) {
  const source = workbook.addWorksheet('Source');
  addSalesRows(source);
  source.addTable({
    name: 'SalesTable',
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    columns: [
      { name: 'Region' },
      { name: 'Product' },
      { name: 'Amount' },
      { name: 'Units' },
    ],
    rows: [
      ['East', 'A', 100, 2],
      ['East', 'B', 120, 3],
      ['West', 'A', 80, 1],
      ['West', 'B', 140, 4],
      ['North', 'A', 90, 2],
    ],
  });

  const pivot = workbook.addWorksheet('Pivot');
  addPivotOutput(pivot);
}

const candidates: Candidate[] = [
  {
    id: 'table-backed-pivot-body-edit',
    issue: 'pivot cache uses a table source name while table body cells are edited',
    configure: addTableBackedPivotWorkbook,
    pivot: {
      cacheId: 19,
      sourceSheetName: 'Source',
      sourceTableName: 'SalesTable',
      pivotSheetIndex: 2,
      pivotLocationRef: 'A3:B7',
      pivotTableName: 'TableBackedPivot',
    },
    edits: [
      { address: 'C5', value: 515 },
      { address: 'B6', value: 'C' },
    ],
  },
];

for (const candidate of candidates) {
  test(`agent15 ${candidate.id}`, async ({ page }, testInfo) => {
    test.setTimeout(600_000);
    const sourcePath = path.join(workRoot, candidate.id, `${candidate.id}.xlsx`);
    const exportedPath = path.join(workRoot, candidate.id, `${candidate.id}.mog-export.xlsx`);

    await writeWorkbook(sourcePath, candidate.configure, (buffer) =>
      addPivotPackage(buffer, candidate.pivot),
    );

    const sourceCheck = await checkWithComRetry(sourcePath);
    testInfo.annotations.push({
      type: 'agent15-source-com',
      description: `${candidate.issue}: ${sourceCheck.status}: ${sourceCheck.message}`,
    });
    expect(sourceCheck.status, `source workbook must be valid: ${sourceCheck.message}`).toBe('ok');

    await exportAfterEdits(page, sourcePath, exportedPath, candidate.edits);

    const exportedCheck = await checkWithComRetry(exportedPath);
    testInfo.annotations.push({
      type: 'agent15-export-com',
      description: `${candidate.issue}: ${exportedCheck.status}: ${exportedCheck.message}`,
    });
    expect(exportedCheck.status, exportedCheck.message).toBe('corrupt');
  });
}
