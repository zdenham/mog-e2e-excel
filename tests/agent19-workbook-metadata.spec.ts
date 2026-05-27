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
  edits: CellEdit[];
};

test.setTimeout(240_000);

const runRoot = path.join(
  tmpdir(),
  `mog-agent19-workbook-metadata-${process.pid}-${Date.now()}`,
);

const packageRelationshipNamespace =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const officeRelationshipNamespace =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const spreadsheetNamespace =
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

async function rewriteZip(
  buffer: Uint8Array,
  rewriter: (zip: JSZip) => Promise<void> | void,
) {
  const zip = await JSZip.loadAsync(buffer);
  await rewriter(zip);
  return zip.generateAsync({ type: 'uint8array' });
}

async function zipText(zip: JSZip, partPath: string) {
  const part = zip.file(partPath);
  if (!part) {
    throw new Error(`Missing ${partPath}.`);
  }
  return part.async('string');
}

function insertBeforeClosingTag(xml: string, closingTag: string, insertion: string) {
  if (!xml.includes(closingTag)) {
    throw new Error(`Expected ${closingTag} in XML part.`);
  }
  return xml.replace(closingTag, `${insertion}${closingTag}`);
}

function replaceRequired(xml: string, pattern: RegExp, replacement: string) {
  if (!pattern.test(xml)) {
    throw new Error(`Pattern not found: ${pattern}`);
  }
  return xml.replace(pattern, replacement);
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

function baseWorkbook(workbook: ExcelJS.Workbook) {
  workbook.creator = 'agent19';
  workbook.lastModifiedBy = 'agent19';
  workbook.subject = 'Workbook metadata corruption probe';

  const sheet = workbook.addWorksheet('Metadata');
  sheet.addRows([
    ['Key', 'Value', 'Formula', 'Notes'],
    ['alpha', 10, { formula: 'B2*2', result: 20 }, 'source'],
    ['beta', 20, { formula: 'SUM(B2:B3)', result: 30 }, 'source'],
    ['gamma', 30, { formula: 'B4+B3', result: 50 }, 'source'],
  ]);
  sheet.getColumn(1).width = 18;
  sheet.getColumn(2).width = 14;
  sheet.getColumn(3).width = 14;
  sheet.getColumn(4).width = 28;
}

async function addMetadataBackedDynamicArray(buffer: Uint8Array) {
  return rewriteZip(buffer, async (zip) => {
    const relsPath = 'xl/_rels/workbook.xml.rels';
    const contentTypesPath = '[Content_Types].xml';
    const sheetPath = 'xl/worksheets/sheet1.xml';
    const relsXml = await zipText(zip, relsPath);
    const contentTypesXml = await zipText(zip, contentTypesPath);
    const sheetXml = await zipText(zip, sheetPath);

    zip.file(
      'xl/metadata.xml',
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        `<metadata xmlns="${spreadsheetNamespace}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">`,
        '<metadataTypes count="1">',
        '<metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>',
        '</metadataTypes>',
        '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>',
        '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>',
        '</metadata>',
      ].join(''),
    );
    zip.file(
      relsPath,
      addRelationship(
        relsXml,
        'rIdAgent19SheetMetadata',
        `${officeRelationshipNamespace}/sheetMetadata`,
        'metadata.xml',
      ),
    );
    zip.file(
      contentTypesPath,
      addContentTypeOverride(
        contentTypesXml,
        '/xl/metadata.xml',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml',
      ),
    );
    zip.file(
      sheetPath,
      replaceRequired(
        sheetXml,
        /<c r="C2">[\s\S]*?<\/c>/,
        '<c r="C2" cm="1"><f>SEQUENCE(2,2,B2,1)</f><v>10</v></c>',
      ),
    );
  });
}

const candidates: Candidate[] = [
  {
    id: 'metadata-backed-dynamic-array-input-edit',
    issue: 'metadata-backed dynamic array formula remains while only its input cell is edited',
    edits: [
      { address: 'B2', value: 919 },
      { address: 'D4', value: 'non-spill edit' },
    ],
  },
  {
    id: 'metadata-backed-dynamic-array-anchor-edit',
    issue: 'metadata-backed dynamic array formula anchor is replaced by a literal edit',
    edits: [
      { address: 'C2', value: 'replace metadata-backed dynamic formula' },
      { address: 'D4', value: 'non-spill edit' },
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
  test(`agent19 ${candidate.id}`, async ({ page }, testInfo) => {
    const { sourcePath, exportedPath } = candidatePaths(candidate.id);

    await writeWorkbook(sourcePath, baseWorkbook, addMetadataBackedDynamicArray);

    const sourceCheck = checkWithComSignal(sourcePath);
    testInfo.annotations.push({
      type: 'agent19-source-com',
      description: `${candidate.issue}: ${sourceCheck.status}: ${sourceCheck.message}`,
    });
    expect(sourceCheck.status, `source workbook must be valid: ${sourceCheck.message}`).toBe('ok');

    await exportAfterEdits(page, sourcePath, exportedPath, candidate.edits);

    const exportedCheck = checkWithComSignal(exportedPath);
    testInfo.annotations.push({
      type: 'agent19-export-com',
      description: `${candidate.issue}: ${exportedCheck.status}: ${exportedCheck.message}`,
    });
    expect(exportedCheck.status, exportedCheck.message).toBe('corrupt');
  });
}
