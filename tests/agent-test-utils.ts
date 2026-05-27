import { expect, type Page, type TestInfo } from '@playwright/test';
import ExcelJS from 'exceljs';
import { spawnSync } from 'node:child_process';
import { mkdirSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

export type CellEdit = {
  address: string;
  value: string | number | boolean | null;
};

export type ExcelCheckResult = {
  status: 'ok' | 'corrupt' | 'error' | 'skipped' | 'unsupported';
  message: string;
};

export const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

export async function writeWorkbook(
  filePath: string,
  configure: (workbook: ExcelJS.Workbook) => Promise<void> | void,
  postprocess?: (buffer: Uint8Array) => Promise<Uint8Array> | Uint8Array,
) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'mog-e2e-excel-agent';
  workbook.created = new Date('2026-01-01T00:00:00Z');
  workbook.modified = new Date('2026-01-01T00:00:00Z');
  await configure(workbook);
  let buffer: Uint8Array = Buffer.from(await workbook.xlsx.writeBuffer());
  if (postprocess) {
    buffer = await postprocess(buffer);
  }
  mkdirSync(path.dirname(filePath), { recursive: true });
  writeFileSync(filePath, buffer);
}

export async function exportAfterEdits(
  page: Page,
  sourcePath: string,
  exportedPath: string,
  edits: CellEdit[],
) {
  await page.goto('/');
  await page.getByTestId('file-input').setInputFiles(sourcePath);
  await expect(page.getByTestId('status')).toContainText(/Loaded/);

  await page.evaluate(async (cellEdits) => {
    if (!window.__mogHarness?.setCells) {
      throw new Error('Mog E2E harness did not expose setCells.');
    }
    await window.__mogHarness.setCells(cellEdits);
  }, edits);
  await expect(page.getByTestId('status')).toContainText(/dirty|clean/);

  const bytes = await page.evaluate(async () => {
    if (!window.__mogHarness?.exportXlsx) {
      throw new Error('Mog E2E harness did not expose exportXlsx.');
    }
    return Array.from(await window.__mogHarness.exportXlsx());
  });
  mkdirSync(path.dirname(exportedPath), { recursive: true });
  writeFileSync(exportedPath, Buffer.from(bytes));
}

export function checkWithCom(filePath: string): ExcelCheckResult {
  const result = spawnSync(
    'powershell.exe',
    [
      '-NoProfile',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      path.join(repoRoot, 'scripts', 'check-excel-com.ps1'),
      '-Path',
      filePath,
    ],
    {
      cwd: repoRoot,
      encoding: 'utf8',
    },
  );
  const output = `${result.stdout || ''}${result.stderr || ''}`.trim();
  try {
    return JSON.parse(output) as ExcelCheckResult;
  } catch {
    return {
      status: 'error',
      message: output || `COM Excel checker exited with ${result.status}`,
    };
  }
}

export function agentPaths(testInfo: TestInfo, id: string) {
  const dir = testInfo.outputPath(id);
  return {
    sourcePath: path.join(dir, `${id}.xlsx`),
    exportedPath: path.join(dir, `${id}.mog-export.xlsx`),
  };
}
