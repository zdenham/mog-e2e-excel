import { expect, test } from '@playwright/test';
import { execFileSync, spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, readdirSync, statSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

type ExcelCheckResult = {
  status: 'ok' | 'corrupt' | 'error' | 'skipped' | 'unsupported';
  message: string;
};

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const corpusDir = path.join(repoRoot, 'corpus');
const downloadsDir = path.join(repoRoot, 'test-results', 'exports');

async function checkWithExcel(filePath: string): Promise<ExcelCheckResult> {
  const result = spawnSync('node', ['scripts/check-excel.mjs', filePath], {
    cwd: repoRoot,
    encoding: 'utf8',
  });
  const output = `${result.stdout || ''}${result.stderr || ''}`.trim();
  try {
    return JSON.parse(output) as ExcelCheckResult;
  } catch {
    return {
      status: 'error',
      message: output || `Excel checker exited with ${result.status}`,
    };
  }
}

function corpusFiles(): string[] {
  if (!existsSync(corpusDir)) {
    execFileSync('npm', ['run', 'corpus:create'], {
      cwd: repoRoot,
      stdio: 'inherit',
    });
  }
  return readdirSync(corpusDir)
    .filter((file) => file.endsWith('.xlsx'))
    .sort()
    .map((file) => path.join(corpusDir, file));
}

test.beforeAll(() => {
  mkdirSync(downloadsDir, { recursive: true });
  execFileSync('npm', ['run', 'corpus:create'], {
    cwd: repoRoot,
    stdio: 'inherit',
  });
});

for (const filePath of corpusFiles()) {
  const fileName = path.basename(filePath);

  test(`imports, edits, exports, and opens in Excel: ${fileName}`, async ({
    page,
  }, testInfo) => {
    await page.goto('/');

    await page.getByTestId('file-input').setInputFiles(filePath);
    await expect(page.getByTestId('status')).toContainText(/Loaded/);
    await expect(page.getByTestId('loaded-file')).toContainText(fileName);

    await page.getByTestId('apply-edit').click();
    await expect(page.getByTestId('status')).toContainText(/dirty|clean/);

    const downloadPromise = page.waitForEvent('download');
    await page.getByTestId('export-xlsx').click();
    const download = await downloadPromise;
    const exportedPath = path.join(downloadsDir, `${path.parse(fileName).name}.mog-export.xlsx`);
    await download.saveAs(exportedPath);

    expect(statSync(exportedPath).size).toBeGreaterThan(0);

    const result = await checkWithExcel(exportedPath);
    testInfo.annotations.push({
      type: 'excel-check',
      description: `${result.status}: ${result.message}`,
    });

    if (result.status === 'skipped' || result.status === 'unsupported') {
      if (process.env.REQUIRE_EXCEL === '1') {
        throw new Error(result.message);
      }
      return;
    }

    expect(result.status, result.message).toBe('ok');
  });
}
