import { expect, test } from '@playwright/test';
import { existsSync, readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

type ExportFailureScenario = {
  id: string;
  file: string;
  expectedExportStatus: 'error';
  expectedMessage: string;
  issue: string;
};

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const corpusDir = path.join(repoRoot, 'corpus');

function exportFailureScenarios(): ExportFailureScenario[] {
  const manifestPath = path.join(corpusDir, 'export-failure-scenarios.json');
  const manifest = JSON.parse(readFileSync(manifestPath, 'utf8')) as {
    scenarios: ExportFailureScenario[];
  };
  return manifest.scenarios;
}

for (const scenario of exportFailureScenarios()) {
  const fileName = scenario.file;
  const filePath = path.join(corpusDir, fileName);

  test(`export failure ${scenario.id}: ${fileName}`, async ({ page }, testInfo) => {
    expect(existsSync(filePath), `${filePath} should exist`).toBe(true);

    await page.goto('/');

    await page.getByTestId('file-input').setInputFiles(filePath);
    await expect(page.getByTestId('status')).toContainText(/Loaded/);
    await expect(page.getByTestId('loaded-file')).toContainText(fileName);

    const result = await page.evaluate(async () => {
      if (!window.__mogHarness?.exportXlsx) {
        throw new Error('Mog E2E harness did not expose exportXlsx.');
      }

      try {
        const bytes = await window.__mogHarness.exportXlsx();
        return { status: 'ok' as const, byteLength: bytes.length, message: '' };
      } catch (error) {
        return {
          status: 'error' as const,
          byteLength: 0,
          message: error instanceof Error ? error.message : String(error),
        };
      }
    });

    testInfo.annotations.push({
      type: 'mog-export-check',
      description: `${scenario.issue}: ${result.status}: ${result.message}`,
    });

    expect(result.status, result.message).toBe(scenario.expectedExportStatus);
    expect(result.byteLength).toBe(0);
    expect(result.message).toContain(scenario.expectedMessage);
  });
}
