import { expect, test } from '@playwright/test';
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

type CorpusScenario = {
  id: string;
  file: string;
  edit: string;
  expectedExcelStatus: 'ok' | 'corrupt';
  issue: string;
};

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const corpusDir = path.join(repoRoot, 'corpus');
const exportDir = process.env.PIPELINE_EXPORT_DIR
  ? path.resolve(process.env.PIPELINE_EXPORT_DIR)
  : path.join(repoRoot, 'test-results', 'pipeline', 'exports');
const scenarioGrep = process.env.SCENARIO_GREP ? new RegExp(process.env.SCENARIO_GREP) : null;

function corpusScenarios(): CorpusScenario[] {
  const manifestPath = path.join(corpusDir, 'scenarios.json');
  if (!existsSync(manifestPath)) {
    execFileSync('npm', ['run', 'corpus:create'], {
      cwd: repoRoot,
      stdio: 'inherit',
    });
  }

  const manifest = JSON.parse(readFileSync(manifestPath, 'utf8')) as {
    scenarios: CorpusScenario[];
  };
  return scenarioGrep
    ? manifest.scenarios.filter((scenario) => scenarioGrep.test(scenario.id))
    : manifest.scenarios;
}

test.beforeAll(() => {
  mkdirSync(exportDir, { recursive: true });
});

for (const scenario of corpusScenarios()) {
  const fileName = scenario.file;
  const filePath = path.join(corpusDir, fileName);

  test(`export artifact ${scenario.id}: ${fileName}`, async ({ page }) => {
    await page.goto('/');

    await page.getByTestId('file-input').setInputFiles(filePath);
    await expect(page.getByTestId('status')).toContainText(/Loaded/);
    await expect(page.getByTestId('loaded-file')).toContainText(fileName);

    await page.evaluate(async (editId) => {
      if (!window.__mogHarness?.applyScenarioEdit) {
        throw new Error('Mog E2E harness did not expose applyScenarioEdit.');
      }
      await window.__mogHarness.applyScenarioEdit(editId);
    }, scenario.edit);
    await expect(page.getByTestId('status')).toContainText(/dirty|clean/);

    const downloadPromise = page.waitForEvent('download');
    await page.getByTestId('export-xlsx').click();
    const download = await downloadPromise;
    const exportedPath = path.join(exportDir, `${scenario.id}.mog-export.xlsx`);
    await download.saveAs(exportedPath);

    const byteLength = statSync(exportedPath).size;
    expect(byteLength).toBeGreaterThan(0);

    writeFileSync(
      path.join(exportDir, `${scenario.id}.json`),
      JSON.stringify(
        {
          ...scenario,
          exportedPath,
          byteLength,
        },
        null,
        2,
      ) + '\n',
    );
  });
}
