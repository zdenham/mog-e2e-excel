import { spawn } from 'node:child_process';
import { mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pipelineDir = path.join(repoRoot, 'test-results', 'pipeline');
const exportDir = path.join(pipelineDir, 'exports');
const resultPath = path.join(pipelineDir, 'results.json');

function usage() {
  return `Usage: node scripts/run-e2e-pipeline.mjs [--parallel N] [--com] [--grep PATTERN]

Runs the Mog import/edit/export phase in parallel, then validates exported XLSX
artifacts with actual Excel.

Options:
  --parallel N   Number of Playwright export workers. Also used as COM validator
                 concurrency on Windows. macOS Excel validation stays serial.
  --com          Use Windows Excel COM/UI Automation validator.
  --grep PATTERN Run only scenarios whose id matches PATTERN.
  --keep         Keep the previous test-results/pipeline directory.
`;
}

function parseArgs(argv) {
  const options = {
    parallel: 1,
    useCom: false,
    grep: '',
    keep: false,
  };

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === '--parallel') {
      options.parallel = Number(argv[++index]);
    } else if (arg === '--com') {
      options.useCom = true;
    } else if (arg === '--grep') {
      options.grep = argv[++index] ?? '';
    } else if (arg === '--keep') {
      options.keep = true;
    } else if (arg === '--help' || arg === '-h') {
      console.log(usage());
      process.exit(0);
    } else {
      throw new Error(`Unknown argument: ${arg}\n\n${usage()}`);
    }
  }

  if (!Number.isInteger(options.parallel) || options.parallel < 1) {
    throw new Error('--parallel must be a positive integer.');
  }

  return options;
}

function commandName(command) {
  return process.platform === 'win32' ? `${command}.cmd` : command;
}

function run(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd: repoRoot,
      stdio: options.capture ? ['ignore', 'pipe', 'pipe'] : 'inherit',
      env: { ...process.env, ...options.env },
      shell: process.platform === 'win32',
    });

    let stdout = '';
    let stderr = '';
    if (options.capture) {
      child.stdout.on('data', (chunk) => {
        stdout += chunk;
      });
      child.stderr.on('data', (chunk) => {
        stderr += chunk;
      });
    }

    child.on('error', reject);
    child.on('close', (code, signal) => {
      const result = { code, signal, stdout, stderr };
      if (options.allowFailure || code === 0) {
        resolve(result);
      } else {
        reject(
          new Error(
            `${command} ${args.join(' ')} failed with ${signal ?? code}\n${stdout}${stderr}`,
          ),
        );
      }
    });
  });
}

function readJson(filePath) {
  return readFile(filePath, 'utf8').then((text) => JSON.parse(text));
}

async function mapLimit(items, limit, worker) {
  const results = new Array(items.length);
  let nextIndex = 0;

  async function runNext() {
    while (nextIndex < items.length) {
      const currentIndex = nextIndex;
      nextIndex += 1;
      results[currentIndex] = await worker(items[currentIndex], currentIndex);
    }
  }

  await Promise.all(
    Array.from({ length: Math.min(limit, items.length) }, () => runNext()),
  );
  return results;
}

async function validateWithMacExcel(filePath) {
  const result = await run('node', ['scripts/check-excel.mjs', filePath], {
    capture: true,
    allowFailure: true,
  });
  const output = `${result.stdout}${result.stderr}`.trim();
  try {
    return JSON.parse(output);
  } catch {
    return {
      status: 'error',
      message: output || `macOS Excel checker exited with ${result.signal ?? result.code}`,
    };
  }
}

async function validateWithCom(filePath) {
  if (process.platform !== 'win32') {
    return {
      status: 'unsupported',
      message: '--com requires Windows because it uses Excel COM automation.',
    };
  }

  const result = await run(
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
    { capture: true, allowFailure: true },
  );
  const output = `${result.stdout}${result.stderr}`.trim();
  try {
    return JSON.parse(output);
  } catch {
    return {
      status: 'error',
      message: output || `COM Excel checker exited with ${result.signal ?? result.code}`,
    };
  }
}

function scenarioMatches(scenario, grep) {
  return !grep || new RegExp(grep).test(scenario.id);
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  if (options.useCom && process.platform !== 'win32') {
    throw new Error('--com can only run on Windows hosts with Microsoft Excel installed.');
  }

  if (!options.keep) {
    await rm(pipelineDir, { recursive: true, force: true });
  }
  await mkdir(exportDir, { recursive: true });

  await run(commandName('npm'), ['run', 'corpus:create']);

  await run(
    commandName('npx'),
    [
      'playwright',
      'test',
      'tests/mog-export-artifacts.spec.ts',
      'tests/mog-export-failure.spec.ts',
      '--workers',
      String(options.parallel),
    ],
    {
      env: {
        PIPELINE_EXPORT_DIR: exportDir,
        SCENARIO_GREP: options.grep,
      },
    },
  );

  const manifest = await readJson(path.join(corpusDir(), 'scenarios.json'));
  const scenarios = manifest.scenarios.filter((scenario) => scenarioMatches(scenario, options.grep));
  const validatorConcurrency = options.useCom ? options.parallel : 1;
  const validatorName = options.useCom ? 'windows-com' : 'macos-applescript';
  const requireExcel = process.env.REQUIRE_EXCEL === '1';

  if (!options.useCom && options.parallel > 1) {
    console.log('Excel validation is serialized on macOS; only the export phase used --parallel.');
  }

  const validations = await mapLimit(scenarios, validatorConcurrency, async (scenario) => {
    const exportedPath = path.join(exportDir, `${scenario.id}.mog-export.xlsx`);
    if (!existsSync(exportedPath)) {
      return {
        ...scenario,
        exportedPath,
        validator: validatorName,
        actualStatus: 'error',
        message: 'Expected export artifact was not created.',
        passed: false,
      };
    }

    const result = options.useCom
      ? await validateWithCom(exportedPath)
      : await validateWithMacExcel(exportedPath);
    const skipped = result.status === 'skipped' || result.status === 'unsupported';
    const passed = skipped && !requireExcel ? true : result.status === scenario.expectedExcelStatus;

    return {
      ...scenario,
      exportedPath,
      validator: validatorName,
      actualStatus: result.status,
      message: result.message,
      passed,
    };
  });

  const failed = validations.filter((result) => !result.passed);
  const summary = {
    generatedAt: new Date().toISOString(),
    parallel: options.parallel,
    validator: validatorName,
    exportDir,
    failed: failed.length,
    passed: validations.length - failed.length,
    results: validations,
  };
  await writeFile(resultPath, JSON.stringify(summary, null, 2) + '\n');

  for (const result of validations) {
    const marker = result.passed ? 'PASS' : 'FAIL';
    console.log(
      `${marker} ${result.id}: expected=${result.expectedExcelStatus} actual=${result.actualStatus}`,
    );
  }

  console.log(`Wrote ${resultPath}`);
  if (failed.length > 0) {
    process.exitCode = 1;
  }
}

function corpusDir() {
  return path.join(repoRoot, 'corpus');
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
