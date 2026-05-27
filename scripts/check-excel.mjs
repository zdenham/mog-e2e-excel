import { access, copyFile, mkdir, mkdtemp, rm, stat, writeFile } from 'node:fs/promises';
import { constants } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawn } from 'node:child_process';

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const excelLockDir = path.join(tmpdir(), 'mog-excel-check.lock');
const excelValidationDir = path.join(repoRoot, 'excel-validation');
const lockWaitMs = 180_000;
const staleLockMs = 300_000;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function run(command, args, options = {}) {
  return new Promise((resolve) => {
    const child = spawn(command, args, {
      stdio: ['ignore', 'pipe', 'pipe'],
      ...options,
    });
    let stdout = '';
    let stderr = '';
    child.stdout.on('data', (chunk) => {
      stdout += chunk;
    });
    child.stderr.on('data', (chunk) => {
      stderr += chunk;
    });
    child.on('close', (code, signal) => resolve({ code, signal, stdout, stderr }));
    child.on('error', (error) =>
      resolve({ code: 1, stdout, stderr: String(error.message ?? error) }),
    );
  });
}

async function assertReadableFile(filePath) {
  await access(filePath, constants.R_OK);
}

async function hasExcel() {
  const result = await run('osascript', ['-e', 'id of application "Microsoft Excel"']);
  return result.code === 0;
}

function validationFileName(filePath) {
  const extension = path.extname(filePath) || '.xlsx';
  const baseName = path
    .basename(filePath, extension)
    .replace(/[^A-Za-z0-9._-]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 80);
  return `${Date.now()}-${process.pid}-${baseName || 'workbook'}${extension}`;
}

async function copyToStableExcelFolder(filePath) {
  await mkdir(excelValidationDir, { recursive: true });
  const stablePath = path.join(excelValidationDir, validationFileName(filePath));
  await copyFile(filePath, stablePath);
  return stablePath;
}

async function acquireExcelLock() {
  const deadline = Date.now() + lockWaitMs;

  while (Date.now() < deadline) {
    try {
      await mkdir(excelLockDir);
      await writeFile(
        path.join(excelLockDir, 'owner.txt'),
        `${process.pid} ${new Date().toISOString()}\n`,
      );
      return async () => {
        await rm(excelLockDir, { recursive: true, force: true });
      };
    } catch (error) {
      if (error?.code !== 'EEXIST') {
        throw error;
      }

      try {
        const lockStats = await stat(excelLockDir);
        if (Date.now() - lockStats.mtimeMs > staleLockMs) {
          await rm(excelLockDir, { recursive: true, force: true });
          continue;
        }
      } catch (statError) {
        if (statError?.code === 'ENOENT') {
          continue;
        }
      }

      await sleep(500);
    }
  }

  throw new Error('Timed out waiting for the Excel automation lock.');
}

async function checkWithExcel(filePath) {
  const absoluteFilePath = path.resolve(filePath);
  await assertReadableFile(absoluteFilePath);

  if (process.platform !== 'darwin') {
    return {
      status: 'unsupported',
      message: 'Actual Excel validation is currently implemented for macOS only.',
    };
  }

  if (!(await hasExcel())) {
    return {
      status: 'skipped',
      message: 'Microsoft Excel is not installed or is not discoverable by LaunchServices.',
    };
  }

  let releaseLock;
  let excelFilePath = absoluteFilePath;
  try {
    releaseLock = await acquireExcelLock();
    excelFilePath = await copyToStableExcelFolder(absoluteFilePath);
    await run('pkill', ['-x', 'Microsoft Excel']);
    await sleep(1000);

    const tempDir = await mkdtemp(path.join(tmpdir(), 'mog-excel-check-'));
    const scriptPath = path.join(tempDir, 'check-excel.applescript');
    await writeFile(
      scriptPath,
      `
on collectWindowText()
  set collectedText to ""
  tell application "System Events"
    if exists process "Microsoft Excel" then
      tell process "Microsoft Excel"
        repeat with currentWindow in windows
          try
            set collectedText to collectedText & " " & (name of currentWindow as text)
          end try
          try
            set collectedText to collectedText & " " & (value of static texts of currentWindow as text)
          end try
          try
            set collectedText to collectedText & " " & (description of buttons of currentWindow as text)
          end try
        end repeat
      end tell
    end if
  end tell
  return collectedText
end collectWindowText

on dismissKnownDialogs()
  tell application "System Events"
    if exists process "Microsoft Excel" then
      tell process "Microsoft Excel"
        try
          click button "No" of window 1
        end try
        try
          click button "Cancel" of window 1
        end try
        try
          click button "Don't Save" of window 1
        end try
        try
          click button "OK" of window 1
        end try
        try
          click button 2 of window 1
        end try
      end tell
    end if
  end tell
end dismissKnownDialogs

on run argv
  set workbookPath to item 1 of argv
  set workbookName to do shell script "basename " & quoted form of workbookPath
  set lowerWorkbookName to do shell script "printf %s " & quoted form of workbookName & " | tr '[:upper:]' '[:lower:]'"
  my dismissKnownDialogs()
  try
    do shell script "open -a " & quoted form of "Microsoft Excel" & " " & quoted form of workbookPath
  on error errMsg number errNo
    return "CORRUPT_OPEN_ERROR: " & errMsg
  end try

  repeat with pollIndex from 1 to 45
    delay 1
    set dialogText to my collectWindowText()
    set lowerDialogText to do shell script "printf %s " & quoted form of dialogText & " | tr '[:upper:]' '[:lower:]'"

    if lowerDialogText contains "found a problem" or lowerDialogText contains "corrupt" or lowerDialogText contains "repair" or lowerDialogText contains "recovered" or lowerDialogText contains "unreadable content" then
      if lowerDialogText does not contain lowerWorkbookName then
        my dismissKnownDialogs()
        delay 1
      else
      my dismissKnownDialogs()
      return "CORRUPT_DIALOG: " & dialogText
      end if
    end if
  end repeat

  tell application "Microsoft Excel"
    try
      close active workbook saving no
    end try
  end tell
  return "OK"
end run
`,
    );

    const result = await run('osascript', [scriptPath, excelFilePath], { timeout: 90_000 });
    const output = `${result.stdout}${result.stderr}`.trim();
    if (result.code !== 0) {
      if (output.toLowerCase().includes('not allowed assistive access')) {
        return {
          status: 'skipped',
          message:
            'Excel is installed, but macOS denied Accessibility automation access for osascript/System Events.',
        };
      }
      if (result.signal) {
        return {
          status: 'skipped',
          message: `Excel automation did not complete before timeout (${result.signal}).`,
        };
      }
      return { status: 'error', message: output || 'osascript failed' };
    }
    if (output.startsWith('OK')) {
      return { status: 'ok', message: output };
    }
    if (output.startsWith('CORRUPT_')) {
      return { status: 'corrupt', message: output };
    }
    return { status: 'error', message: output || 'Unexpected empty Excel check output' };
  } catch (error) {
    return { status: 'error', message: String(error?.message ?? error) };
  } finally {
    if (excelFilePath !== absoluteFilePath) {
      await rm(excelFilePath, { force: true }).catch(() => undefined);
    }
    if (releaseLock) {
      await releaseLock();
    }
  }
}

const invokedPath = process.argv[1] ? path.resolve(process.argv[1]) : '';
const modulePath = path.resolve(fileURLToPath(import.meta.url));

if (invokedPath === modulePath) {
  const filePath = process.argv[2];
  if (!filePath) {
    console.error('Usage: npm run excel:check -- /absolute/path/to/workbook.xlsx');
    process.exit(2);
  }
  const result = await checkWithExcel(filePath);
  console.log(JSON.stringify(result, null, 2));
  if (result.status === 'corrupt' || result.status === 'error') {
    process.exit(1);
  }
}

export { checkWithExcel };
