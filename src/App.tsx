import {
  FileDown,
  FileSpreadsheet,
  Loader2,
  Pencil,
  Upload,
} from 'lucide-react';
import {
  MogSpreadsheetApp,
  createSpreadsheetRuntime,
  type SpreadsheetAppAttachmentHandle,
  type SpreadsheetAppError,
  type SpreadsheetDirtyState,
  type SpreadsheetRuntime,
  type SpreadsheetSaveRequest,
  type SpreadsheetSaveResult,
  type SpreadsheetWorkbookSession,
} from '@mog-sdk/spreadsheet-app';
import type { Workbook } from '@mog-sdk/contracts/api';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';

type LoadedFile = {
  id: string;
  name: string;
  bytes: Uint8Array;
  versionId: string;
};

type WorkbookState =
  | { status: 'idle' }
  | { status: 'loading'; fileName: string }
  | {
      status: 'ready';
      runtime: SpreadsheetRuntime;
      workbook: SpreadsheetWorkbookSession;
      file: LoadedFile;
    }
  | { status: 'error'; message: string };

type HarnessApi = {
  loadFile: (file: File) => Promise<void>;
  exportXlsx: () => Promise<Uint8Array>;
  applyEdit: () => Promise<void>;
  getStatus: () => string;
};

declare global {
  interface Window {
    __mogHarness?: HarnessApi;
  }
}

function hashBytes(bytes: Uint8Array): string {
  let hash = 2166136261;
  for (const byte of bytes) {
    hash ^= byte;
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(16).padStart(8, '0');
}

function toUint8Array(input: ArrayBuffer | Uint8Array): Uint8Array {
  return input instanceof Uint8Array ? input : new Uint8Array(input);
}

function downloadBytes(bytes: Uint8Array, fileName: string) {
  const downloadBuffer = bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
  const blob = new Blob([downloadBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

async function createRuntime(): Promise<SpreadsheetRuntime> {
  const runtime = await createSpreadsheetRuntime({
    host: {
      persistenceMode: 'host-owned-ephemeral',
      beforeUnloadPrompt: false,
    },
    onSaveRequest: async (request: SpreadsheetSaveRequest): Promise<SpreadsheetSaveResult> => {
      return {
        status: 'saved',
        workbookId: request.workbookId,
        epoch: request.epoch,
        baseVersionId: request.baseVersionId,
        dirtyEpoch: request.dirtyEpoch,
        changeSequence: request.changeSequence,
        saveRequestId: request.saveRequestId,
        bytesHash: request.bytesHash,
        versionId: `host-save-${Date.now()}`,
      };
    },
  });
  await runtime.ready;
  return runtime;
}

function exportedFileName(inputName: string): string {
  return inputName.replace(/\.xlsx$/i, '') + '.mog-export.xlsx';
}

export function App() {
  const [state, setState] = useState<WorkbookState>({ status: 'idle' });
  const [dirtyState, setDirtyState] = useState<SpreadsheetDirtyState | null>(null);
  const [lastExportName, setLastExportName] = useState<string>('');
  const [lastError, setLastError] = useState<string>('');
  const attachmentRef = useRef<SpreadsheetAppAttachmentHandle | null>(null);
  const stateRef = useRef<WorkbookState>(state);

  useEffect(() => {
    stateRef.current = state;
  }, [state]);

  const disposeCurrentWorkbook = useCallback(async () => {
    const current = stateRef.current;
    if (current.status !== 'ready') return;
    attachmentRef.current = null;
    await current.workbook.dispose().catch(() => undefined);
    await current.runtime.dispose().catch(() => undefined);
  }, []);

  const loadFile = useCallback(
    async (file: File) => {
      setLastError('');
      setLastExportName('');
      setDirtyState(null);
      setState({ status: 'loading', fileName: file.name });

      await disposeCurrentWorkbook();

      try {
        const bytes = new Uint8Array(await file.arrayBuffer());
        const loadedFile: LoadedFile = {
          id: `file-${hashBytes(bytes)}-${Date.now()}`,
          name: file.name,
          bytes,
          versionId: `import-${hashBytes(bytes)}`,
        };

        const runtime = await createRuntime();
        const workbook = await runtime.openWorkbook({
          workbookId: loadedFile.id,
          displayName: loadedFile.name,
          source: {
            kind: 'xlsx-bytes',
            bytes: loadedFile.bytes,
            fileName: loadedFile.name,
            versionId: loadedFile.versionId,
          },
        });
        await workbook.ready;
        setDirtyState({
          status: 'clean',
          workbookId: workbook.workbookId,
          epoch: workbook.epoch,
          changeSequence: 0,
          versionId: loadedFile.versionId,
        });
        setState({ status: 'ready', runtime, workbook, file: loadedFile });
      } catch (error) {
        setState({
          status: 'error',
          message: error instanceof Error ? error.message : String(error),
        });
      }
    },
    [disposeCurrentWorkbook],
  );

  const currentReadyState = state.status === 'ready' ? state : null;

  useEffect(() => {
    if (!currentReadyState) return undefined;
    return currentReadyState.workbook.onDirtyChange(setDirtyState);
  }, [currentReadyState]);

  useEffect(() => {
    return () => {
      void disposeCurrentWorkbook();
    };
  }, [disposeCurrentWorkbook]);

  const exportXlsx = useCallback(async (): Promise<Uint8Array> => {
    const current = stateRef.current;
    if (current.status !== 'ready') {
      throw new Error('No workbook is loaded.');
    }
    const bytes = toUint8Array(await current.workbook.exportXlsx());
    const name = exportedFileName(current.file.name);
    downloadBytes(bytes, name);
    setLastExportName(name);
    return bytes;
  }, []);

  const applyEdit = useCallback(async () => {
    const current = stateRef.current;
    if (current.status !== 'ready') {
      throw new Error('No workbook is loaded.');
    }
    const workbookApi = current.workbook.getWorkbook();
    await workbookApi.batch('E2E harness edit', async (workbook: Workbook) => {
      const sheet = workbook.activeSheet;
      await sheet.setCell('A1', 'Mog E2E export smoke test');
      await sheet.setCell('B1', new Date().toISOString());
      await sheet.setCell('C1', '=LEN(A1)');
    });
  }, []);

  useEffect(() => {
    window.__mogHarness = {
      loadFile,
      exportXlsx,
      applyEdit,
      getStatus: () => stateRef.current.status,
    };
    return () => {
      delete window.__mogHarness;
    };
  }, [applyEdit, exportXlsx, loadFile]);

  const statusLabel = useMemo(() => {
    if (state.status === 'idle') return 'No workbook loaded';
    if (state.status === 'loading') return `Loading ${state.fileName}`;
    if (state.status === 'error') return 'Import failed';
    return dirtyState?.status === 'dirty' ? 'Loaded, dirty' : 'Loaded, clean';
  }, [dirtyState?.status, state]);

  const handleExport = async () => {
    setLastError('');
    try {
      await exportXlsx();
    } catch (error) {
      setLastError(error instanceof Error ? error.message : String(error));
    }
  };

  const handleEdit = async () => {
    setLastError('');
    try {
      await applyEdit();
    } catch (error) {
      setLastError(error instanceof Error ? error.message : String(error));
    }
  };

  const handleMogError = (error: SpreadsheetAppError) => {
    setLastError(error.message);
  };

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="brand">
          <FileSpreadsheet aria-hidden="true" size={22} />
          <span>Mog XLSX E2E</span>
        </div>
        <div className="toolbar" aria-label="Workbook actions">
          <label className="icon-button file-picker" title="Upload XLSX">
            <Upload aria-hidden="true" size={18} />
            <input
              data-testid="file-input"
              type="file"
              accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              onChange={(event) => {
                const file = event.currentTarget.files?.[0];
                if (file) void loadFile(file);
                event.currentTarget.value = '';
              }}
            />
          </label>
          <button
            data-testid="apply-edit"
            type="button"
            className="icon-button"
            disabled={state.status !== 'ready'}
            onClick={handleEdit}
            title="Apply test edit"
          >
            <Pencil aria-hidden="true" size={18} />
          </button>
          <button
            data-testid="export-xlsx"
            type="button"
            className="primary-button"
            disabled={state.status !== 'ready'}
            onClick={handleExport}
          >
            <FileDown aria-hidden="true" size={18} />
            Export
          </button>
        </div>
      </header>

      <section className="status-strip" aria-live="polite">
        <span data-testid="status">{statusLabel}</span>
        {state.status === 'ready' ? (
          <span className="muted" data-testid="loaded-file">
            {state.file.name}
          </span>
        ) : null}
        {lastExportName ? (
          <span className="muted" data-testid="last-export">
            {lastExportName}
          </span>
        ) : null}
        {lastError || state.status === 'error' ? (
          <span className="error" data-testid="error">
            {lastError || (state.status === 'error' ? state.message : '')}
          </span>
        ) : null}
      </section>

      <section className="spreadsheet-region" data-testid="spreadsheet-region">
        {state.status === 'idle' ? (
          <div className="empty-state">
            <Upload aria-hidden="true" size={28} />
            <span>Upload an XLSX workbook to start.</span>
          </div>
        ) : null}
        {state.status === 'loading' ? (
          <div className="empty-state">
            <Loader2 className="spin" aria-hidden="true" size={28} />
            <span>Opening workbook in Mog.</span>
          </div>
        ) : null}
        {state.status === 'ready' ? (
          <MogSpreadsheetApp
            runtime={state.runtime}
            workbook={state.workbook}
            workspace={{
              mode: 'single-document',
              fileExplorer: false,
              appSwitcher: false,
              settings: false,
            }}
            chrome={{
              fileMenu: false,
              commandBar: {
                mode: 'mog',
                tabs: ['home', 'insert', 'formulas', 'data', 'view'],
                hiddenGroups: ['charts'],
              },
              formulaBar: true,
              sheetTabs: true,
              statusBar: true,
            }}
            commands={{
              save: 'host',
              open: 'host',
              import: 'disabled',
              export: 'host',
              print: 'disabled',
            }}
            editModel={{
              user: 'write',
              agents: 'write',
              automation: 'write',
            }}
            onReady={(attachment) => {
              attachmentRef.current = attachment;
              void attachment.ready;
            }}
            onError={handleMogError}
            loadingFallback={
              <div className="empty-state">
                <Loader2 className="spin" aria-hidden="true" size={28} />
                <span>Attaching Mog UI.</span>
              </div>
            }
          />
        ) : null}
      </section>
    </main>
  );
}
