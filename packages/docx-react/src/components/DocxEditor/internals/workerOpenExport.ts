import type {
  DocxExportResult,
  DocxLayoutMap,
  DocxPageExportOptions,
  DocxPagedStructuredContent,
  ResidentEngineWorkerClient,
  YrsSession,
} from '@betteroffice/docx/yrs';
import { ResidentWorkerFailureError } from '@betteroffice/docx/yrs';
import { awaitWorkerOpenReplica } from './workerOpenReplica';
import { clearWorkerExportVersions, type WorkerExportVersions } from './workerExportVersions';

export const LAYOUT_WAIT_MS = 2_000;
export const VIEWER_LAYOUT_WAIT_MS = 60_000;
export const LAYOUT_POLL_MS = 16;
export const LAYOUT_REFUSALS: ReadonlySet<string> = new Set([
  'stale-document', 'stale-layout', 'layout-unavailable',
]);

export type WorkerPageExportResult = DocxExportResult<DocxPagedStructuredContent<DocxLayoutMap>>;

export interface WorkerOpenExportContext {
  current(): boolean;
  flush(): Promise<void>;
  request(): Promise<string | null>;
  settleLayout(timeoutMs: number, requestLayout: boolean): Promise<void>;
}

export interface WorkerOpenExport {
  export(options: DocxPageExportOptions, context: WorkerOpenExportContext): Promise<WorkerPageExportResult>;
}

const exporters = new WeakMap<YrsSession, WorkerOpenExport>();

export function registerWorkerOpenExport(session: YrsSession, operation: WorkerOpenExport): () => void {
  exporters.set(session, operation);
  return () => {
    if (exporters.get(session) === operation) exporters.delete(session);
  };
}

export function workerOpenExport(session: YrsSession): WorkerOpenExport | null {
  return exporters.get(session) ?? null;
}

export function retireWorkerOpenExport(session: YrsSession): void {
  exporters.delete(session);
  clearWorkerExportVersions(session);
}

export async function exportWorkerOpenPages(
  peer: YrsSession,
  options: DocxPageExportOptions,
  context: WorkerOpenExportContext,
  worker: {
    assertCurrent(): void;
    catchUp(): Promise<{ P: string; W: string; changed: boolean }>;
    read: ResidentEngineWorkerClient['documentReadAt'];
    versions: WorkerExportVersions;
    serialize<T>(operation: () => Promise<T>): Promise<T>;
  }
): Promise<WorkerPageExportResult> {
  const assertCurrent = (): void => {
    if (!context.current()) throw new Error('The document changed while exporting');
    worker.assertCurrent();
  };
  const refusal = (version: string, code: 'stale-document' | 'layout-unavailable', message: string): WorkerPageExportResult => ({
    ok: false, version, failure: { code, target: null, message },
  });
  assertCurrent();
  const ready = awaitWorkerOpenReplica(peer);
  if (ready) {
    let timer: ReturnType<typeof setTimeout> | undefined;
    let hydrated: boolean;
    try {
      hydrated = await Promise.race([
        ready.then(() => true),
        new Promise<false>((resolve) => { timer = setTimeout(() => resolve(false), VIEWER_LAYOUT_WAIT_MS); }),
      ]);
    } finally {
      if (timer !== undefined) clearTimeout(timer);
    }
    assertCurrent();
    if (!hydrated) return refusal(peer.version(), 'layout-unavailable', 'The document is not laid out yet.');
  }
  assertCurrent();
  await context.flush();
  assertCurrent();
  const result = await worker.serialize(exportCapturedPages);
  assertCurrent();
  return result;

  async function exportCapturedPages(): Promise<WorkerPageExportResult> {
    assertCurrent();
    let capture = await worker.catchUp();
    assertCurrent();
    const moved = (): WorkerPageExportResult => refusal(capture.P, 'stale-document', 'The document changed while exporting.');
    if (capture.changed) return moved();
    const readOptions = options.expectLayoutVersion === undefined ? options : {
      ...options, expectLayoutVersion: worker.versions.workerLayoutVersion(options.expectLayoutVersion),
    };
    let request: string | null = null;
    const attempt = async (): Promise<WorkerPageExportResult> => {
      request = await context.request();
      assertCurrent();
      if (request === null) return refusal(capture.P, 'layout-unavailable', 'The fonts this document uses are not loaded yet.');
      const read = await worker.read({ kind: 'exportStructuredWithPages', options: readOptions, currentRequest: request }, capture.W);
      assertCurrent();
      if (read.status === 'superseded') return moved();
      if (read.version !== capture.W) throw new ResidentWorkerFailureError('Resident engine worker returned an unexpected read version');
      return worker.versions.adapt(JSON.parse(read.value) as WorkerPageExportResult, capture.P, capture.W);
    };
    let result = await attempt();
    assertCurrent();
    if (options.expectLayoutVersion !== undefined) return result;
    const retryable = (): boolean => {
      if (result.ok) return false;
      if (LAYOUT_REFUSALS.has(result.failure.code)) return true;
      if (result.failure.code !== 'unsupported-revision-layout' || request === null) return false;
      const preview = (JSON.parse(request) as { renderEnv?: { revisionPreview?: object } }).renderEnv?.revisionPreview;
      return !preview || Object.keys(preview).length === 0;
    };
    let deadline = 0;
    for (let attempts = 0; retryable() && attempts < Math.ceil(LAYOUT_WAIT_MS / LAYOUT_POLL_MS) + 1; attempts += 1) {
      if (peer.version() !== capture.P) return moved();
      if (attempts > 0 && Date.now() >= deadline) break;
      try {
        await context.settleLayout(attempts === 0 ? VIEWER_LAYOUT_WAIT_MS : Math.max(0, deadline - Date.now()), attempts === 0);
      } catch {
        assertCurrent();
      }
      assertCurrent();
      if (peer.version() !== capture.P) return moved();
      if (attempts === 0 || (!result.ok && result.failure.code === 'stale-document')) {
        const nextCapture = await worker.catchUp();
        assertCurrent();
        if (nextCapture.changed || nextCapture.P !== capture.P) return moved();
        capture = nextCapture;
      }
      result = await attempt();
      assertCurrent();
      if (attempts === 0) deadline = Date.now() + LAYOUT_WAIT_MS;
      if (retryable()) {
        await new Promise((resolve) => setTimeout(resolve, LAYOUT_POLL_MS));
        assertCurrent();
      }
    }
    return result;
  }
}
