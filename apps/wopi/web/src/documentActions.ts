import type { SaveOptions } from "./postMessage";
import type { SaveBarrier, SocketSaveBarrier } from "../../src/relay/sequence";

interface InputEditor {
  flushPendingInput(): Promise<void>;
  hasPendingUpdates(): boolean;
}

interface SaveDocumentOptions extends SaveOptions {
  signal?: AbortSignal;
  captureBarrier?(): SocketSaveBarrier | null;
  hasPendingResync?(): boolean;
  onFlushed?(): void;
  onProgress?(patch: Record<string, unknown>): void;
  inputTimeoutMs?: number;
  saveTimeoutMs?: number;
}

interface SaveApi {
  save(barrier: SaveBarrier, signal?: AbortSignal): Promise<unknown>;
}

interface CloseDocumentOptions extends SaveDocumentOptions {
  closeTimeoutMs?: number;
  setInputBlocked(blocked: boolean): void;
  onFailure?(error: unknown): void;
}

async function bounded<T>(operation: () => Promise<T>, timeoutMs: number, reason: string, signal?: AbortSignal): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  let onAbort: (() => void) | undefined;
  try {
    signal?.throwIfAborted();
    return await Promise.race([
      Promise.resolve().then(() => {
        signal?.throwIfAborted();
        return operation();
      }),
      new Promise<never>((_resolve, reject) => {
        timer = setTimeout(() => reject(new Error(reason)), timeoutMs);
      }),
      new Promise<never>((_resolve, reject) => {
        onAbort = () => reject(signal?.reason);
        signal?.addEventListener("abort", onAbort, { once: true });
      }),
    ]);
  } finally {
    clearTimeout(timer);
    if (onAbort) signal?.removeEventListener("abort", onAbort);
  }
}

export async function saveDocument(
  editor: InputEditor | null,
  api: SaveApi,
  options: SaveDocumentOptions,
): Promise<void> {
  let step = "started";
  const controller = new AbortController();
  const progress = (patch: Record<string, unknown>) => options.onProgress?.({ saveStep: step, ...patch });
  progress({
    saveStartedAt: Date.now(), saveFlushedAt: null, saveInputFlushed: null, saveStateVectorBytes: null,
    savePostedAt: null, saveResult: null, saveError: null, saveFailedStep: null,
    saveConnectionId: null, saveSequence: null,
  });
  try {
    options.signal?.throwIfAborted();
    if (!editor) throw new Error("editorNotReady");
    step = "flush";
    const flushDeadline = Date.now() + (options.inputTimeoutMs ?? 3000);
    if (!options.DontTerminateEdit) {
      await bounded(() => editor.flushPendingInput(), options.inputTimeoutMs ?? 3000, "flushPendingInputTimeout", options.signal);
    }
    options.signal?.throwIfAborted();
    options.onFlushed?.();
    step = "flushed";
    progress({ saveFlushedAt: Date.now(), saveInputFlushed: !options.DontTerminateEdit });
    step = "drain";
    while (options.hasPendingResync?.() || editor.hasPendingUpdates()) {
      const remaining = flushDeadline - Date.now();
      if (remaining <= 0) throw new Error("notSynced");
      await bounded(() => new Promise<void>((resolve) => setTimeout(resolve, Math.min(50, remaining))), remaining, "notSynced", options.signal);
    }
    options.signal?.throwIfAborted();
    if (Date.now() >= flushDeadline) throw new Error("notSynced");
    step = "sequence";
    const barrier = options.captureBarrier?.();
    if (!barrier) throw new Error("notSynced");
    const capturedBarrier = { connectionId: barrier.connectionId, seq: barrier.seq };
    progress({ saveConnectionId: barrier.connectionId, saveSequence: barrier.seq });
    step = "posted";
    await bounded(() => {
      const saving = api.save(capturedBarrier, controller.signal);
      progress({ savePostedAt: Date.now() });
      return saving;
    }, options.saveTimeoutMs ?? 10_000, "saveTimeout", options.signal);
    options.signal?.throwIfAborted();
    step = "result";
    progress({ saveResult: "saved" });
  } catch (error) {
    controller.abort();
    const message = error instanceof Error ? error.message : "saveFailed";
    progress({ saveStep: "result", saveResult: "failed", saveError: message, saveFailedStep: step });
    console.error(`WOPI save failed at ${step}: ${message}`);
    throw error;
  }
}

export async function closeDocument(
  editor: InputEditor | null,
  api: SaveApi & { close(signal?: AbortSignal): Promise<void> },
  options: CloseDocumentOptions,
  finish: () => void,
): Promise<void> {
  options.setInputBlocked(true);
  const controller = new AbortController();
  try {
    await saveDocument(editor, api, { ...options, DontTerminateEdit: false, DontSaveIfUnmodified: false });
    options.signal?.throwIfAborted();
    await bounded(() => api.close(controller.signal), options.closeTimeoutMs ?? 10_000, "closeTimeout", options.signal);
    options.signal?.throwIfAborted();
    finish();
  } catch (error) {
    controller.abort();
    options.setInputBlocked(false);
    options.onFailure?.(error);
    throw error;
  }
}

export function finishClose(
  config: { closePostMessage: boolean; closeUrl: string },
  bridge: { enabled: boolean; uiClose(): void },
  win: { top: { location: { href: string } } | null },
  showClosed: () => void,
): void {
  if (config.closePostMessage && bridge.enabled) bridge.uiClose();
  else if (config.closeUrl !== "" && win.top) win.top.location.href = config.closeUrl;
  else showClosed();
}
