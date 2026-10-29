import type { SaveOptions } from "./postMessage";
import type { SaveBarrier, SocketSaveBarrier } from "../../src/relay/sequence";

interface InputEditor {
  flushPendingInput(): Promise<void>;
  hasPendingUpdates(): boolean;
}

interface RecoveryOptions {
  signal?: AbortSignal;
  hasUnacknowledgedUpdates(): boolean;
  upload?(bytes: Uint8Array, signal: AbortSignal): Promise<{ copyName: string }>;
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
  whenClosed?: Promise<void>;
  finalSaved?: boolean;
  onFinalSaved?(): void;
  setInputBlocked(blocked: boolean): void;
  onFailure?(error: unknown): void;
}

async function bounded<T>(operation: () => Promise<T>, timeoutMs: number, reason: string | Error, signal?: AbortSignal): Promise<T> {
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
        timer = setTimeout(() => reject(typeof reason === "string" ? new Error(reason) : reason), timeoutMs);
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

export async function recoverDocument(
  editor: (InputEditor & { serialize(): Promise<Uint8Array> }) | null,
  options: RecoveryOptions,
): Promise<{ bytes: Uint8Array; copyName?: string } | null> {
  const controller = new AbortController();
  try {
    if (!editor) throw new Error("The editor is not ready to preserve local work.");
    await bounded(() => editor.flushPendingInput(), 3000, "Recovery input flush timed out.", options.signal);
    options.signal?.throwIfAborted();
    if (!editor.hasPendingUpdates() && !options.hasUnacknowledgedUpdates()) return null;
    const bytes = await bounded(() => editor.serialize(), 10_000, "Recovery serialization timed out.", options.signal);
    options.signal?.throwIfAborted();
    if (!options.upload) return { bytes };
    const result = await bounded(() => options.upload!(bytes, controller.signal), 10_000, "Recovery upload timed out.", options.signal);
    options.signal?.throwIfAborted();
    return { bytes, copyName: result.copyName };
  } finally {
    controller.abort();
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
  const timeout = new Error("closeTimeout");
  let saved = options.finalSaved ?? false;
  try {
    if (!saved) {
      await saveDocument(editor, api, { ...options, DontTerminateEdit: false, DontSaveIfUnmodified: false });
      saved = true;
      options.onFinalSaved?.();
    }
    options.signal?.throwIfAborted();
    await bounded(() => {
      const request = api.close(controller.signal);
      return options.whenClosed ? Promise.race([request, options.whenClosed]) : request;
    }, options.closeTimeoutMs ?? 10_000, timeout, options.signal);
    options.signal?.throwIfAborted();
  } catch (error) {
    if (!saved) options.setInputBlocked(false);
    const aborted = options.signal?.aborted || (error instanceof Error && error.name === "AbortError");
    if (!saved || (error !== timeout && !aborted)) {
      options.onFailure?.(error);
      throw error;
    }
  } finally {
    controller.abort();
  }
  finish();
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
