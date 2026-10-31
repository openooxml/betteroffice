import type { SaveOptions } from "./postMessage";
import type { SaveBarrier, SocketSaveBarrier } from "../../src/relay/sequence";

interface InputEditor {
  flushPendingInput(options?: { terminateEdit?: boolean }): Promise<void>;
  hasPendingUpdates(): boolean;
}

interface InputFlushOptions {
  signal?: AbortSignal;
  hasActiveComposition?(): boolean;
  blockInput?(): void;
}

interface RecoveryOptions extends InputFlushOptions {
  onFlushed?(): void;
  hasUnsavedLocalChanges?(): boolean;
  hasUnacknowledgedUpdates(): boolean;
  upload?(bytes: Uint8Array, signal: AbortSignal): Promise<{ copyName: string }>;
}

interface SaveDocumentOptions extends SaveOptions {
  signal?: AbortSignal;
  captureBarrier?(): SocketSaveBarrier | null;
  canWrite?(): boolean;
  hasActiveComposition?(): boolean;
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
  preserveLocalWork?(): Promise<boolean>;
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

async function waitForComposition(
  hasActiveComposition: (() => boolean) | undefined,
  deadline: number,
  reason: string,
  signal?: AbortSignal,
): Promise<void> {
  while (hasActiveComposition?.()) {
    const remaining = deadline - Date.now();
    if (remaining <= 0) throw new Error(reason);
    await bounded(() => new Promise<void>((resolve) => setTimeout(resolve, Math.min(50, remaining))), remaining, reason, signal);
  }
}

export async function flushDocumentInput(
  editor: InputEditor | null,
  options: InputFlushOptions,
  compositionError = "compositionActive",
  flushError = "flushPendingInputTimeout",
): Promise<void> {
  if (!editor) throw new Error("The editor is not ready to preserve local work.");
  const deadline = Date.now() + 3000;
  do {
    await waitForComposition(options.hasActiveComposition, deadline, compositionError, options.signal);
    options.signal?.throwIfAborted();
    options.blockInput?.();
    await bounded(() => editor.flushPendingInput(), Math.max(0, deadline - Date.now()), flushError, options.signal);
    options.signal?.throwIfAborted();
  } while (options.hasActiveComposition?.());
}

export async function recoverDocument(
  editor: (InputEditor & { serialize(): Promise<Uint8Array> }) | null,
  options: RecoveryOptions,
): Promise<{ bytes: Uint8Array; copyName?: string } | null> {
  const controller = new AbortController();
  try {
    if (!editor) throw new Error("The editor is not ready to preserve local work.");
    await flushDocumentInput(editor, options, "Recovery composition timed out. Finish typing, then retry recovery.", "Recovery input flush timed out.");
    options.onFlushed?.();
    if (!options.hasUnsavedLocalChanges?.() && !editor.hasPendingUpdates() && !options.hasUnacknowledgedUpdates()) return null;
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
  const checkWritable = () => {
    if (options.canWrite?.() === false) throw new Error("notAllowed");
  };
  progress({
    saveStartedAt: Date.now(), saveFlushedAt: null, saveInputFlushed: null, saveStateVectorBytes: null,
    savePostedAt: null, saveResult: null, saveError: null, saveFailedStep: null,
    saveConnectionId: null, saveSequence: null,
  });
  try {
    options.signal?.throwIfAborted();
    checkWritable();
    if (!editor) throw new Error("editorNotReady");
    const flushDeadline = Date.now() + (options.inputTimeoutMs ?? 3000);
    do {
      step = "composition";
      checkWritable();
      await waitForComposition(options.hasActiveComposition, flushDeadline, "compositionActive", options.signal);
      checkWritable();
      step = "flush";
      await bounded(() => editor.flushPendingInput({ terminateEdit: !options.DontTerminateEdit }), Math.max(0, flushDeadline - Date.now()), "flushPendingInputTimeout", options.signal);
      options.signal?.throwIfAborted();
      options.onFlushed?.();
      step = "flushed";
      progress({ saveFlushedAt: Date.now(), saveInputFlushed: true });
      step = "drain";
      for (;;) {
        checkWritable();
        if (!options.hasPendingResync?.() && !editor.hasPendingUpdates()) break;
        const remaining = flushDeadline - Date.now();
        if (remaining <= 0) throw new Error("notSynced");
        await bounded(() => new Promise<void>((resolve) => setTimeout(resolve, Math.min(50, remaining))), remaining, "notSynced", options.signal);
      }
    } while (options.hasActiveComposition?.());
    options.signal?.throwIfAborted();
    checkWritable();
    if (Date.now() >= flushDeadline) throw new Error("notSynced");
    step = "sequence";
    const barrier = options.captureBarrier?.();
    if (!barrier) throw new Error("notSynced");
    const capturedBarrier = { connectionId: barrier.connectionId, seq: barrier.seq };
    progress({ saveConnectionId: barrier.connectionId, saveSequence: barrier.seq });
    step = "posted";
    await bounded(() => {
      checkWritable();
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
      if (options.canWrite?.() !== false) {
        await saveDocument(editor, api, { ...options, DontTerminateEdit: false, DontSaveIfUnmodified: false });
      } else if (options.preserveLocalWork) {
        if (!await options.preserveLocalWork()) return;
      } else if (editor?.hasPendingUpdates()) {
        throw new Error("localRecoveryRequired");
      }
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
