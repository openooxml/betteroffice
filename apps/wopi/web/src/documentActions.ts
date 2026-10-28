import type { SaveOptions } from "./postMessage";

interface InputEditor {
  flushPendingInput(): Promise<void>;
  encodeStateVector(): Promise<Uint8Array>;
}

interface SaveDocumentOptions extends SaveOptions {
  isModified(): boolean;
  onFlushed?(): void;
  onProgress?(patch: Record<string, unknown>): void;
  inputTimeoutMs?: number;
  saveTimeoutMs?: number;
}

async function bounded<T>(operation: () => Promise<T>, timeoutMs: number, reason: string): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      Promise.resolve().then(operation),
      new Promise<never>((_resolve, reject) => {
        timer = setTimeout(() => reject(new Error(reason)), timeoutMs);
      }),
    ]);
  } finally {
    clearTimeout(timer);
  }
}

export async function saveDocument(
  editor: InputEditor | null,
  api: { save(stateVector: Uint8Array, signal?: AbortSignal): Promise<unknown> },
  options: SaveDocumentOptions,
): Promise<void> {
  let step = "started";
  const controller = new AbortController();
  const progress = (patch: Record<string, unknown>) => options.onProgress?.({ saveStep: step, ...patch });
  progress({
    saveStartedAt: Date.now(), saveFlushedAt: null, saveInputFlushed: null, saveStateVectorBytes: null,
    savePostedAt: null, saveResult: null, saveError: null, saveFailedStep: null,
  });
  try {
    if (!editor) throw new Error("editorNotReady");
    step = "flush";
    if (!options.DontTerminateEdit) {
      await bounded(() => editor.flushPendingInput(), options.inputTimeoutMs ?? 3000, "flushPendingInputTimeout");
    }
    options.onFlushed?.();
    step = "flushed";
    progress({ saveFlushedAt: Date.now(), saveInputFlushed: !options.DontTerminateEdit });
    if (options.DontSaveIfUnmodified && !options.isModified()) {
      step = "result";
      progress({ saveResult: "skipped" });
      return;
    }
    step = "stateVector";
    const stateVector = await bounded(() => editor.encodeStateVector(), options.inputTimeoutMs ?? 3000, "encodeStateVectorTimeout");
    progress({ saveStateVectorBytes: stateVector.byteLength });
    step = "posted";
    await bounded(() => {
      const saving = api.save(stateVector, controller.signal);
      progress({ savePostedAt: Date.now() });
      return saving;
    }, options.saveTimeoutMs ?? 10_000, "saveTimeout");
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
  editor: Pick<InputEditor, "flushPendingInput">,
  actions: { save(): Promise<void>; close(): Promise<void> },
  isModified: () => boolean,
  finish: () => void,
): Promise<void> {
  await editor.flushPendingInput();
  if (isModified()) await actions.save();
  await actions.close();
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
