import type { SaveOptions } from "./postMessage";

interface InputEditor {
  flushPendingInput(): Promise<void>;
  encodeStateVector(): Promise<Uint8Array>;
}

export async function saveDocument(
  editor: InputEditor,
  api: { save(stateVector: Uint8Array): Promise<unknown> },
  options: SaveOptions & { isModified(): boolean },
): Promise<void> {
  if (!options.DontTerminateEdit) await editor.flushPendingInput();
  if (options.DontSaveIfUnmodified && !options.isModified()) return;
  await api.save(await editor.encodeStateVector());
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
