import type { DocxClipboardSlice, DocxCommandSelection, YrsSelectionText } from '@betteroffice/docx/yrs';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import { clipboardFragmentHtml } from '../internals/clipboardHtml';
import { flushAdmittedInput } from './inputAdmission';

interface SelectionRead {
  binding: PeerlessEditSessionBinding;
  readSelection(selection: DocxCommandSelection): Promise<YrsSelectionText>;
  isCurrent(): boolean;
}

export async function readWorkerSelection(options: SelectionRead) {
  const { binding } = options;
  await flushAdmittedInput(binding);
  if (!options.isCurrent()) throw new Error('Document was replaced');
  const state = binding.viewStore.getSnapshot();
  if (!state.workerSelection) return null;
  const selection = structuredClone(state.workerSelection);
  const { anchor, focus } = selection;
  if (anchor.story !== focus.story) throw new Error('Cross-story selection is unavailable');
  const info = await options.readSelection(selection);
  if (!options.isCurrent()) throw new Error('Document was replaced');
  if (state.version !== binding.viewStore.getSnapshot().version) throw new Error('Selection read was superseded');
  return { selection, text: info.selectedText, version: state.version, info };
}

export async function readWorkerClipboardSlice(options: SelectionRead & {
  readClipboardSlice(selection: DocxCommandSelection): Promise<DocxClipboardSlice>;
}) {
  const { binding } = options;
  await flushAdmittedInput(binding);
  if (!options.isCurrent()) throw new Error('Document was replaced');
  const state = binding.viewStore.getSnapshot();
  if (!state.workerSelection) return null;
  const selection = structuredClone(state.workerSelection);
  const { anchor, focus } = selection;
  if (anchor.story !== focus.story) throw new Error('Cross-story selection is unavailable');
  const slice = await options.readClipboardSlice(selection);
  if (!options.isCurrent()) throw new Error('Document was replaced');
  if (state.version !== binding.viewStore.getSnapshot().version) throw new Error('Selection read was superseded');
  return { selection, slice, version: state.version };
}

export function isPasteFragmentAllowed(commands: {
  getState(id: 'insertTable', args?: unknown): { enabled: boolean };
}): boolean {
  return commands.getState('insertTable', { rows: 1, columns: 1 }).enabled === true;
}

/** @internal A clipboard slice read for one selection at one document version. */
export interface ClipboardCapture {
  selection: DocxCommandSelection;
  slice: DocxClipboardSlice;
  version: number;
}

const CLIPBOARD_PREFETCH_SETTLE_MS = 100;

function selectionKey(version: number, selection: DocxCommandSelection | null | undefined): string | null {
  if (!selection) return null;
  const { anchor: a, focus: f } = selection;
  if (a.story !== f.story || (a.paraId === f.paraId && a.offset === f.offset)) return null;
  return `${version}\u0000${a.story}\u0000${a.paraId}\u0000${a.offset}\u0000${f.paraId}\u0000${f.offset}`;
}

function snapshotKey(binding: PeerlessEditSessionBinding): string | null {
  const { version, workerSelection } = binding.viewStore.getSnapshot();
  return selectionKey(version, workerSelection);
}

function inputSettled(binding: PeerlessEditSessionBinding): boolean {
  const admission = binding.admission?.getSnapshot();
  return !admission || (!admission.tail && admission.compositionSlot === null && admission.heldIntents === 0);
}

/** @internal Reads the settled selection's clipboard slice ahead of a copy so the copy event can fill its clipboardData. */
export function createClipboardPrefetch(options: SelectionRead & {
  readClipboardSlice(selection: DocxCommandSelection): Promise<DocxClipboardSlice>;
}) {
  const { binding } = options;
  let cached: { key: string; capture: ClipboardCapture } | null = null;
  let scheduled: string | null = null;
  let timer: ReturnType<typeof setTimeout> | undefined;
  let reading = false;
  let disposed = false;
  const dispose = () => {
    disposed = true;
    clearTimeout(timer);
    unsubscribe();
    cached = null;
  };
  const run = () => {
    scheduled = null;
    if (disposed || reading || !snapshotKey(binding)) return;
    reading = true;
    void readWorkerClipboardSlice(options).then((captured) => {
      const key = captured?.slice.text ? selectionKey(captured.version, captured.selection) : null;
      if (!disposed && captured && key) cached = { key, capture: captured };
    }, () => {}).finally(() => { reading = false; schedule(); });
  };
  const schedule = () => {
    if (disposed) return;
    if (!options.isCurrent()) { dispose(); return; }
    const key = snapshotKey(binding);
    if (!key || key === cached?.key || key === scheduled || reading) return;
    clearTimeout(timer);
    scheduled = key;
    timer = setTimeout(run, CLIPBOARD_PREFETCH_SETTLE_MS);
  };
  const unsubscribe = binding.viewStore.subscribe(schedule);
  schedule();
  return {
    take(): ClipboardCapture | null {
      return !disposed && cached && options.isCurrent() && inputSettled(binding) && snapshotKey(binding) === cached.key
        ? cached.capture : null;
    },
    dispose,
  };
}

/** @internal Fills a copy event's clipboardData from a capture; false when there is nothing to write. */
function writeCapture(data: DataTransfer, capture: ClipboardCapture): boolean {
  if (!capture.slice.text) return false;
  data.setData('text/plain', capture.slice.text);
  if (capture.slice.fragment) data.setData('text/html', clipboardFragmentHtml(capture.slice.fragment));
  return true;
}

export async function copyWorkerSelection(options: SelectionRead & { write(text: string): Promise<void>; cut?: boolean;
  readClipboardSlice?(selection: DocxCommandSelection): Promise<DocxClipboardSlice>;
  writeSlice?(text: string, html: string): Promise<void>;
  data?: DataTransfer | null; prefetched?: ClipboardCapture | null }) {
  const { binding } = options;
  if (options.cut && binding.viewStore.getSnapshot().editState?.readOnly) return false;
  let captured: { selection: DocxCommandSelection; version: number };
  if (options.data && options.prefetched) {
    if (!writeCapture(options.data, options.prefetched)) return false;
    captured = options.prefetched;
  } else if (options.readClipboardSlice) {
    const sliceCaptured = await readWorkerClipboardSlice({ ...options, readClipboardSlice: options.readClipboardSlice });
    if (!sliceCaptured?.slice.text) return false;
    if (sliceCaptured.slice.fragment && options.writeSlice) {
      await options.writeSlice(sliceCaptured.slice.text, clipboardFragmentHtml(sliceCaptured.slice.fragment));
    } else {
      await options.write(sliceCaptured.slice.text);
    }
    captured = sliceCaptured;
  } else {
    const textCaptured = await readWorkerSelection(options);
    if (!textCaptured?.text) return false;
    await options.write(textCaptured.text);
    captured = textCaptured;
  }
  if (!options.isCurrent()) return false;
  if (!options.cut) return true;
  if (binding.viewStore.getSnapshot().editState?.readOnly) return false;
  if (binding.viewStore.getSnapshot().editState?.tableContext?.hasMultiCellSelection) return true;
  if (binding.viewStore.getSnapshot().version !== captured.version &&
    (!captured.selection.anchor.sticky || !captured.selection.focus.sticky)) return false;
  const result = binding.commands.submit({ kind: 'replaceSelection', payload: { text: '', selection: captured.selection } },
    { source: 'paste', fence: true });
  if (result.status !== 'submitted') return false;
  const outcome = await result.ticket.ack;
  return outcome.status === 'applied' && outcome.committed !== false;
}
