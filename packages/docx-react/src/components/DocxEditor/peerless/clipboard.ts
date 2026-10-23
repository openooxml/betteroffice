import { clipboardFragmentHasPictures } from '@betteroffice/docx/yrs';
import type { DocxClipboardSlice, DocxCommandSelection, YrsSelectionText } from '@betteroffice/docx/yrs';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
import { clipboardFragmentHtml } from '../internals/clipboardHtml';
import { flushAdmittedInput } from './inputAdmission';
import type { PeerlessViewIdentity } from './viewStore';

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
  const identity = { ...state.identity };
  if (!state.workerSelection) return null;
  const selection = structuredClone(state.workerSelection);
  const { anchor, focus } = selection;
  if (anchor.story !== focus.story) throw new Error('Cross-story selection is unavailable');
  const info = await options.readSelection(selection);
  if (!options.isCurrent()) throw new Error('Document was replaced');
  if (!sameIdentity(identity, binding.viewStore.getSnapshot().identity) || state.version !== binding.viewStore.getSnapshot().version) throw new Error('Selection read was superseded');
  return { selection, text: info.selectedText, version: state.version, identity, info };
}

export async function readWorkerClipboardSlice(options: SelectionRead & {
  readClipboardSlice(selection: DocxCommandSelection): Promise<DocxClipboardSlice>;
}) {
  const { binding } = options;
  await flushAdmittedInput(binding);
  if (!options.isCurrent()) throw new Error('Document was replaced');
  const state = binding.viewStore.getSnapshot();
  const identity = { ...state.identity };
  if (!state.workerSelection) return null;
  const selection = structuredClone(state.workerSelection);
  const { anchor, focus } = selection;
  if (anchor.story !== focus.story) throw new Error('Cross-story selection is unavailable');
  const slice = await options.readClipboardSlice(selection);
  if (!options.isCurrent()) throw new Error('Document was replaced');
  if (!sameIdentity(identity, binding.viewStore.getSnapshot().identity) || state.version !== binding.viewStore.getSnapshot().version) throw new Error('Selection read was superseded');
  return { selection, slice, version: state.version, identity, html: slice.fragment ? clipboardFragmentHtml(slice.fragment) : null };
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
  identity: PeerlessViewIdentity;
  html: string | null;
}

function sameIdentity(a: PeerlessViewIdentity, b: PeerlessViewIdentity): boolean {
  return a?.session === b?.session && a?.incarnation === b?.incarnation && a?.epoch === b?.epoch;
}

function selectionKey(identity: PeerlessViewIdentity, version: number, selection: DocxCommandSelection | null | undefined): string | null {
  if (!selection) return null;
  const { anchor: a, focus: f } = selection;
  if (a.story !== f.story || (a.paraId === f.paraId && a.offset === f.offset)) return null;
  return JSON.stringify([identity?.session, identity?.incarnation, identity?.epoch, version, a.story, a.paraId, a.offset, f.paraId, f.offset]);
}

function snapshotKey(binding: PeerlessEditSessionBinding): string | null {
  const { identity, version, workerSelection } = binding.viewStore.getSnapshot();
  return selectionKey(identity, version, workerSelection);
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
  let cached: { key: string; capture: ClipboardCapture | null; error?: unknown } | null = null;
  let reading: Promise<void> | null = null;
  let disposed = false;
  const dispose = () => {
    disposed = true;
    unsubscribe();
    cached = null;
  };
  const schedule = () => {
    if (disposed) return;
    if (!options.isCurrent()) { dispose(); return; }
    const key = snapshotKey(binding);
    if (cached && key !== cached.key) cached = null;
    if (!key || key === cached?.key || reading) return;
    reading = readWorkerClipboardSlice(options).then(captured => {
      if (disposed) return;
      const capturedKey = captured ? selectionKey(captured.identity, captured.version, captured.selection) : key;
      if (capturedKey) cached = { key: capturedKey, capture: (captured?.slice.text || captured?.slice.fragment) ? captured : null };
    }, error => {
      if (!disposed) cached = { key, capture: null, error };
    }).finally(() => { reading = null; schedule(); });
  };
  const unsubscribe = binding.viewStore.subscribe(schedule);
  const take = (): ClipboardCapture | null =>
    !disposed && cached && options.isCurrent() && inputSettled(binding) && snapshotKey(binding) === cached.key
      ? cached.capture : null;
  schedule();
  return {
    take,
    async read(): Promise<ClipboardCapture | null> {
      await flushAdmittedInput(binding);
      const key = snapshotKey(binding);
      if (!key) return readWorkerClipboardSlice(options);
      const identity = { ...binding.viewStore.getSnapshot().identity };
      schedule();
      while (!disposed && options.isCurrent()) {
        if (!sameIdentity(identity, binding.viewStore.getSnapshot().identity)) return null;
        if (snapshotKey(binding) !== key) throw new Error('Selection read was superseded');
        if (reading) { await reading; continue; }
        if (cached?.key === snapshotKey(binding) && cached?.error) throw cached.error;
        return take();
      }
      return null;
    },
    dispose,
  };
}

/** @internal Fills a copy event's clipboardData from a capture; false when there is nothing to write. */
export function writeClipboardCapture(data: DataTransfer, capture: ClipboardCapture): boolean {
  if (!capture.slice.text && !capture.slice.fragment) return false;
  data.setData('text/plain', capture.slice.text);
  if (capture.html) data.setData('text/html', capture.html);
  return true;
}

export async function copyWorkerSelection(options: SelectionRead & { write(text: string): Promise<void>; cut?: boolean;
  readClipboardSlice?(selection: DocxCommandSelection): Promise<DocxClipboardSlice>;
  writeSlice?(text: string, html: string): Promise<void>;
  data?: DataTransfer | null; prefetched?: ClipboardCapture | null;
  readCapture?(): Promise<ClipboardCapture | null>;
  writeNative?(capture: ClipboardCapture): boolean }) {
  const { binding } = options;
  const identity = { ...binding.viewStore.getSnapshot().identity };
  if (options.cut && binding.viewStore.getSnapshot().editState?.readOnly) return false;
  let captured: { selection: DocxCommandSelection; version: number; identity?: PeerlessViewIdentity };
  if (options.data && options.prefetched) {
    if (!sameIdentity(options.prefetched.identity, binding.viewStore.getSnapshot().identity) ||
      !writeClipboardCapture(options.data, options.prefetched)) return false;
    captured = options.prefetched;
  } else if (options.readClipboardSlice || options.readCapture) {
    const sliceCaptured = options.readCapture ? await options.readCapture()
      : await readWorkerClipboardSlice({ ...options, readClipboardSlice: options.readClipboardSlice! });
    if (!sliceCaptured || (!sliceCaptured.slice.text && !sliceCaptured.slice.fragment)) return false;
    if (!options.isCurrent() || !sameIdentity(identity, sliceCaptured.identity) ||
      !sameIdentity(sliceCaptured.identity, binding.viewStore.getSnapshot().identity)) return false;
    if (!options.writeNative?.(sliceCaptured)) {
      if (sliceCaptured.html) {
        try {
          if (!options.writeSlice) throw new Error('Rich clipboard writing is unavailable');
          await options.writeSlice(sliceCaptured.slice.text, sliceCaptured.html);
        } catch (error) {
          if (clipboardFragmentHasPictures(sliceCaptured.slice.fragment)) throw error;
          await options.write(sliceCaptured.slice.text);
        }
      } else await options.write(sliceCaptured.slice.text);
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
  if (captured.identity && !sameIdentity(captured.identity, binding.viewStore.getSnapshot().identity)) return false;
  if (binding.viewStore.getSnapshot().editState?.tableContext?.hasMultiCellSelection) return true;
  if (binding.viewStore.getSnapshot().version !== captured.version &&
    (!captured.selection.anchor.sticky || !captured.selection.focus.sticky)) return false;
  const result = binding.commands.submit({ kind: 'replaceSelection', payload: { text: '', selection: captured.selection } },
    { source: 'paste', fence: true });
  if (result.status !== 'submitted') return false;
  const outcome = await result.ticket.ack;
  return outcome.status === 'applied' && outcome.committed !== false;
}
