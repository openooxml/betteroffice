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

export async function copyWorkerSelection(options: SelectionRead & { write(text: string): Promise<void>; cut?: boolean;
  readClipboardSlice?(selection: DocxCommandSelection): Promise<DocxClipboardSlice>;
  writeSlice?(text: string, html: string): Promise<void> }) {
  const { binding } = options;
  if (options.cut && binding.viewStore.getSnapshot().editState?.readOnly) return false;
  let captured: { selection: DocxCommandSelection; version: number };
  if (options.readClipboardSlice) {
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
