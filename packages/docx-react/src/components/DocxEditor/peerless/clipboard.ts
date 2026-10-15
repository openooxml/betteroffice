import type { DocxClipboardFragment, DocxClipboardSlice, DocxCommandSelection, YrsSelectionText } from '@betteroffice/docx/yrs';
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

export function clipboardSliceHtml(fragment: DocxClipboardFragment): string {
  return clipboardFragmentHtml(fragment);
}

export function isPasteFragmentAllowed(commands: {
  getState(id: 'insertTable', args?: unknown): { enabled: boolean };
}): boolean {
  try {
    return commands.getState('insertTable', { rows: 1, columns: 1 }).enabled === true;
  } catch {
    return false;
  }
}

export async function copyWorkerSelection(options: SelectionRead & { write(text: string): Promise<void>; cut?: boolean;
  readClipboardSlice?(selection: DocxCommandSelection): Promise<DocxClipboardSlice>;
  writeSlice?(text: string, html: string): Promise<void> }) {
  const { binding } = options;
  if (options.cut && binding.viewStore.getSnapshot().editState?.readOnly) return false;
  if (options.readClipboardSlice) {
    const captured = await readWorkerClipboardSlice({ ...options, readClipboardSlice: options.readClipboardSlice });
    if (!captured?.slice.text) return false;
    if (captured.slice.fragment && options.writeSlice) {
      await options.writeSlice(captured.slice.text, clipboardFragmentHtml(captured.slice.fragment));
    } else {
      await options.write(captured.slice.text);
    }
    if (!options.isCurrent()) return false;
    if (!options.cut) return true;
    if (binding.viewStore.getSnapshot().editState?.readOnly) return false;
    const tableContext = binding.viewStore.getSnapshot().editState?.tableContext as { hasMultiCellSelection?: boolean } | null | undefined;
    if (tableContext?.hasMultiCellSelection) return true;
    if (binding.viewStore.getSnapshot().version !== captured.version &&
      (!captured.selection.anchor.sticky || !captured.selection.focus.sticky)) return false;
    const result = binding.commands.submit({ kind: 'replaceSelection', payload: { text: '', selection: captured.selection } },
      { source: 'paste', fence: true });
    if (result.status !== 'submitted') return false;
    const outcome = await result.ticket.ack;
    return outcome.status === 'applied' && outcome.committed !== false;
  }
  const captured = await readWorkerSelection(options);
  if (!captured?.text) return false;
  await options.write(captured.text);
  if (!options.isCurrent()) return false;
  if (!options.cut) return true;
  if (binding.viewStore.getSnapshot().editState?.readOnly) return false;
  if (binding.viewStore.getSnapshot().version !== captured.version &&
    (!captured.selection.anchor.sticky || !captured.selection.focus.sticky)) return false;
  const result = binding.commands.submit({ kind: 'replaceSelection', payload: { text: '', selection: captured.selection } },
    { source: 'paste', fence: true });
  if (result.status !== 'submitted') return false;
  const outcome = await result.ticket.ack;
  return outcome.status === 'applied' && outcome.committed !== false;
}
