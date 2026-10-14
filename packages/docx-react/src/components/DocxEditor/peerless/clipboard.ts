import type { DocxCommandSelection, YrsSelectionText } from '@betteroffice/docx/yrs';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';
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

export async function copyWorkerSelection(options: SelectionRead & { write(text: string): Promise<void>; cut?: boolean }) {
  const { binding } = options;
  if (options.cut && binding.viewStore.getSnapshot().editState?.readOnly) return false;
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
