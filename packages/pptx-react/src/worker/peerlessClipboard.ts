import { PptxCommandAdmissionError } from '../commands/createPptxCommandStore';
import { indexGeometry, type PresentationCommandAck } from '@betteroffice/pptx/worker';
import { peerlessKnownState } from './peerlessViewStore';
import type { PeerlessPresentation } from './usePeerlessPresentation';

export async function peerlessClipboard(run: PeerlessPresentation, cut: boolean, immediate = false): Promise<PresentationCommandAck | null> {
  run.require();
  if (typeof navigator === 'undefined' || typeof navigator.clipboard?.writeText !== 'function')
    throw new PptxCommandAdmissionError('clipboard-unavailable');
  if (!immediate && !peerlessKnownState(run.store)) {
    await run.store.fence();
    run.require();
  }
  const view = run.store.getSnapshot();
  const selection = view.selection;
  if (!selection || selection.kind === 'object' || selection.kind === 'shape' || selection.anchor === selection.focus || view.composition || cut && view.editState.readOnly)
    throw new PptxCommandAdmissionError('target-changed');
  const captured = JSON.stringify(selection);
  if (!peerlessKnownState(run.store) || !view.confirmed || view.confirmed.frame.slideId !== selection.slideId)
    throw new PptxCommandAdmissionError('target-changed');
  const box = indexGeometry(view.confirmed.displayList).hits.map((hit) => hit.primitive)
    .find((primitive) => primitive.kind === 'textBox' && primitive.storyId === selection.storyId && primitive.shapeId === selection.shapeId);
  if (!box || box.kind !== 'textBox') throw new PptxCommandAdmissionError('target-changed');
  const text = box.paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join('')).join('\n');
  const version = run.client.state.version;
  await navigator.clipboard.writeText(text.slice(Math.min(selection.anchor, selection.focus), Math.max(selection.anchor, selection.focus)));
  run.require();
  if (!cut) return null;
  await run.store.fence();
  run.require();
  if (run.store.getSnapshot().composition || !run.inputEnabled || !peerlessKnownState(run.store) || version !== run.client.state.version ||
    captured !== JSON.stringify(run.store.getSnapshot().selection)) throw new PptxCommandAdmissionError('target-changed');
  return run.store.dispatch({ kind: 'replaceText', payload: { at: { source: 'confirmed', selection },
    text: '', inputTime: performance.now(), undoGroup: 'cut' } });
}

/** @internal */
export async function peerlessPaste(run: PeerlessPresentation) {
  run.require();
  if (typeof navigator === 'undefined' || typeof navigator.clipboard?.readText !== 'function')
    throw new PptxCommandAdmissionError('clipboard-unavailable');
  const view = run.store.getSnapshot();
  const selection = view.selection;
  if (!selection || selection.kind === 'object' || selection.kind === 'shape' || view.composition)
    throw new PptxCommandAdmissionError('target-changed');
  const captured = JSON.stringify(selection);
  const version = run.client.state.version;
  const text = await navigator.clipboard.readText();
  await run.store.fence();
  run.require();
  if (run.store.getSnapshot().composition || !run.inputEnabled || !peerlessKnownState(run.store) ||
    version !== run.client.state.version || captured !== JSON.stringify(run.store.getSnapshot().selection))
    throw new PptxCommandAdmissionError('target-changed');
  if (!text) return null;
  return run.store.dispatch({ kind: 'replaceText', payload: { at: { source: 'confirmed', selection },
    text, inputTime: performance.now(), undoGroup: `paste-${++nextPasteGroup}` } });
}

let nextPasteGroup = 0;
