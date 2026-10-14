import { PptxCommandAdmissionError } from '../commands/createPptxCommandStore';
import { peerlessKnownState } from './peerlessViewStore';
import type { PeerlessPresentation } from './usePeerlessPresentation';

export async function peerlessClipboard(run: PeerlessPresentation, cut: boolean): Promise<void> {
  run.require();
  await run.store.fence();
  run.require();
  const view = run.store.getSnapshot();
  const selection = view.selection;
  if (!selection || selection.kind === 'object' || selection.anchor === selection.focus || view.composition || cut && view.editState.readOnly) return;
  const captured = JSON.stringify(selection);
  const result = (await run.client.read({ slideIds: [selection.slideId] })).result;
  run.require();
  if (!result.ok) throw new Error(result.failure.message);
  if (captured !== JSON.stringify(run.store.getSnapshot().selection)) throw new PptxCommandAdmissionError('target-changed');
  const story = result.stories.find((story) => story.storyId === selection.storyId && story.shapeId === selection.shapeId);
  if (!story) throw new PptxCommandAdmissionError('target-changed');
  await navigator.clipboard.writeText(story.text.slice(Math.min(selection.anchor, selection.focus), Math.max(selection.anchor, selection.focus)));
  run.require();
  if (!cut) return;
  await run.store.fence();
  run.require();
  if (run.store.getSnapshot().composition || !run.inputEnabled || !peerlessKnownState(run.store) || result.version !== run.client.state.version ||
    captured !== JSON.stringify(run.store.getSnapshot().selection)) throw new PptxCommandAdmissionError('target-changed');
  await run.store.dispatch({ kind: 'replaceText', payload: { at: { source: 'confirmed', selection },
    text: '', inputTime: performance.now(), undoGroup: 'cut' } });
}
