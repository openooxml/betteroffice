import { PptxCommandAdmissionError } from '../commands/createPptxCommandStore';
import type { PresentationCommandAck } from '@betteroffice/pptx/worker';
import { peerlessKnownState } from './peerlessViewStore';
import type { PeerlessPresentation } from './usePeerlessPresentation';

export async function peerlessClipboard(run: PeerlessPresentation, cut: boolean, strict = false): Promise<PresentationCommandAck | null> {
  run.require();
  if (typeof navigator === 'undefined' || typeof navigator.clipboard?.writeText !== 'function') {
    if (strict) throw new PptxCommandAdmissionError('clipboard-unavailable');
    return null;
  }
  const view = run.store.getSnapshot();
  const selection = view.selection;
  if (!selection || selection.kind === 'object' || selection.kind === 'shape' || selection.anchor === selection.focus || view.composition || cut && view.editState.readOnly) {
    if (strict) throw new PptxCommandAdmissionError('target-changed');
    return null;
  }
  const captured = JSON.stringify(selection);
  const read = (async () => {
    await run.store.fence();
    run.require();
    const result = (await run.client.read({ slideIds: [selection.slideId] })).result;
    run.require();
    if (!result.ok) throw new Error(result.failure.message);
    if (captured !== JSON.stringify(run.store.getSnapshot().selection)) throw new PptxCommandAdmissionError('target-changed');
    const story = result.stories.find((story) => story.storyId === selection.storyId && story.shapeId === selection.shapeId);
    if (!story) throw new PptxCommandAdmissionError('target-changed');
    return { text: story.text.slice(Math.min(selection.anchor, selection.focus), Math.max(selection.anchor, selection.focus)), version: result.version };
  })();
  read.catch(() => {});
  if (typeof ClipboardItem === 'function' && typeof navigator.clipboard.write === 'function')
    await navigator.clipboard.write([new ClipboardItem({ 'text/plain': read.then(({ text }) => new Blob([text], { type: 'text/plain' })) })]);
  else await navigator.clipboard.writeText((await read).text);
  const { version } = await read;
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
