import type { PresentationCommandAck, PresentationCommandMutation } from '@betteroffice/pptx';
import type { PptxCommandArgs } from '../commands/types';
import type { PeerlessPresentation } from './usePeerlessPresentation';

export function invokePeerless(run: PeerlessPresentation, mutation: PresentationCommandMutation,
  options?: { reportRefusal?: boolean }): Promise<PresentationCommandAck> {
  run.require();
  return run.store.dispatch({ kind: 'invoke', payload: { ...mutation, inputTime: performance.now(), undoGroup: 'object' } }, options);
}

export async function peerlessObjectCommand(run: PeerlessPresentation, action:
  { kind: 'fill'; color: string | null } | { kind: 'strokeColor'; color: string | null } |
  { kind: 'strokeWidth'; points: number | null } | { kind: 'adjust'; name: string; value: number } |
  { kind: 'zOrder'; value: PptxCommandArgs['zOrder']['value'] } | { kind: 'remove' }) {
  const selected = run.store.getSnapshot().selection;
  if (selected?.kind !== 'object') return null;
  await run.store.fence();
  run.require();
  const view = run.store.getSnapshot();
  if (view.selection?.kind !== 'object' || view.selection.shapeId !== selected.shapeId || view.selection.slideId !== selected.slideId) return null;
  const ids = [selected.slideId, selected.shapeId] as const;
  const context = view.editState.object;
  let mutation: PresentationCommandMutation;
  switch (action.kind) {
    case 'fill': mutation = { method: 'setShapeFill', args: [...ids, action.color] }; break;
    case 'strokeColor': mutation = { method: 'setShapeStroke', args: [...ids, action.color === null ? {} : { color: action.color }] }; break;
    case 'strokeWidth': mutation = { method: 'setShapeStroke', args: [...ids, action.points === null ? {} : { widthPt: action.points }] }; break;
    case 'adjust': mutation = { method: 'setShapeAdjust', args: [...ids,
      { ...context?.adjustments, [action.name]: action.value }] }; break;
    case 'remove': mutation = { method: 'removeShape', args: [...ids] }; break;
    case 'zOrder': {
      const method = { front: 'bringShapeToFront', forward: 'bringShapeForward', backward: 'sendShapeBackward', back: 'sendShapeToBack' } as const;
      mutation = { method: method[action.value], args: [...ids] }; break;
    }
  }
  return invokePeerless(run, mutation);
}
