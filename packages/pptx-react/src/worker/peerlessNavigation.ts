import type { KeyboardEvent } from 'react';
import type { PresentationCommandAck } from '@betteroffice/pptx/worker';
import type { PeerlessPresentation } from './usePeerlessPresentation';
import { commitComposition } from './peerlessViewStore';

const targets = new WeakMap<PeerlessPresentation, { slideId: string }>();

export function activatePeerlessSlide(run: PeerlessPresentation, slideId: string): Promise<PresentationCommandAck> {
  const composition = run.store.getSnapshot().composition;
  if (composition) void commitComposition(run.store, composition.text, performance.now()).catch(() => {});
  const target = { slideId };
  targets.set(run, target);
  return run.store.dispatch({ kind: 'activateSlide', payload: { slideId } }).finally(() => {
    if (targets.get(run) === target) targets.delete(run);
  });
}

export function navigatePeerlessSlide(event: KeyboardEvent<HTMLDivElement>, run: PeerlessPresentation): void {
  const view = run.store.getSnapshot();
  if (!run.current || event.target !== event.currentTarget ||
    event.currentTarget.ownerDocument.activeElement !== event.currentTarget ||
    event.altKey || event.ctrlKey || event.metaKey || event.shiftKey || event.nativeEvent.isComposing ||
    view.focused || view.composition) return;
  const slides = run.client.state.slides;
  const slideId = targets.get(run)?.slideId ?? view.confirmed?.frame.slideId;
  const index = slides.findIndex((slide) => slide.id === slideId);
  if (index < 0) return;
  const next = event.key === 'PageUp' ? Math.max(0, index - 1)
    : event.key === 'PageDown' ? Math.min(slides.length - 1, index + 1)
    : event.key === 'Home' ? 0 : event.key === 'End' ? slides.length - 1 : null;
  if (next === null) return;
  event.preventDefault();
  if (next !== index) void activatePeerlessSlide(run, slides[next]!.id).catch(() => {});
}
