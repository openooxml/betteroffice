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

export function jumpPeerlessSlide(run: PeerlessPresentation, jump: 'next' | 'previous' | 'first' | 'last'): void {
  if (!run.current) return;
  const slides = run.client.state.slides;
  const slideId = targets.get(run)?.slideId ?? run.store.getSnapshot().confirmed?.frame.slideId;
  const index = slides.findIndex((slide) => slide.id === slideId);
  if (index < 0) return;
  const next = jumpIndex(index, slides.length, jump);
  if (next !== index) void activatePeerlessSlide(run, slides[next]!.id).catch(() => {});
}

function jumpIndex(index: number, length: number, jump: 'next' | 'previous' | 'first' | 'last'): number {
  return jump === 'previous' ? Math.max(0, index - 1)
    : jump === 'next' ? Math.min(length - 1, index + 1)
    : jump === 'first' ? 0 : length - 1;
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
  const jump = event.key === 'PageUp' ? 'previous' : event.key === 'PageDown' ? 'next'
    : event.key === 'Home' ? 'first' : event.key === 'End' ? 'last' : null;
  const next = jump === null ? null : jumpIndex(index, slides.length, jump);
  if (next === null) return;
  event.preventDefault();
  if (next !== index) void activatePeerlessSlide(run, slides[next]!.id).catch(() => {});
}
