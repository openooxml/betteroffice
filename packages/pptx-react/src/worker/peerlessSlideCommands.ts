import { invokePeerless } from './peerlessObjectCommands';
import type { PeerlessPresentation } from './usePeerlessPresentation';

export function peerlessInsertSlide(run: PeerlessPresentation, layout?: string | null) {
  const active = run.client.state.slides.find((slide) => slide.id === run.client.state.activeSlideId);
  return invokePeerless(run, { method: 'insertSlide', args: [active ? active.index + 1 : 0,
    layout === undefined ? active?.layoutPartPath ?? undefined : layout ?? undefined] });
}

export function peerlessMoveSlide(run: PeerlessPresentation, id: string, index: number) {
  return invokePeerless(run, { method: 'moveSlide', args: [id, index] });
}

export function peerlessDeleteSlide(run: PeerlessPresentation, id: string) {
  return invokePeerless(run, { method: 'deleteSlide', args: [id] });
}
