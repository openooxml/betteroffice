import { defaultInsertLayoutPath } from '../slideLayouts';
import { invokePeerless } from './peerlessObjectCommands';
import type { PeerlessPresentation } from './usePeerlessPresentation';

export function peerlessInsertSlide(run: PeerlessPresentation, layout?: string | null) {
  run.require();
  const { slides, layouts, activeSlideId } = run.client.state;
  const active = slides.find((slide) => slide.id === activeSlideId);
  const index = active ? active.index + 1 : 0;
  const layoutPartPath =
    layout === undefined
      ? defaultInsertLayoutPath(layouts, active?.layoutPartPath)
      : layout ?? undefined;
  return run.store.dispatch({
    kind: 'insertSlide',
    payload: {
      index,
      ...(layoutPartPath === undefined ? {} : { layoutPartPath }),
      inputTime: performance.now(),
      undoGroup: 'slide',
    },
  });
}

export function peerlessMoveSlide(run: PeerlessPresentation, id: string, index: number) {
  return invokePeerless(run, { method: 'moveSlide', args: [id, index] });
}

export function peerlessDeleteSlide(run: PeerlessPresentation, id: string) {
  return invokePeerless(run, { method: 'deleteSlide', args: [id] });
}
