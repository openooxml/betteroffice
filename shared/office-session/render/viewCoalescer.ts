import type { FrameScheduler, ViewMessage } from './types';

/** @internal Latest view fields merge once per animation callback. */
export function createViewCoalescer<H = number>(options: FrameScheduler<H> & {
  send(view: ViewMessage): void;
  initial?: Omit<ViewMessage, 'generation'>;
}) {
  let view = options.initial ?? { scale: 1, surfaces: [], overscan: 0 };
  let generation = 0;
  let pending: { handle: H } | undefined;
  let disposed = false;
  return {
    update(partial: Partial<Omit<ViewMessage, 'generation'>>) {
      if (disposed) return;
      view = { ...view, ...partial };
      if (!pending) pending = { handle: options.requestFrame(() => {
        pending = undefined;
        if (!disposed) options.send({ ...view, generation: ++generation });
      }) };
    },
    dispose() {
      disposed = true;
      if (pending) options.cancelFrame(pending.handle);
      pending = undefined;
    },
  };
}
