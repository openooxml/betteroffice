import { createBitmapOwnership, type BitmapOwnership } from './ownership';
import { surfaceKeyId, type ClosableBitmap, type ConsumedMessage, type FrameScheduler,
  type RenderFrame, type ResyncMessage, type SurfaceKey } from './types';

/** @internal Validation precedes the synchronous compositor transaction. */
export type RenderPresenterOptions<B extends ClosableBitmap, H> = FrameScheduler<H> & {
  validate(frame: RenderFrame<B>, presentedSurfaces: ReadonlySet<string>, presentedEpoch: number | undefined): 'ok' | 'stale' | 'mismatch';
  present(frame: RenderFrame<B>): void;
  consumed(message: ConsumedMessage): void;
  resync(message: ResyncMessage): void;
  onStall?(): void;
  onError?(error: unknown): void;
  ownership?: BitmapOwnership<B>;
};

/** @internal Latest frame per callback; successful plates use the plate object as their holder. */
export function createRenderPresenter<B extends ClosableBitmap, H = number>(options: RenderPresenterOptions<B, H>) {
  const ownership = options.ownership ?? createBitmapOwnership<B>();
  const presentedSurfaces = new Set<string>();
  const counters = { received: 0, presented: 0, superseded: 0, discarded: 0 };
  let identity: { session: string; incarnation: number; epoch: number } | undefined;
  let latest: RenderFrame<B> | undefined;
  let pending: { handle: H } | undefined;
  let requests = 2;
  let stalled = false;
  let disposed = false;

  function bitmaps(frame: RenderFrame<B>): B[] {
    return [...frame.patches, ...(frame.plates ?? [])].map(({ bitmap }) => bitmap);
  }
  function finish(frame: RenderFrame<B>, status: ConsumedMessage['status']): void {
    for (const bitmap of bitmaps(frame)) ownership.release(bitmap, frame);
    counters[status] += 1;
    options.consumed({ session: frame.session, incarnation: frame.incarnation, epoch: frame.epoch, status });
  }
  function resync(surfaces: SurfaceKey[], reason: string): void {
    if (stalled) return;
    if (requests === 0) { stalled = true; options.onStall?.(); return; }
    requests -= 1;
    if (surfaces.length === 0) options.resync({ reason });
    else for (const surface of surfaces) options.resync({ surface, reason });
  }
  function failed(frame: RenderFrame<B>, error: unknown): void {
    for (const plate of frame.plates ?? []) ownership.release(plate.bitmap, plate);
    finish(frame, 'discarded');
    resync([], 'present-failed');
    options.onError?.(error);
  }
  function draw(): void {
    pending = undefined;
    const frame = latest;
    latest = undefined;
    if (!frame || disposed) return;
    const same = identity?.session === frame.session && identity.incarnation === frame.incarnation;
    const surfaces = same ? presentedSurfaces : new Set<string>();
    const presentedEpoch = same ? identity?.epoch : undefined;
    const mismatched = new Map<string, SurfaceKey>();
    for (const patch of frame.patches) {
      if (!patch.full && (!same || frame.baseEpoch !== presentedEpoch || !surfaces.has(surfaceKeyId(patch.surface)))) {
        mismatched.set(surfaceKeyId(patch.surface), patch.surface);
      }
    }
    let validation: 'ok' | 'stale' | 'mismatch';
    try { validation = options.validate(frame, surfaces, presentedEpoch); }
    catch (error) { failed(frame, error); return; }
    if (validation === 'stale' || (identity?.session === frame.session &&
      (frame.incarnation < identity.incarnation || (same && frame.epoch <= identity.epoch)))) {
      finish(frame, 'discarded');
      return;
    }
    if (validation === 'mismatch' || mismatched.size > 0) {
      const surfaces = mismatched.size > 0 ? [...mismatched.values()] :
        [...new Map(frame.patches.map(({ surface }) => [surfaceKeyId(surface), surface])).values()];
      finish(frame, 'discarded');
      resync(surfaces, 'base-mismatch');
      return;
    }
    try {
      for (const plate of frame.plates ?? []) ownership.retain(plate.bitmap, plate);
      options.present(frame);
    } catch (error) { failed(frame, error); return; }
    if (!same) presentedSurfaces.clear();
    for (const patch of frame.patches) presentedSurfaces.add(surfaceKeyId(patch.surface));
    for (const key of frame.retire) presentedSurfaces.delete(surfaceKeyId(key));
    identity = { session: frame.session, incarnation: frame.incarnation, epoch: frame.epoch };
    requests = 2;
    stalled = false;
    finish(frame, 'presented');
  }
  return {
    counters,
    ownership,
    receive(frame: RenderFrame<B>) {
      counters.received += 1;
      for (const bitmap of bitmaps(frame)) ownership.retain(bitmap, frame);
      if (disposed) { finish(frame, 'discarded'); return; }
      const previous = latest;
      latest = frame;
      if (previous) finish(previous, 'superseded');
      if (!pending) pending = { handle: options.requestFrame(draw) };
    },
    invalidate() { if (!disposed) { requests = 2; stalled = false; } },
    dispose() {
      disposed = true;
      if (pending) options.cancelFrame(pending.handle);
      pending = undefined;
      const frame = latest;
      latest = undefined;
      if (frame) finish(frame, 'discarded');
    },
  };
}
