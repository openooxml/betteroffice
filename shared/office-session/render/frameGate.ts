import { intersects, isEmpty, union } from './rects';
import type { ConsumedMessage, DeviceRect } from './types';

/** @internal Damage and optional changed geometry for one surface. */
export type SurfaceWork = { rects: DeviceRect[]; full: boolean; geometryRevision?: number };

/** @internal One credited frame snapshot with a gate-assigned epoch. */
export type FrameWork = { epoch: number; baseEpoch: number | undefined; surfaces: Map<string, SurfaceWork> };

/** @internal Session-local pixel credit; receipts and resync bypass this gate. */
export function createFrameGate() {
  let pending = new Map<string, SurfaceWork>();
  const known = new Set<string>();
  let outstanding: FrameWork | undefined;
  let presentedEpoch: number | undefined;
  let epoch = 0;
  function surface(id: string): SurfaceWork {
    known.add(id);
    let work = pending.get(id);
    if (!work) { work = { rects: [], full: false }; pending.set(id, work); }
    return work;
  }
  function full(id: string): void { const work = surface(id); work.full = true; work.rects = []; }
  return {
    markDamage(id: string, rect: DeviceRect | 'full') {
      if (rect === 'full') { full(id); return; }
      if (isEmpty(rect)) return;
      const work = surface(id);
      if (work.full) return;
      let merged = { ...rect };
      for (let index = 0; index < work.rects.length;) {
        if (intersects(merged, work.rects[index]!)) {
          merged = union(merged, work.rects.splice(index, 1)[0]!);
          index = 0;
        } else index += 1;
      }
      work.rects.push(merged);
    },
    markGeometry(id: string, revision: number) {
      const work = surface(id);
      work.geometryRevision = Math.max(work.geometryRevision ?? revision, revision);
    },
    canPost() { return outstanding === undefined; },
    take(): FrameWork | undefined {
      if (outstanding || pending.size === 0) return;
      if (presentedEpoch === undefined) {
        for (const work of pending.values()) { work.full = true; work.rects = []; }
      }
      outstanding = { epoch: ++epoch, baseEpoch: presentedEpoch, surfaces: pending };
      pending = new Map();
      return outstanding;
    },
    onConsumed(message: ConsumedMessage) {
      if (!outstanding || message.epoch !== outstanding.epoch) return;
      if (message.status === 'presented') presentedEpoch = outstanding.epoch;
      else {
        for (const [id, previous] of outstanding.surfaces) {
          full(id);
          if (previous.geometryRevision !== undefined) {
            const work = surface(id);
            work.geometryRevision = Math.max(work.geometryRevision ?? previous.geometryRevision,
              previous.geometryRevision);
          }
        }
      }
      outstanding = undefined;
    },
    onResync(id?: string) { if (id !== undefined) full(id); else for (const key of known) full(key); },
  };
}
