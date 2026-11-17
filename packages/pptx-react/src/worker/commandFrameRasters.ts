import type { RasterEntry, SlideRaster } from '../../../../shared/office-session/render/pptxSlideRaster';
import type { PeerlessConfirmedFrame } from './peerlessViewStore';

export type CommandFrameRaster = SlideRaster<HTMLCanvasElement>;

export interface CommandFrameLease {
  readonly raster: CommandFrameRaster;
  readonly ready: Promise<void>;
  retain(): CommandFrameLease;
  release(): void;
}

const canvasVersions = new WeakMap<HTMLCanvasElement, number>();
export function commandFrameCanvasVersion(canvas: HTMLCanvasElement): number { return canvasVersions.get(canvas) ?? 0; }
export function invalidateCommandFrameCanvas(canvas: HTMLCanvasElement): void {
  canvasVersions.set(canvas, commandFrameCanvasVersion(canvas) + 1);
}

const rasters = new WeakMap<PeerlessConfirmedFrame, Map<number, RasterEntry<HTMLCanvasElement>>>();
export function commandFrameRaster(frame: PeerlessConfirmedFrame, scale = 1,
  dpr = typeof window === 'undefined' ? 1 : window.devicePixelRatio || 1): CommandFrameRaster | undefined {
  return rasters.get(frame)?.get(scale * dpr)?.raster;
}

const softReferences = new WeakMap<RasterEntry<HTMLCanvasElement>, number>();
export function hardReferences(entry: RasterEntry<HTMLCanvasElement>): number {
  return entry.references - (softReferences.get(entry) ?? 0);
}

export function lease(entry: RasterEntry<HTMLCanvasElement>, soft = false): CommandFrameLease {
  if (entry.cancelled) throw new Error('Slide raster is released');
  entry.references += 1;
  if (soft) softReferences.set(entry, (softReferences.get(entry) ?? 0) + 1);
  let released = false;
  return { raster: entry.raster, ready: entry.ready, retain: () => lease(entry), release() {
    if (released) return;
    released = true;
    if (soft) softReferences.set(entry, (softReferences.get(entry) ?? 1) - 1);
    if (--entry.references === 0) entry.release();
  } };
}

export function retainCommandFrameRaster(frame: PeerlessConfirmedFrame, factor: number, soft = false): CommandFrameLease | undefined {
  const entry = rasters.get(frame)?.get(factor);
  return entry && !entry.cancelled ? lease(entry, soft) : undefined;
}

export { rasters };
