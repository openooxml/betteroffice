/** @internal Stable retained surface identity. */
export type SurfaceKey = { kind: 'page' | 'slide' | 'pane' | 'thumb'; id: string };

/** @internal Collision-free surface lookup key. */
export function surfaceKeyId(key: SurfaceKey): string { return `${key.kind}:${key.id}`; }

/** @internal Integer device-pixel bounds. */
export type DeviceRect = { x: number; y: number; width: number; height: number };

/** @internal Transferable bitmap contract, also satisfied by ImageBitmap. */
export interface ClosableBitmap { readonly width: number; readonly height: number; close(): void }

/** @internal Replacement pixels cropped from the authoritative backing. */
export type Patch<B extends ClosableBitmap> = { surface: SurfaceKey; rect: DeviceRect; bitmap: B; full: boolean };

/** @internal Retained text-free echo background. */
export type Plate<B extends ClosableBitmap> = Omit<Patch<B>, 'full'>;

/** @internal Content-independent frame metadata; revisions use surfaceKeyId keys. */
export type FrameHeader = {
  session: string;
  incarnation: number;
  epoch: number;
  baseEpoch?: number;
  coversThroughSeq: number;
  viewportGeneration: number;
  scale: number;
  retire: SurfaceKey[];
  progressive: boolean;
  geometryRevision: Record<string, number>;
};

/** @internal Private pixel frame with binary format-owned projections. */
export type RenderFrame<B extends ClosableBitmap> = FrameHeader & {
  patches: Patch<B>[];
  plates?: Plate<B>[];
  geometry?: ArrayBuffer;
  echo?: ArrayBuffer;
};

/** @internal Latest viewport demand, independent of pixel credit. */
export type ViewMessage = {
  generation: number;
  scale: number;
  surfaces: { key: SurfaceKey; visibleRect: DeviceRect; priority: number }[];
  overscan: number;
};

/** @internal Exactly one terminal acknowledgement per received frame. */
export type ConsumedMessage = Pick<FrameHeader, 'session' | 'incarnation' | 'epoch'> & {
  status: 'presented' | 'discarded' | 'superseded';
};

/** @internal Full replacement request, accepted independently of pixel credit. */
export type ResyncMessage = { surface?: SurfaceKey; reason: string };

/** @internal Injected main-thread animation scheduling. */
export type FrameScheduler<H> = { requestFrame(callback: () => void): H; cancelFrame(handle: H): void };

/** @internal Maximum cloned metadata size in UTF-8 bytes. */
export const HEADER_BUDGET_BYTES = 2048;
const HEADER_FIELDS = ['session', 'incarnation', 'epoch', 'baseEpoch', 'coversThroughSeq',
  'viewportGeneration', 'scale', 'retire', 'progressive', 'geometryRevision'] as const;

/** @internal Unique transferables in first-occurrence order without absent projections. */
export function transferList<B extends ClosableBitmap>(frame: RenderFrame<B>): (ArrayBuffer | B)[] {
  return [...new Set<ArrayBuffer | B>([...(frame.geometry ? [frame.geometry] : []), ...(frame.echo ? [frame.echo] : []),
    ...frame.patches.map(({ bitmap }) => bitmap), ...(frame.plates ?? []).map(({ bitmap }) => bitmap)])];
}

/** @internal UTF-8 size of cloned headers and patch/plate descriptors, excluding transferables. */
export function headerBytes(frame: FrameHeader & Partial<RenderFrame<ClosableBitmap>>): number {
  const header = Object.fromEntries(HEADER_FIELDS.map((key) => [key, frame[key]]));
  const metadata = { ...header,
    patches: frame.patches?.map(({ surface, rect, full }) => ({ surface, rect, full })),
    plates: frame.plates?.map(({ surface, rect }) => ({ surface, rect })) };
  return new TextEncoder().encode(JSON.stringify(metadata)).byteLength;
}

/** @internal Strict header guard excluding undeclared object graphs. */
export function isRenderFrameHeader(value: unknown): value is FrameHeader {
  return record(value) && Object.keys(value).every((key) => HEADER_FIELDS.some((field) => field === key)) &&
    typeof value.session === 'string' && counter(value.incarnation) && counter(value.epoch) &&
    (value.baseEpoch === undefined || counter(value.baseEpoch)) && counter(value.coversThroughSeq) &&
    counter(value.viewportGeneration) && typeof value.scale === 'number' &&
    Number.isFinite(value.scale) && value.scale > 0 && typeof value.progressive === 'boolean' &&
    Array.isArray(value.retire) && value.retire.every(surface) && record(value.geometryRevision) &&
    Object.values(value.geometryRevision).every(counter);
}

function record(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && Object.getPrototypeOf(value) === Object.prototype;
}

function counter(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;
}

function surface(value: unknown): value is SurfaceKey {
  return record(value) && Object.keys(value).every((key) => key === 'kind' || key === 'id') &&
    (value.kind === 'page' || value.kind === 'slide' || value.kind === 'pane' || value.kind === 'thumb') &&
    typeof value.id === 'string';
}
