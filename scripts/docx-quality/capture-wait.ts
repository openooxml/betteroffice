export const LAYOUT_WAIT_BASE_MS = 120_000;
export const LAYOUT_WAIT_PER_PAGE_MS = 2_000;
export const LAYOUT_WAIT_MAX_MS = 480_000;

export function layoutWaitMs(referencePages?: number) {
  const pages = Number.isInteger(referencePages) && referencePages! > 0 ? referencePages! : 0;
  return Math.min(LAYOUT_WAIT_MAX_MS, LAYOUT_WAIT_BASE_MS + pages * LAYOUT_WAIT_PER_PAGE_MS);
}

export function hasInk(pixels: ArrayLike<number>) {
  for (let index = 3; index < pixels.length; index += 4) if (pixels[index] !== 0) return true;
  return false;
}

export function isUnpaintedCanvas(width: number, height: number, inked: boolean) {
  return width === 300 && height === 150 && !inked;
}

export interface CanvasFrame {
  width: number;
  height: number;
  value: string;
  painted?: boolean;
}

export interface SettleCanvasOptions {
  read: () => CanvasFrame | null;
  matches: (width: number, height: number) => boolean;
  sleep: (ms: number) => Promise<unknown>;
  now: () => number;
  intervalMs?: number;
  stableReads?: number;
  mismatchGraceMs?: number;
}

export async function settleCanvas({
  read,
  matches,
  sleep,
  now,
  intervalMs = 200,
  stableReads = 3,
  mismatchGraceMs = 15_000,
}: SettleCanvasOptions): Promise<CanvasFrame> {
  let previous: CanvasFrame | null = null;
  let equal = 0;
  let mismatchSince: number | null = null;
  for (;;) {
    const frame = read();
    if (frame) {
      const same =
        previous !== null &&
        frame.width === previous.width &&
        frame.height === previous.height &&
        frame.value === previous.value &&
        frame.painted === previous.painted;
      equal = same ? equal + 1 : 1;
      if (!same) mismatchSince = null;
      previous = frame;
      if (equal >= stableReads) {
        if (frame.painted !== false && matches(frame.width, frame.height)) return frame;
        mismatchSince ??= now();
        if (now() - mismatchSince >= mismatchGraceMs) return frame;
      }
    }
    await sleep(intervalMs);
  }
}
