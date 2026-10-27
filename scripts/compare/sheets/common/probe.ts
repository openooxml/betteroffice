export const params = new URLSearchParams(location.search);
export const workbookName = params.get('wb') ?? 'w100k';
export const errors: string[] = [];
addEventListener('error', (event) => errors.push(String(event.message).slice(0, 300)));
addEventListener('unhandledrejection', (event) => errors.push(String(event.reason).slice(0, 300)));

export const longtasks: [number, number][] = [];
try {
  new PerformanceObserver((list) => {
    for (const e of list.getEntries()) longtasks.push([Math.round(performance.timeOrigin + e.startTime), Math.round(e.duration)]);
  }).observe({ type: 'longtask', buffered: true });
} catch {}

let rafTimes: number[] | null = null;
function tick(now: number) {
  if (rafTimes) rafTimes.push(performance.timeOrigin + now);
  requestAnimationFrame(tick);
}
requestAnimationFrame(tick);

export interface Rect { x: number; y: number; width: number; height: number }
export interface Cell { row: number; col: number }
export interface ArmProbe {
  /** The painted cell area (headers included); used for content and blank-frame crops. */
  gridRect(): Rect | null;
  /** Active cell, zero-based. */
  active(): Promise<Cell | null>;
  /** Raw value at an A1 address through the arm's own API. */
  read(a1: string): Promise<unknown>;
  /** Writes A1 through the arm's API; resolves once the write is accepted. */
  setA1(value: number): Promise<unknown>;
  /** Scrolls so `row` is at the top; zero-based. */
  scrollToRow(row: number): Promise<unknown>;
  /** '' while no in-cell editor is open, its text while open, null when the arm can't tell. */
  editorText(): string | null;
  /** Makes the named sheet active; resolves true once asked. */
  activateSheet?(name: string): Promise<boolean>;
}

export interface Cmp extends ArmProbe {
  arm: string;
  wb: string;
  t0Epoch: number;
  readyEpoch?: number;
  ready: Promise<unknown>;
  longtasks: [number, number][];
  errors: string[];
  info: Record<string, unknown>;
  rafStart(): void;
  rafStop(): number[];
}

export const rectOf = (element: Element): Rect => {
  const r = element.getBoundingClientRect();
  return { x: r.x, y: r.y, width: r.width, height: r.height };
};

export function largestCanvas(root: ParentNode = document) {
  let best: HTMLCanvasElement | null = null;
  for (const c of root.querySelectorAll('canvas')) if (!best || c.clientWidth * c.clientHeight > best.clientWidth * best.clientHeight) best = c;
  return best;
}

export const nextFrame = () => new Promise((done) => requestAnimationFrame(() => requestAnimationFrame(done)));

export function expose(arm: string, t0: number, probe: ArmProbe, ready: Promise<unknown>, info: Record<string, unknown> = {}) {
  const cmp: Cmp = {
    ...probe, arm, wb: workbookName, t0Epoch: performance.timeOrigin + t0, ready, longtasks, errors, info,
    rafStart() { rafTimes = []; },
    rafStop() { const out = rafTimes ?? []; rafTimes = null; return out; },
  };
  ready.then(() => { cmp.readyEpoch = performance.timeOrigin + performance.now(); }, (error) => errors.push(`ready: ${String(error)}`));
  Object.assign(window, { cmp });
  return cmp;
}
