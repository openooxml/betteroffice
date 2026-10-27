export const params = new URLSearchParams(location.search);
export const docName = params.get('doc') ?? 'syn13';
export const errors: string[] = [];
addEventListener('error', (event) => errors.push(String(event.message).slice(0, 300)));
addEventListener('unhandledrejection', (event) => errors.push(String(event.reason).slice(0, 300)));

export interface InputEvent { type: string; key?: string; t: number }
export const inputs: InputEvent[] = [];
for (const type of ['keydown', 'mousedown', 'pointerdown']) {
  addEventListener(type, (event) => {
    inputs.push({ type, key: (event as KeyboardEvent).key, t: performance.timeOrigin + event.timeStamp });
  }, { capture: true });
}

export const longtasks: [number, number][] = [];
try {
  new PerformanceObserver((list) => { for (const e of list.getEntries()) longtasks.push([Math.round(e.startTime), Math.round(e.duration)]); })
    .observe({ type: 'longtask', buffered: true });
} catch {}

export interface Rect { x: number; y: number; width: number; height: number }
export interface ArmProbe {
  pageCount(): number;
  hasText(text: string): Promise<boolean>;
  pageRects(): Rect[];
  toMiddle(): Promise<Record<string, unknown>>;
  scrollBy(dy: number): Promise<void>;
  focus?(): void;
}

const sleep = (ms: number) => new Promise((done) => setTimeout(done, ms));

export async function fetchBytes(path: string) {
  const response = await fetch(path);
  if (!response.ok) throw new Error(`fetch ${path}: ${response.status}`);
  return new Uint8Array(await response.arrayBuffer());
}

export function rectOf(element: Element): Rect {
  const r = element.getBoundingClientRect();
  return { x: r.x, y: r.y, width: r.width, height: r.height };
}

// Page rects clipped to the scroll viewport, so chrome drawn over scrolled-away page parts is never taken for text.
export function visible(rects: Rect[]) {
  const scroller = findScroller();
  const v = scroller && scroller !== document.scrollingElement ? scroller.getBoundingClientRect() : { top: 0, bottom: innerHeight, left: 0, right: innerWidth };
  return rects.map((r) => {
    const top = Math.max(r.y, v.top), bottom = Math.min(r.y + r.height, v.bottom);
    const left = Math.max(r.x, v.left), right = Math.min(r.x + r.width, v.right);
    return { x: left, y: top, width: right - left, height: bottom - top };
  }).filter((r) => r.height > 0 && r.width > 50);
}

let cachedScroller: HTMLElement | null = null;
export function findScroller(): HTMLElement | null {
  if (cachedScroller?.isConnected && cachedScroller.scrollHeight - cachedScroller.clientHeight >= 200) return cachedScroller;
  let best: HTMLElement | null = null;
  for (const element of document.querySelectorAll<HTMLElement>('*')) {
    if (element.scrollHeight - element.clientHeight < 200) continue;
    const overflow = getComputedStyle(element).overflowY;
    if (overflow !== 'auto' && overflow !== 'scroll') continue;
    if (!best || element.clientHeight * element.clientWidth > best.clientHeight * best.clientWidth) best = element;
  }
  cachedScroller = best;
  return best ?? (document.scrollingElement as HTMLElement | null);
}

export async function scrollFraction(fraction: number) {
  const scroller = findScroller();
  if (!scroller) throw new Error('no scroll container');
  const range = scroller.scrollHeight - scroller.clientHeight;
  scroller.scrollTop = Math.round(range * fraction);
  await sleep(300);
  const settled = scroller.scrollHeight - scroller.clientHeight;
  if (Math.abs(settled - range) > 50) scroller.scrollTop = Math.round(settled * fraction);
  return { method: 'scrollTop', scroller: scroller.className.slice(0, 60) || scroller.tagName, range: settled, top: scroller.scrollTop };
}

export async function scrollByPx(dy: number) {
  const scroller = findScroller();
  if (scroller) scroller.scrollTop += dy;
  await sleep(300);
}

export async function untilStable(read: () => number, t0: number, quietMs = 3000, timeoutMs = 300_000) {
  let value = read();
  let changedAt = performance.now();
  const deadline = performance.now() + timeoutMs;
  while (performance.now() < deadline) {
    await sleep(100);
    const now = read();
    if (now !== value) { value = now; changedAt = performance.now(); }
    else if (value > 0 && performance.now() - changedAt >= quietMs) break;
  }
  return { count: value, ms: Math.round(changedAt - t0), stable: performance.now() < deadline };
}

export async function firstPaint(present: () => boolean, t0: number, timeoutMs = 300_000) {
  const deadline = performance.now() + timeoutMs;
  for (;;) {
    if (present()) return Math.round(performance.now() - t0);
    if (performance.now() > deadline) throw new Error('first page never appeared');
    await new Promise((done) => requestAnimationFrame(done));
  }
}

export interface Cmp extends ArmProbe {
  arm: string;
  doc: string;
  t0Epoch: number;
  timeOrigin: number;
  longtasks: [number, number][];
  ready: Promise<unknown>;
  layoutDone: Promise<unknown>;
  openMs?: number;
  firstPageMs?: number;
  layoutDoneMs?: number;
  layoutStable?: boolean;
  errors: string[];
  inputs: InputEvent[];
  info: Record<string, unknown>;
}

export function expose(arm: string, t0: number, probe: ArmProbe, ready: Promise<unknown>, layoutDone: Promise<unknown>, info: Record<string, unknown> = {}) {
  const cmp: Cmp = { ...probe, arm, doc: docName, t0Epoch: performance.timeOrigin + t0, timeOrigin: performance.timeOrigin, longtasks, ready, layoutDone, errors, inputs, info };
  ready.catch((error) => errors.push(`ready: ${String(error)}`));
  layoutDone.catch((error) => errors.push(`layout: ${String(error)}`));
  Object.assign(window, { cmp });
  return cmp;
}
