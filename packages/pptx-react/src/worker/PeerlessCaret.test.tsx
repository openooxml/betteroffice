import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, test } from 'bun:test';
import type { SlideDisplayList } from '@betteroffice/pptx';
import type { CommandRect } from '@betteroffice/pptx/worker';
import type { PeerlessViewSnapshot } from './peerlessViewStore';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render } = await import('@testing-library/react');
const { PeerlessCaret } = await import('./PeerlessCaret');

afterEach(cleanup);
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

const list = { contractVersion: 1, width: 960, height: 540, background: { kind: 'solid', color: '#0b1f4b' },
  primitives: [] } as unknown as SlideDisplayList;

function presentation(caret: CommandRect) {
  let view = { focused: true, caret, selection: { kind: 'text', anchor: 2, focus: 2 }, activeLine: null, echo: null,
    confirmed: { displayList: list }, editState: { toolbar: {} } } as unknown as PeerlessViewSnapshot;
  const listeners = new Set<() => void>();
  const publish = (next: Partial<PeerlessViewSnapshot>) => {
    view = { ...view, ...next } as PeerlessViewSnapshot;
    for (const listener of listeners) listener();
  };
  return {
    store: { getSnapshot: () => view, subscribePresentation: (listener: () => void) => {
      listeners.add(listener);
      return () => { listeners.delete(listener); };
    } },
    move: (next: CommandRect) => publish({ caret: next }),
    commit: () => publish({ confirmed: { displayList: list } as PeerlessViewSnapshot['confirmed'] }),
  };
}

function slideCanvas(shade: number, fail = false) {
  const reads: number[][] = [];
  const contexts: unknown[][] = [];
  const context = {
    clearRect: () => {},
    drawImage: (_source: unknown, x: number, y: number, width: number, height: number) => { reads.push([x, y, width, height]); },
    getImageData: (_x: number, _y: number, width: number, height: number) => {
      if (fail) throw new DOMException('tainted', 'SecurityError');
      return { data: new Uint8ClampedArray(width * height * 4).map((_, index) => index % 4 === 3 ? 255 : shade) };
    },
  };
  const scratch = { width: 0, height: 0, getContext: (...args: unknown[]) => { contexts.push(['scratch', ...args]); return context; } };
  const canvas = { width: 960 * 2 * 1.25, height: 540 * 2 * 1.25, ownerDocument: { createElement: () => scratch },
    getContext: (...args: unknown[]) => { contexts.push(['slide', ...args]); return null; } } as unknown as HTMLCanvasElement;
  return { canvas, reads, contexts };
}

const ink = (element: HTMLElement) => {
  const value = element.style.background || element.style.backgroundColor;
  return value === 'rgb(0, 0, 0)' ? '#000000' : value === 'rgb(255, 255, 255)' ? '#ffffff' : value;
};

function blinking() {
  const original = window.setInterval;
  const ticks: (() => void)[] = [];
  window.setInterval = ((tick: () => void) => { ticks.push(tick); return ticks.length; }) as unknown as typeof window.setInterval;
  return { ticks, restore: () => { window.setInterval = original; } };
}

test('the worker caret is black or white by the painted pixels, one small read per caret move, frame or zoom, none per blink', () => {
  const timers = blinking();
  try {
    const run = presentation({ x: 120, y: 100, width: 1, height: 20 });
    const painted = slideCanvas(245);
    const ref = { current: painted.canvas };
    const view = render(<PeerlessCaret store={run.store} canvas={ref} scale={1.25} />);
    const caret = view.getByTestId('pptx-caret');
    expect(ink(caret)).toBe('#000000');
    expect(painted.reads).toEqual([[300, 248, 3, 54]]);
    for (let blink = 0; blink < 6; blink += 1) act(() => timers.ticks[timers.ticks.length - 1]!());
    expect(painted.reads).toHaveLength(1);
    act(() => run.move({ x: 140, y: 100, width: 1, height: 20 }));
    act(() => run.commit());
    view.rerender(<PeerlessCaret store={run.store} canvas={ref} scale={1.5} />);
    expect(painted.reads).toEqual([[300, 248, 3, 54], [350, 248, 3, 54], [350, 248, 3, 54], [350, 248, 3, 54]]);
    expect(painted.contexts).toEqual([['scratch', '2d', { willReadFrequently: true }]]);
  } finally { timers.restore(); }
});

test('the worker caret falls back to black or white by the fill estimate when the pixels cannot be read', () => {
  const run = presentation({ x: 120, y: 100, width: 1, height: 20 });
  const tainted = slideCanvas(245, true);
  const view = render(<PeerlessCaret store={run.store} canvas={{ current: tainted.canvas }} />);
  expect(tainted.reads).toHaveLength(1);
  expect(ink(view.getByTestId('pptx-caret'))).toBe('#ffffff');
  cleanup();
  expect(ink(render(<PeerlessCaret store={run.store} />).getByTestId('pptx-caret'))).toBe('#ffffff');
});
