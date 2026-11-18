import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, spyOn, test } from 'bun:test';
import type { DisplayList, TextCmd } from '@betteroffice/xlsx';
import { defaultFontPainterModule, useDefaultFonts } from './useDefaultFonts';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');
const restorers: (() => void)[] = [];

afterEach(() => {
  cleanup();
  for (const restore of restorers.splice(0).reverse()) restore();
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

const frame = (fontFamily?: string): DisplayList => ({ width: 10, height: 10,
  commands: [{ op: 'text', x: 0, y: 0, text: 'a', fontSize: 11, color: '#000', fontFamily } as TextCmd] });

function painterModule() {
  let release!: () => void;
  const ready = new Promise<void>((resolve) => { release = resolve; });
  const painter = { adopt: mock((_list: DisplayList) => {}),
    paint: mock((_: DisplayList, paint: () => void) => { paint(); return true; }), dispose: mock(() => {}) };
  const create = mock(() => painter);
  const load = spyOn(defaultFontPainterModule, 'load').mockImplementation(async () => {
    await ready;
    return { createDefaultFontPainter: create } as unknown as Awaited<ReturnType<typeof defaultFontPainterModule.load>>;
  });
  restorers.push(() => load.mockRestore());
  return { release, painter, create, load };
}

const decoder = async () => ({ font: new ArrayBuffer(0) });

test('paints before the painter module loads, then hands it the latest frame', async () => {
  const module = painterModule();
  const { result, unmount } = renderHook(() => useDefaultFonts('owner', decoder, undefined, () => {}));
  const paint = mock(() => {});
  const text = frame('Calibri');
  expect(result.current(frame(), paint)).toBe(false);
  expect(result.current(text, paint)).toBe(false);
  expect(paint).toHaveBeenCalledTimes(2);
  await act(async () => { module.release(); await Promise.resolve(); await Promise.resolve(); });
  expect(module.painter.adopt.mock.calls.map(([list]) => list)).toEqual([text]);
  expect(result.current(text, paint)).toBe(true);
  expect(module.painter.paint).toHaveBeenCalledTimes(1);
  expect(paint).toHaveBeenCalledTimes(3);
  unmount();
  expect(module.painter.dispose).toHaveBeenCalledTimes(1);
});

test('opting out never loads the painter module', async () => {
  const module = painterModule();
  const { result } = renderHook(() => useDefaultFonts('owner', decoder, false, () => {}));
  const paint = mock(() => {});
  expect(result.current(frame('Calibri'), paint)).toBe(true);
  module.release();
  await Promise.resolve();
  expect(paint).toHaveBeenCalledTimes(1);
  expect(module.load).not.toHaveBeenCalled();
});

test('unmounting before the module arrives creates no painter', async () => {
  const module = painterModule();
  const { result, unmount } = renderHook(() => useDefaultFonts('owner', decoder, undefined, () => {}));
  result.current(frame('Calibri'), () => {});
  unmount();
  await act(async () => { module.release(); await Promise.resolve(); await Promise.resolve(); });
  expect(module.create).not.toHaveBeenCalled();
});

function ownerHook(initial: unknown) {
  return renderHook(({ owner }) => useDefaultFonts(owner, decoder, undefined, () => {}), { initialProps: { owner: initial } });
}

const settleModule = (module: ReturnType<typeof painterModule>) =>
  act(async () => { module.release(); await Promise.resolve(); await Promise.resolve(); });

test('a frame painted before the workbook owner commits passes to the owner\'s painter', async () => {
  const module = painterModule();
  const { result, rerender } = ownerHook(null);
  const text = frame('Calibri');
  await settleModule(module);
  result.current(text, () => {});
  module.painter.adopt.mockClear();
  rerender({ owner: 'workbook' });
  await act(async () => { await Promise.resolve(); await Promise.resolve(); });
  expect(module.create).toHaveBeenCalledTimes(2);
  expect(module.painter.adopt.mock.calls.map(([list]) => list)).toEqual([text]);
});

test('a frame from a replaced workbook does not pass to the next painter', async () => {
  const module = painterModule();
  const { result, rerender } = ownerHook('first');
  await settleModule(module);
  result.current(frame('Calibri'), () => {});
  module.painter.adopt.mockClear();
  rerender({ owner: 'second' });
  await act(async () => { await Promise.resolve(); await Promise.resolve(); });
  rerender({ owner: null });
  await act(async () => { await Promise.resolve(); await Promise.resolve(); });
  rerender({ owner: 'third' });
  await act(async () => { await Promise.resolve(); await Promise.resolve(); });
  expect(module.create).toHaveBeenCalledTimes(4);
  expect(module.painter.adopt).not.toHaveBeenCalled();
});
