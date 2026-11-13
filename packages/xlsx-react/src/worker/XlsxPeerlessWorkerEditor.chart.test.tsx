import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it, mock, spyOn } from 'bun:test';
import { readFile } from 'node:fs/promises';
import * as xlsx from '@betteroffice/xlsx';
import { initWasm } from '../../../xlsx/src/wasm/loader';
import { createTestWorkbookSession } from '../../../xlsx/src/session/testHelpers';
import { XlsxWorkerEditor } from './XlsxWorkerEditor';
import { editableWorkbookSessionBackend } from './useEditableSessionWorkbook';
import type { XlsxAsyncEditorApi } from './createAsyncEditorApi';
import type { RenderResult } from '@testing-library/react';
import { MOUNTED } from '../__fixtures__/mounted';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
// static import would load testing-library before happy-dom installs its globals.
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');
const restorers: (() => void)[] = [];
let frames: Map<number, FrameRequestCallback>;
let nextFrame = 0;
let sessionToDispose: xlsx.WorkbookSession | undefined;

beforeAll(async () => {
  await initWasm(await readFile(new URL('../../../xlsx/src/wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)));
});

beforeEach(() => {
  frames = new Map();
  const raf = globalThis.requestAnimationFrame;
  const cancel = globalThis.cancelAnimationFrame;
  const resize = globalThis.ResizeObserver;
  globalThis.requestAnimationFrame = (callback) => { const id = ++nextFrame; frames.set(id, callback); return id; };
  globalThis.cancelAnimationFrame = (id) => { frames.delete(id); };
  globalThis.ResizeObserver = class { observe() {} disconnect() {} } as unknown as typeof ResizeObserver;
  restorers.push(() => {
    globalThis.requestAnimationFrame = raf;
    globalThis.cancelAnimationFrame = cancel;
    globalThis.ResizeObserver = resize;
  });
  for (const [key, value] of [['clientWidth', 800], ['clientHeight', 600]] as const) {
    const original = Object.getOwnPropertyDescriptor(HTMLElement.prototype, key);
    Object.defineProperty(HTMLElement.prototype, key, { configurable: true, get: () => value });
    restorers.push(() => {
      if (original) Object.defineProperty(HTMLElement.prototype, key, original);
      else Reflect.deleteProperty(HTMLElement.prototype, key);
    });
  }
  const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue({
    save() {}, restore() {}, setTransform() {}, beginPath() {}, rect() {}, clip() {},
  } as unknown as CanvasRenderingContext2D);
  const paint = spyOn(xlsx, 'paintDisplayList').mockImplementation(() => {});
  const proposals = spyOn(xlsx, 'isProposalsAvailable').mockReturnValue(false);
  restorers.push(() => context.mockRestore(), () => paint.mockRestore(), () => proposals.mockRestore());
});

afterEach(async () => {
  await act(async () => { cleanup(); });
  await sessionToDispose?.dispose();
  sessionToDispose = undefined;
  for (const restore of restorers.reverse()) restore();
  restorers.length = 0;
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

async function flush() {
  await act(async () => {
    const callbacks = [...frames.values()];
    frames.clear();
    for (const callback of callbacks) callback(0);
  });
}

async function mountCharts(readOnly = false) {
  const bytes = new Uint8Array(await readFile(
    new URL('../../../xlsx/test-fixtures/charts.xlsx', import.meta.url)
  ));
  const session = await createTestWorkbookSession(bytes, undefined, { includeEditState: true });
  sessionToDispose = session;
  const open = spyOn(editableWorkbookSessionBackend, 'open').mockResolvedValue(session);
  restorers.push(() => open.mockRestore());
  const ready = mock((_api: XlsxAsyncEditorApi) => {});
  const view = render(<XlsxWorkerEditor backend="peerless" file={bytes} readOnly={readOnly} onReady={ready} />);
  await waitFor(async () => {
    await flush();
    expect(ready).toHaveBeenCalledTimes(1);
  }, MOUNTED);
  const frame = await session.call.frame({ x: 0, y: 0, width: 800, height: 600 }, { sheet: 0 });
  return { view, session, charts: frame.displayList.charts ?? [] };
}

async function selectChart(view: RenderResult, chart: xlsx.ChartRegion) {
  const surface = view.getByTestId('xlsx-scroll');
  const point = { clientX: chart.clip.x + chart.clip.w / 2, clientY: chart.clip.y + chart.clip.h / 2 };
  fireEvent.mouseDown(surface, point);
  await waitFor(() => expect(view.queryByTestId('xlsx-chart-selection')).not.toBeNull(), MOUNTED);
  fireEvent.mouseUp(window, point);
}

describe('peerless chart resize', () => {
  it('draws handles and submits one resize for a south-east drag', async () => {
    const { view, session, charts } = await mountCharts();
    const chart = charts.find((candidate) => candidate.resizable)!;
    await selectChart(view, chart);

    const handles = view.queryAllByTestId('xlsx-chart-resize-handle');
    expect(handles.map((handle) => handle.dataset.handle)).toEqual(
      ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']
    );

    const submit = spyOn(session, 'submit');
    const outline = view.getByTestId('xlsx-chart-selection');
    const width = parseFloat(outline.style.width);
    const se = view.container.querySelector<HTMLElement>('[data-handle="se"]')!;

    fireEvent.pointerDown(se, { button: 0, pointerId: 1, clientX: 200, clientY: 200 });
    fireEvent.pointerMove(window, { pointerId: 1, clientX: 220, clientY: 210 });
    expect(parseFloat(outline.style.width)).toBeCloseTo(width + 20);

    fireEvent.pointerUp(window, { pointerId: 1, button: 0, clientX: 220, clientY: 210 });
    await waitFor(async () => {
      await flush();
      const resizes = submit.mock.calls.filter(([command]) => command.kind === 'resizeChart');
      expect(resizes).toHaveLength(1);
    }, MOUNTED);
    const resize = submit.mock.calls.find(([command]) => command.kind === 'resizeChart')![0];
    expect(resize.payload).toMatchObject({
      args: [0, chart.id, { left: 0, top: 0, right: 20, bottom: 10 }],
    });

    await waitFor(async () => {
      await flush();
      const frame = await session.call.frame({ x: 0, y: 0, width: 800, height: 600 }, { sheet: 0 });
      const resized = (frame.displayList.charts ?? []).find((candidate) => candidate.id === chart.id);
      expect(resized!.rect.w).toBeCloseTo(chart.rect.w + 20, 1);
      expect(resized!.rect.h).toBeCloseTo(chart.rect.h + 10, 1);
    }, MOUNTED);
  });

  it('keeps the move preview and commits one move on release', async () => {
    const { view, session, charts } = await mountCharts();
    const chart = charts.find((candidate) => candidate.resizable)!;
    await selectChart(view, chart);
    const submit = spyOn(session, 'submit');
    const surface = view.getByTestId('xlsx-scroll');
    const start = { clientX: chart.clip.x + chart.clip.w / 2, clientY: chart.clip.y + chart.clip.h / 2 };

    fireEvent.mouseDown(surface, start);
    fireEvent.mouseMove(surface, { clientX: start.clientX + 30, clientY: start.clientY + 20, buttons: 1 });
    expect(view.getByTestId('xlsx-chart-selection').style.transform).toBe('translate(30px, 20px)');
    expect(view.queryAllByTestId('xlsx-chart-resize-handle')).toHaveLength(0);

    fireEvent.mouseUp(window, { clientX: start.clientX + 30, clientY: start.clientY + 20 });
    await waitFor(async () => {
      await flush();
      expect(submit.mock.calls.filter(([command]) => command.kind === 'moveChart')).toHaveLength(1);
    }, MOUNTED);
    expect(submit.mock.calls.find(([command]) => command.kind === 'moveChart')![0].payload)
      .toMatchObject({ args: [0, chart.id, 30, 20] });
  });

  it('draws no handles for a chart the engine cannot resize', async () => {
    const { view, charts } = await mountCharts();
    await selectChart(view, charts.find((candidate) => !candidate.resizable)!);
    expect(view.queryAllByTestId('xlsx-chart-resize-handle')).toHaveLength(0);
  });

  it('draws no handles when the editor is read-only', async () => {
    const { view, charts } = await mountCharts(true);
    await selectChart(view, charts.find((candidate) => candidate.resizable)!);
    expect(view.queryAllByTestId('xlsx-chart-resize-handle')).toHaveLength(0);
  });
});
