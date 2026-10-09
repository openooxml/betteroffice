import { createRoot } from 'react-dom/client';
import * as React from 'react';
import JSZip from 'jszip';
import { XlsxEditor, type XlsxEditorApi, type XlsxWorkerViewerApi } from '@betteroffice/xlsx-react';
import sampleUrl from '../../../packages/xlsx/test-fixtures/sample.xlsx?url';
import chartsUrl from '../../../packages/xlsx/test-fixtures/charts.xlsx?url';
import { compareCanvasPixels } from './xlsx-canvas-parity';
import type { ViewerArm, WorkerViewerProbe } from './xlsx-worker-viewer-probe';

interface CanvasPaint {
  serial: number;
  depth: number;
  finished: boolean;
  scale: number;
  scrollLeft: number;
  scrollTop: number;
  texts: Set<string>;
  targetBox: { left: number; top: number; right: number; bottom: number } | null;
}

function mainCanvas(arm: ViewerArm) {
  return document.querySelector<HTMLCanvasElement>(`[data-arm="${arm}"] [data-testid="xlsx-scroll"] canvas`);
}

function scrollElement(arm: ViewerArm) {
  return document.querySelector<HTMLElement>(`[data-arm="${arm}"] [data-testid="xlsx-scroll"]`)!;
}

function sheetTabs(arm: ViewerArm) {
  return Array.from(document.querySelectorAll<HTMLButtonElement>(
    `[data-arm="${arm}"] [data-testid="xlsx-sheet-tabs"] [role="tab"]`
  ));
}

function installPaintProbe() {
  const paints = new WeakMap<HTMLCanvasElement, CanvasPaint>();
  let serial = 0;
  const state = (canvas: HTMLCanvasElement) => {
    const existing = paints.get(canvas);
    if (existing) return existing;
    if (!canvas.isConnected || !canvas.closest('[data-arm]')) return;
    const paint: CanvasPaint = {
      serial: 0, depth: 0, finished: false, scale: 0,
      scrollLeft: 0, scrollTop: 0, texts: new Set(), targetBox: null,
    };
    paints.set(canvas, paint);
    return paint;
  };
  for (const dimension of ['width', 'height'] as const) {
    const descriptor = Object.getOwnPropertyDescriptor(HTMLCanvasElement.prototype, dimension)!;
    Object.defineProperty(HTMLCanvasElement.prototype, dimension, {
      ...descriptor,
      set(this: HTMLCanvasElement, value: number) {
        descriptor.set!.call(this, value);
        const paint = state(this);
        if (paint) { paint.depth = 0; paint.finished = false; paint.scale = 0; }
      },
    });
  }
  const nativeSave = CanvasRenderingContext2D.prototype.save;
  CanvasRenderingContext2D.prototype.save = function (this: CanvasRenderingContext2D) {
    nativeSave.call(this);
    const paint = state(this.canvas);
    if (!paint || paint.depth++ !== 0) return;
    const scroll = this.canvas.closest<HTMLElement>('[data-testid="xlsx-scroll"]')!;
    paint.serial = ++serial;
    paint.finished = false;
    paint.scrollLeft = scroll.scrollLeft;
    paint.scrollTop = scroll.scrollTop;
    paint.texts.clear();
    paint.targetBox = null;
  };
  const nativeSetTransform = CanvasRenderingContext2D.prototype.setTransform;
  CanvasRenderingContext2D.prototype.setTransform = function (
    this: CanvasRenderingContext2D, ...args: unknown[]
  ) {
    Reflect.apply(nativeSetTransform, this, args);
    const paint = state(this.canvas);
    if (paint?.depth === 1) paint.scale = this.getTransform().a;
  } as CanvasRenderingContext2D['setTransform'];
  const nativeFillText = CanvasRenderingContext2D.prototype.fillText;
  CanvasRenderingContext2D.prototype.fillText = function (
    this: CanvasRenderingContext2D, ...args: Parameters<typeof nativeFillText>
  ) {
    nativeFillText.apply(this, args);
    const paint = state(this.canvas);
    if (!paint || paint.depth === 0) return;
    paint.texts.add(args[0]);
    if (args[0] === 'B91/25') {
      const metrics = this.measureText(args[0]);
      const transform = this.getTransform();
      const start = new DOMPoint(args[1] - metrics.actualBoundingBoxLeft, args[2] - metrics.actualBoundingBoxAscent)
        .matrixTransform(transform);
      const end = new DOMPoint(args[1] + metrics.actualBoundingBoxRight, args[2] + metrics.actualBoundingBoxDescent)
        .matrixTransform(transform);
      paint.targetBox = { left: start.x, top: start.y, right: end.x, bottom: end.y };
    }
  };
  const nativeRestore = CanvasRenderingContext2D.prototype.restore;
  CanvasRenderingContext2D.prototype.restore = function (this: CanvasRenderingContext2D) {
    nativeRestore.call(this);
    const paint = state(this.canvas);
    if (paint && paint.depth > 0 && --paint.depth === 0) paint.finished = true;
  };
  return { get: (canvas: HTMLCanvasElement | null) => canvas ? paints.get(canvas) : undefined };
}

async function until<T>(read: () => T | null | undefined, label: string): Promise<T> {
  const deadline = performance.now() + 120_000;
  for (;;) {
    const error = document.querySelector('[data-testid="xlsx-error"], [data-testid="xlsx-render-error"]');
    if (error) throw new Error(error.textContent ?? 'Editor error');
    const value = read();
    if (value != null) return value;
    if (performance.now() > deadline) throw new Error(`Timed out waiting for ${label}`);
    await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
  }
}

function syntheticSheet(prefix: string, frozen: boolean) {
  const rows = Array.from({ length: 120 }, (_, row) => {
    const cells = Array.from({ length: 30 }, (_, col) => {
      if (frozen && ((row === 0 && col >= 3 && col <= 4) ||
        (row >= 24 && row <= 25 && col >= 5 && col <= 7 && (row !== 24 || col !== 5)))) return '';
      const column = col < 26 ? String.fromCharCode(65 + col) : `A${String.fromCharCode(65 + col - 26)}`;
      return `<c r="${column}${row + 1}" t="inlineStr"><is><t>${prefix}${row + 1}/${col + 1}</t></is></c>`;
    }).join('');
    return `<row r="${row + 1}" ht="24" customHeight="1">${cells}</row>`;
  }).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:AD120"/>
  <sheetViews><sheetView workbookViewId="0">${frozen ?
    '<pane xSplit="2" ySplit="2" topLeftCell="C3" activePane="bottomRight" state="frozen"/>' : ''}</sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="24"/>
  <cols><col min="1" max="30" width="14" customWidth="1"/></cols>
  <sheetData>${rows}</sheetData>
  ${frozen ? '<mergeCells count="2"><mergeCell ref="C1:E1"/><mergeCell ref="F25:H26"/></mergeCells>' : ''}
</worksheet>`;
}

async function loadViewerInput(charts: boolean) {
  const response = await fetch(charts ? chartsUrl : sampleUrl);
  if (!response.ok) throw new Error(`Workbook fetch failed: ${response.status}`);
  const file = new Uint8Array(await response.arrayBuffer());
  await document.fonts.ready;
  const zip = await JSZip.loadAsync(file);
  if (charts) {
    const path = 'xl/worksheets/sheet1.xml';
    const sheet = await zip.file(path)!.async('string');
    if (!sheet.includes('</sheetData>')) throw new Error('Chart fixture has no sheet data');
    zip.file(path, sheet.replace('</sheetData>',
      '<row r="120"><c r="AD120" t="inlineStr"><is><t>Extent</t></is></c></row></sheetData>'));
  } else {
    zip.file('xl/worksheets/sheet1.xml', syntheticSheet('B', true));
    zip.file('xl/worksheets/sheet2.xml', syntheticSheet('S', false));
  }
  return zip.generateAsync({ type: 'uint8array' });
}

function deferred<T>() {
  let resolve!: (value: T) => void;
  let settled = false;
  const promise = new Promise<T>((done) => {
    resolve = (value) => { settled = true; done(value); };
  });
  return { promise, resolve, get settled() { return settled; } };
}

async function mount(root: HTMLElement, probe: WorkerViewerProbe) {
  const params = new URLSearchParams(location.search);
  const workerOnly = params.get('arm') === 'worker';
  const charts = params.get('fixture') === 'charts';
  const file = await loadViewerInput(charts);
  const paints = installPaintProbe();
  const localReady = deferred<XlsxEditorApi>();
  const workerReady = deferred<XlsxWorkerViewerApi>();
  const common = { file, readOnly: true as const, showToolbar: false };
  createRoot(root).render(<React.Fragment>
    {!workerOnly && <div data-arm="in-thread">
      <XlsxEditor {...common} onReady={localReady.resolve} />
    </div>}
    <div data-arm="worker">
      <XlsxEditor {...common} experimentalWorkerOpen onReady={workerReady.resolve} />
    </div>
  </React.Fragment>);
  await until(() => workerReady.settled && (workerOnly || localReady.settled) ? true : null,
    'editor onReady callbacks');
  const [local, worker] = await Promise.all([
    workerOnly ? Promise.resolve(null) : localReady.promise, workerReady.promise,
  ]);
  let zoom = 1;
  const arms: ViewerArm[] = local ? ['in-thread', 'worker'] : ['worker'];
  const waitForPaint = (arm: ViewerArm, sheet: number, after = 0) => until(() => {
    const canvas = mainCanvas(arm);
    const paint = paints.get(canvas);
    const scroll = scrollElement(arm);
    if (!canvas || !paint || !scroll || !paint.finished || paint.depth !== 0 || paint.serial <= after) return;
    if (canvas.width !== Math.round(scroll.clientWidth * devicePixelRatio) ||
      canvas.height !== Math.round(scroll.clientHeight * devicePixelRatio) || canvas.width === 0 || canvas.height === 0) return;
    if (paint.scale !== zoom * devicePixelRatio || paint.scrollLeft !== scroll.scrollLeft ||
      paint.scrollTop !== scroll.scrollTop) return;
    if (sheetTabs(arm).findIndex((tab) => tab.getAttribute('aria-selected') === 'true') !== sheet) return;
    const marker = sheet === 0 ? /^B\d+\/\d+$/ : /^S\d+\/\d+$/;
    if (charts ? !paint.texts.has('Quarter') : !Array.from(paint.texts).some((text) => marker.test(text))) return;
    return canvas;
  }, `${arm} sheet ${sheet} completed canvas paint at zoom ${zoom}`);
  await Promise.all(arms.map((arm) => waitForPaint(arm, 0)));

  const compareCurrent = async (sheet: number) => {
    if (!local) throw new Error('Parity requires two editors');
    const [a, b] = await Promise.all(arms.map((arm) => waitForPaint(arm, sheet)));
    const left = scrollElement('in-thread');
    const right = scrollElement('worker');
    if (left.scrollLeft !== right.scrollLeft || left.scrollTop !== right.scrollTop) {
      throw new Error('Editor scroll positions differ');
    }
    return {
      ...compareCanvasPixels(a, b), sheet, zoom, dpr: devicePixelRatio,
      width: a.width, height: a.height, scrollLeft: left.scrollLeft, scrollTop: left.scrollTop,
    };
  };
  probe.compareCurrent = compareCurrent;
  probe.show = async (nextZoom, position) => {
    if (!local) throw new Error('Missing local editor');
    const sheet = sheetTabs('worker').findIndex((tab) => tab.getAttribute('aria-selected') === 'true');
    if (zoom !== nextZoom) {
      const before = arms.map((arm) => paints.get(mainCanvas(arm))?.serial ?? 0);
      zoom = nextZoom;
      const results = await Promise.all([local, worker].map((api) => api.commands.execute('zoom', { scale: zoom })));
      if (results.some((result) => !result.ok)) throw new Error('Zoom command failed');
      await Promise.all(arms.map((arm, index) => waitForPaint(arm, sheet, before[index])));
    }
    const left = position === 'origin' ? 0 : 400 * zoom;
    const top = position === 'origin' ? 0 : 480 * zoom;
    const before = arms.map((arm) => {
      const scroll = scrollElement(arm);
      const after = scroll.scrollLeft === left && scroll.scrollTop === top ? 0 : paints.get(mainCanvas(arm))?.serial ?? 0;
      scroll.scrollLeft = left;
      scroll.scrollTop = top;
      if (scroll.scrollLeft !== left || scroll.scrollTop !== top) throw new Error('Fixture does not support requested scroll');
      return after;
    });
    await Promise.all(arms.map((arm, index) => waitForPaint(arm, sheet, before[index])));
    return compareCurrent(sheet);
  };
  probe.contract = async () => {
    const scroll = scrollElement('worker');
    const target = { row: 90, col: 24 };
    const selection = { anchor: target, focus: target };
    const handleIsNull = worker.handle === null;
    const syncSaveIsNull = worker.save() === null;
    const syncSelection = worker.selectCells(0, selection);
    const scrollBefore = { left: scroll.scrollLeft, top: scroll.scrollTop };
    const before = paints.get(mainCanvas('worker'))?.serial ?? 0;
    const asyncSelection = await worker.selectCellsAsync(0, selection);
    const paint = paints.get(mainCanvas('worker'));
    const targetPaintedAtResolution = !!paint?.finished && paint.serial > before && paint.texts.has('B91/25');
    const box = paint?.targetBox;
    const canvas = mainCanvas('worker')!;
    const targetVisibleAtResolution = !!box && box.right > box.left && box.bottom > box.top &&
      box.left >= 0 && box.top >= 0 && box.right <= canvas.width && box.bottom <= canvas.height;
    const scrollAfter = { left: scroll.scrollLeft, top: scroll.scrollTop };
    const bytes = await worker.saveAsync();
    return {
      handleIsNull, syncSaveIsNull, syncSelection, asyncSelection, scrollBefore, scrollAfter,
      targetPaintedAtResolution, targetVisibleAtResolution,
      savedLength: bytes?.length ?? 0, signature: Array.from(bytes?.slice(0, 4) ?? []),
    };
  };
  return { sheetCount: sheetTabs('worker').length };
}

const root = document.getElementById('worker-viewer-root');
if (root) {
  const probe: WorkerViewerProbe = {
    ready: Promise.resolve({ sheetCount: 0 }), errors: [],
    show: async () => { throw new Error('Not ready'); },
    compareCurrent: async () => { throw new Error('Not ready'); },
    contract: async () => { throw new Error('Not ready'); },
  };
  window.__xlsxWorkerViewer = probe;
  probe.ready = mount(root, probe).catch((error: Error) => {
    probe.errors.push(error.message);
    throw error;
  });
}
