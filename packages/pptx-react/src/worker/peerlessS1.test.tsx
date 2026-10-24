import './peerlessTestCapabilities';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import type {
  PptxWorkerSession, PresentationCommandClient, PresentationCommandEvents, PresentationCommandMethods,
  PresentationCommandSelection,
} from '@betteroffice/pptx/worker';
import { createSessionHost } from '../../../../shared/office-session/host';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { createPresentationCommandClient } from '../../../pptx/src/session/commandClient';
import { registerPresentationCommandHandlers } from '../../../pptx/src/session/commandHost';
import { initWasm } from '../../../pptx/src/wasm/loader';
import { PptxEditor } from '../PptxEditor';
import type { PptxAsyncEditorApi } from './createPeerlessEditorApi';
import { peerlessPainterFactories, peerlessSessionOpener } from './usePeerlessPresentation';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, render, waitFor } = await import('@testing-library/react');
const restorers: (() => void)[] = [];
const sessions: PptxWorkerSession[] = [];
const clients: PresentationCommandClient[] = [];
const canvasCalls: { canvas: HTMLCanvasElement; kind: 'paint' | 'clear'; bounds?: number[] }[] = [];
let delayedOpen: ReturnType<typeof deferred<PresentationCommandClient>> | undefined;

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

async function deck(count = 2): Promise<Uint8Array> {
  const zip = new JSZip();
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ns = `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${rel}"`;
  const slides = Array.from({ length: count }, (_, index) => index + 1);
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${slides.map((id) => `<Override PartName="/ppt/slides/slide${id}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')}</Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation ${ns}><p:sldIdLst>${slides.map((id) => `<p:sldId id="${255 + id}" r:id="rId${id}"/>`).join('')}</p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${slides.map((id) => `<Relationship Id="rId${id}" Type="${rel}/slide" Target="slides/slide${id}.xml"/>`).join('')}</Relationships>`);
  for (const id of slides) zip.file(`ppt/slides/slide${id}.xml`, `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="6000000" cy="1200000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="2400"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>slide ${id}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  return zip.generateAsync({ type: 'uint8array' });
}

beforeEach(() => {
  delayedOpen = undefined;
  canvasCalls.length = 0;
  const worker = Object.getOwnPropertyDescriptor(globalThis, 'Worker');
  Object.defineProperty(globalThis, 'Worker', { configurable: true, value: class {} });
  restorers.push(() => {
    if (worker) Object.defineProperty(globalThis, 'Worker', worker);
    else Reflect.deleteProperty(globalThis, 'Worker');
  });
  const canvas = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(function (this: HTMLCanvasElement) {
    return new Proxy({ canvas: this, measureText: (text: string) => ({ width: text.length * 10 }),
      clearRect: (x: number, y: number, width: number, height: number) => {
        canvasCalls.push({ canvas: this, kind: 'clear', bounds: [x, y, width, height] });
      },
      drawImage: () => { canvasCalls.push({ canvas: this, kind: 'paint' }); },
    }, {
      get(target, key) { return key in target ? Reflect.get(target, key) : () => {}; },
    }) as unknown as CanvasRenderingContext2D;
  } as unknown as HTMLCanvasElement['getContext']);
  restorers.push(() => canvas.mockRestore());
  for (const [property, extent] of [['clientWidth', 1000], ['clientHeight', 580]] as const) {
    const original = Object.getOwnPropertyDescriptor(HTMLElement.prototype, property);
    Object.defineProperty(HTMLElement.prototype, property, { configurable: true,
      get(this: HTMLElement) { return this.getAttribute('role') === 'application' ? extent : 0; } });
    restorers.push(() => {
      if (original) Object.defineProperty(HTMLElement.prototype, property, original);
      else Reflect.deleteProperty(HTMLElement.prototype, property);
    });
  }
  const opener = spyOn(peerlessSessionOpener, 'open').mockImplementation((bytes, options = {}) => {
    const pair = createInProcessPair();
    const registration = registerPresentationCommandHandlers({ initWasm: async () => {
      await initWasm(await readFile(new URL('../../../pptx/src/wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
    } });
    createSessionHost<PresentationCommandMethods, PresentationCommandEvents, null>(pair.host, {
      handlers: registration.handlers, policies: registration.policies,
      context: null, onDispose: registration.dispose });
    const session = `s1-${sessions.length}`;
    const ready = readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url))
      .then((font) => createPresentationCommandClient(bytes, { session, incarnation: `${session}-worker`,
        readOnly: options.readOnly, activeSlideIndex: options.activeSlideIndex,
        fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array(font).buffer }] }, pair.client))
      .then((client) => { clients.push(client); return client; });
    void ready.catch(() => {});
    const owner: PptxWorkerSession = { ready, failure: undefined, subscribeFailure: () => () => {},
      dispose: mock(async () => { await ready.then((client) => client.dispose(), () => {}); registration.dispose(); }) };
    sessions.push(owner);
    return sessions.length === 1 && delayedOpen ? { ...owner, ready: delayedOpen.promise } : owner;
  });
  restorers.push(() => opener.mockRestore());
});

afterEach(async () => {
  cleanup();
  await Promise.all(sessions.splice(0).map((owner) => owner.dispose()));
  clients.length = 0;
  for (const restore of restorers.splice(0).reverse()) restore();
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

test('initial_bounds', async () => {
  const file = await deck();
  for (const [initialSlide, index] of [[-4, 0], [99, 1]]) {
    const ready = mock((_api: PptxAsyncEditorApi) => {});
    const errors = mock((_error: Error) => {});
    const view = render(<PptxEditor fonts={[]} file={file} initialSlide={initialSlide} onReady={ready} onError={errors} />);
    await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
    expect(errors).not.toHaveBeenCalled();
    expect(clients[clients.length - 1].state.frame?.slideIndex).toBe(index);
    view.unmount();
  }
});

test('initial_rerender', async () => {
  const file = await deck();
  const ready = mock((_api: PptxAsyncEditorApi) => {});
  const view = render(<PptxEditor fonts={[]} file={file} initialSlide={1} onReady={ready} />);
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  const client = clients[0];
  const content = (await client.read()).result;
  if (!content.ok) throw new Error(content.failure.message);
  const story = content.stories[0];
  const selection: PresentationCommandSelection = { slideId: story.slideId, shapeId: story.shapeId,
    storyId: story.storyId, anchor: 0, focus: 0, affinity: 'downstream' };
  await act(async () => {
    await client.dispatch({ kind: 'insertText', payload: { text: 'edited ', at: { source: 'confirmed', selection },
      inputTime: 1, undoGroup: 's1' } });
  });
  const api = ready.mock.calls[0][0];
  await act(async () => { view.rerender(<PptxEditor fonts={[]} file={file} initialSlide={2} onReady={ready} />); });
  expect(sessions).toHaveLength(1);
  expect(sessions[0].dispose).not.toHaveBeenCalled();
  expect(ready).toHaveBeenCalledTimes(1);
  await expect(api.readContent()).resolves.toMatchObject({ ok: true, stories: [expect.objectContaining({ text: 'edited slide 1' }), expect.anything()] });
  expect((await api.saveAsync()).byteLength).toBeGreaterThan(0);
  const replacement = await deck();
  await act(async () => { view.rerender(<PptxEditor fonts={[]} file={replacement} initialSlide={2} onReady={ready} />); });
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(2));
  expect(clients[clients.length - 1].state.frame?.slideIndex).toBe(1);
});

test('empty_ready', async () => {
  const ready = mock((_api: PptxAsyncEditorApi) => {});
  const painter = spyOn(peerlessPainterFactories, 'frame');
  restorers.push(() => painter.mockRestore());
  const view = render(<PptxEditor fonts={[]} file={await deck(0)} onReady={ready} />);
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  expect(view.getByRole('status').textContent).toContain('No slides');
  expect(painter).not.toHaveBeenCalled();
  expect(view.container.querySelector('textarea')).toBeNull();
  expect(clients[0].state.activeSlideId).toBeNull();
  expect(clients[0].state.frame).toBeNull();
  const api = ready.mock.calls[0][0];
  await expect(api.whenHydrated()).resolves.toBeUndefined();
  await expect(api.readContent()).resolves.toMatchObject({ ok: true, slides: [], stories: [] });
  expect((await api.saveAsync()).byteLength).toBeGreaterThan(0);
  expect(ready).toHaveBeenCalledTimes(1);
});

test('no_file', async () => {
  const view = render(<PptxEditor fonts={[]} />);
  await waitFor(() => expect(view.getByRole('status').textContent).toContain('Open'));
  expect(view.getByRole('status').textContent).not.toContain('Preparing');
  expect(sessions).toHaveLength(0);
});

test('replacement_to_empty', async () => {
  let raster = new Uint8Array(0);
  let context: CanvasRenderingContext2D | undefined;
  let transform = [1, 0, 0, 1, 0, 0];
  const saved: number[][] = [];
  const canvasMock = spyOn(HTMLCanvasElement.prototype, 'getContext');
  const getContext = canvasMock.getMockImplementation()!;
  canvasMock.mockImplementation(function (this: HTMLCanvasElement) {
    if (this.dataset.testid !== 'pptx-slide-canvas') return getContext.call(this, '2d');
    if (context) return context;
    const reset = () => {
      raster = new Uint8Array(this.width * this.height);
      transform = [1, 0, 0, 1, 0, 0];
      saved.length = 0;
    };
    for (const property of ['width', 'height'] as const) {
      const descriptor = Object.getOwnPropertyDescriptor(HTMLCanvasElement.prototype, property)!;
      Object.defineProperty(this, property, { configurable: true,
        get: () => descriptor.get!.call(this),
        set: (value: number) => { descriptor.set!.call(this, value); reset(); } });
    }
    reset();
    const update = (value: number, x: number, y: number, width: number, height: number, matrix = transform) => {
      const [a, b, c, d, e, f] = matrix;
      const determinant = a * d - b * c;
      if (!determinant) return;
      const rasterWidth = this.width;
      const rasterHeight = this.height;
      for (let row = 0; row < rasterHeight; row++) {
        for (let column = 0; column < rasterWidth; column++) {
          const dx = column + 0.5 - e;
          const dy = row + 0.5 - f;
          const px = (d * dx - c * dy) / determinant;
          const py = (a * dy - b * dx) / determinant;
          if (px >= Math.min(x, x + width) && px < Math.max(x, x + width)
            && py >= Math.min(y, y + height) && py < Math.max(y, y + height))
            raster[row * rasterWidth + column] = value;
        }
      }
    };
    const paint = (x: number, y: number, width: number, height: number) => {
      canvasCalls.push({ canvas: this, kind: 'paint' });
      update(1, x, y, width, height);
    };
    context = new Proxy({ canvas: this, measureText: (text: string) => ({ width: text.length * 10 }),
      save: () => { saved.push([...transform]); },
      restore: () => { transform = saved.pop() ?? transform; },
      setTransform: (...matrix: number[]) => { transform = matrix; },
      resetTransform: () => { transform = [1, 0, 0, 1, 0, 0]; },
      clearRect: (x: number, y: number, width: number, height: number) => {
        canvasCalls.push({ canvas: this, kind: 'clear', bounds: [x, y, width, height] });
        update(0, x, y, width, height);
      },
      drawImage: (source: HTMLCanvasElement, ...bounds: number[]) => {
        const [x, y, width, height] = bounds.length === 8 ? bounds.slice(4) : bounds;
        paint(x, y, width ?? source.width, height ?? source.height);
      },
      fillRect: paint,
      fillText: (text: string, x: number, y: number) => { paint(x, y - 10, text.length * 10, 10); },
      putImageData: (data: ImageData, x: number, y: number) => {
        canvasCalls.push({ canvas: this, kind: 'paint' });
        update(1, x, y, data.width, data.height, [1, 0, 0, 1, 0, 0]);
      },
    }, {
      get(target, key) { return key in target ? Reflect.get(target, key) : () => {}; },
    }) as unknown as CanvasRenderingContext2D;
    return context;
  } as typeof HTMLCanvasElement.prototype.getContext);
  const ready = mock((_api: PptxAsyncEditorApi) => {});
  const view = render(<PptxEditor fonts={[]} file={await deck()} onReady={ready} />);
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  const canvas = view.getByTestId('pptx-slide-canvas') as HTMLCanvasElement;
  expect(raster.some((pixel) => pixel !== 0)).toBe(true);
  const paints = canvasCalls.filter((call) => call.canvas === canvas && call.kind === 'paint');
  expect(paints.length).toBeGreaterThan(0);
  const lastPaint = canvasCalls.lastIndexOf(paints[paints.length - 1]);
  const bounds = [0, 0, canvas.width, canvas.height];
  const replacement = await deck(0);
  await act(async () => { view.rerender(<PptxEditor fonts={[]} file={replacement} onReady={ready} />); });
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(2));
  expect(view.getByTestId('pptx-slide-canvas')).toBe(canvas);
  expect(view.getByRole('status').textContent).toContain('No slides');
  expect(clients[clients.length - 1].state.frame).toBeNull();
  expect(raster.every((pixel) => pixel === 0)).toBe(true);
  const retired = canvasCalls.slice(lastPaint + 1).filter((call) => call.canvas === canvas);
  expect(retired).toContainEqual({ canvas, kind: 'clear', bounds });
  expect(retired.some((call) => call.kind === 'paint')).toBe(false);
});

test('replacement_open', async () => {
  const gate = deferred<PresentationCommandClient>();
  delayedOpen = gate;
  const ready = mock((_api: PptxAsyncEditorApi) => {});
  const errors = mock((_error: Error) => {});
  const painter = spyOn(peerlessPainterFactories, 'frame');
  restorers.push(() => painter.mockRestore());
  const view = render(<PptxEditor fonts={[]} file={await deck()} onReady={ready} onError={errors} />);
  await waitFor(() => expect(sessions).toHaveLength(1));
  const retired = await sessions[0].ready;
  const replacement = await deck(0);
  await act(async () => { view.rerender(<PptxEditor fonts={[]} file={replacement} onReady={ready} onError={errors} />); });
  expect(sessions[0].dispose).toHaveBeenCalledTimes(1);
  await act(async () => { gate.resolve(retired); });
  await waitFor(() => expect(ready).toHaveBeenCalledTimes(1));
  expect(errors).not.toHaveBeenCalled();
  expect(painter).not.toHaveBeenCalled();
  expect(view.container.querySelector('aside')).toBeNull();
  expect(view.container.querySelector('textarea')).toBeNull();
  expect(view.getByRole('status').textContent).toContain('No slides');
  expect((await ready.mock.calls[0][0].saveAsync()).byteLength).toBeGreaterThan(0);
});
