import { afterEach, beforeAll, describe, expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import * as fonts from '@betteroffice/fonts';
import { initWasm, openPresentation } from '../index';
import { fontAssetResponse, fontRuntimeWithoutPacks, withBrowserFontGlobals } from '../../../fonts/src/test-assets';
import * as workerFonts from '../../../fonts/src/worker';
import { getFontCdnConfiguration, setFontCdnConfiguration } from '../../../../shared/office-session/fontCdn';
import { createFontPackResponder, createWorkerFontPacks, setWorkerFontAssets } from '../../../../shared/office-session/optionalFontPacks';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { createPresentationCommandClient } from '../session/commandClient';
import { createPresentationSessionHost } from '../session/host';
import { engineKind, PptxRenderer } from './engine';

const root = resolve(import.meta.dir, '../../../..');
let fixture: Uint8Array;

beforeAll(async () => {
  const [wasm, pptx] = await Promise.all([
    readFile(resolve(import.meta.dir, 'generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
  ]);
  await initWasm(wasm);
  fixture = new Uint8Array(pptx);
});

let restore: (() => void) | undefined;
afterEach(() => { restore?.(); restore = undefined; });

function stubCanvas() {
  class Canvas {
    getContext() {
      return new Proxy({} as Record<string, unknown>, {
        get(target, property) {
          if (property === 'createLinearGradient' || property === 'createRadialGradient' || property === 'createPattern') return () => ({ addColorStop: () => undefined });
          if (property === 'measureText') return (text: string) => ({ width: text.length });
          return property in target ? target[property as string] : () => undefined;
        },
        set(target, property, value) { target[property as string] = value; return true; },
      });
    }
    convertToBlob(request: { type: string }) { return Promise.resolve(new Blob(['png'], { type: request.type })); }
  }
  class Face { constructor(public family: string) {} load() { return Promise.resolve(this); } }
  const held: Face[] = [];
  const scope = globalThis as Record<string, unknown>;
  const saved = { OffscreenCanvas: scope.OffscreenCanvas, FontFace: scope.FontFace, fonts: scope.fonts, Image: scope.Image, createImageBitmap: scope.createImageBitmap };
  Object.assign(scope, {
    OffscreenCanvas: Canvas, FontFace: Face, Image: undefined,
    createImageBitmap: () => Promise.reject(new Error('no decoder in this test')),
    fonts: {
      add(face: Face) { held.push(face); },
      delete(face: Face) { held.splice(held.indexOf(face), 1); return true; },
      load(request: string) { return Promise.resolve(held.filter((face) => request.includes(`"${face.family}"`))); },
    },
  });
  restore = () => { Object.assign(scope, saved); };
}

describe.skipIf(engineKind === 'native')('canvas renderPng with a bundled face that cannot load', () => {
  test('rejects with font-unavailable naming the face, and exports once the face loads', async () => {
    stubCanvas();
    const failing = spyOn(fonts, 'loadBundledFontBytes').mockImplementation(() => Promise.reject(new Error('offline')));
    const handle = openPresentation(fixture, { clientId: 7301 });
    try {
      await expect(handle.renderPng({ slide: 0 })).rejects.toMatchObject({
        name: 'PngExportError', code: 'font-unavailable', message: expect.stringContaining(' could not load: offline'),
      });
      failing.mockRestore();
      expect(new TextDecoder().decode(await handle.renderPng({ slide: 0 }))).toBe('png');
    } finally {
      failing.mockRestore();
      handle.dispose();
    }
  });
});

describe.skipIf(engineKind === 'native')('canvas renderPng while edits keep asking for new faces', () => {
  test('rejects with font-unavailable after a bounded number of loading rounds', async () => {
    stubCanvas();
    const original = fonts.bundledFaceInstalled;
    const handle = openPresentation(fixture, { clientId: 7302 });
    const slideId = handle.snapshot().slides[0]!.id;
    let edits = 0;
    const loading = spyOn(fonts, 'bundledFaceInstalled').mockImplementation((...args: Parameters<typeof original>) => {
      edits += 1;
      handle.addTextBox(slideId, { name: `Edit ${edits}`, rect: { x: 100000, y: 100000, width: 3000000, height: 800000 }, text: `Edit ${edits}`, style: { fontFamily: `Edited Family ${edits}`, fontSizePt: 18 } });
      return original(...args);
    });
    const extra = { family: 'Supplied Only Here', bytes: new Uint8Array(await readFile(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'))) };
    const stray: unknown[] = [];
    const onStray = (error: unknown) => { stray.push(error); };
    process.on('unhandledRejection', onStray);
    try {
      await expect(handle.renderPng({ slide: 0 })).rejects.toMatchObject({ name: 'PngExportError', code: 'font-unavailable' });
      await expect(handle.renderPng({ slide: 0 }, { fonts: [extra] })).rejects.toMatchObject({ name: 'PngExportError', code: 'font-unavailable' });
      await new Promise((settled) => setTimeout(settled, 50));
      expect(stray).toEqual([]);
      loading.mockRestore();
      expect(new TextDecoder().decode(await handle.renderPng({ slide: 0 }))).toBe('png');
    } finally {
      process.off('unhandledRejection', onStray);
      loading.mockRestore();
      handle.dispose();
    }
  });
});

async function withUninstalledPacks(run: () => Promise<void>) {
  const runtime = fontRuntimeWithoutPacks();
  const installed = spyOn(fonts, 'bundledFaceInstalled').mockImplementation(runtime.bundledFaceInstalled);
  const load = spyOn(fonts, 'loadBundledFontBytes').mockImplementation(runtime.loadBundledFontBytes);
  const provider = spyOn(fonts, 'createFontProvider').mockImplementation(runtime.createFontProvider);
  const previous = getFontCdnConfiguration();
  setFontCdnConfiguration(undefined);
  try { await run(); }
  finally {
    installed.mockRestore();
    load.mockRestore();
    provider.mockRestore();
    setFontCdnConfiguration(previous);
    Reflect.deleteProperty(globalThis, Symbol.for('betteroffice.fonts.cdn-warned'));
  }
}

describe.skipIf(engineKind === 'native')('PPTX PNG with uninstalled font packs', () => {
  test('T12: renderCommandPng rejects a primary CDN failure without keeping the editor substitute', async () => withUninstalledPacks(async () => {
    stubCanvas();
    const handle = openPresentation(fixture, { clientId: 7307 });
    const { slideId } = handle.insertSlide(0);
    handle.addTextBox(slideId, { name: 'CDN', rect: { x: 100000, y: 100000, width: 3000000, height: 800000 },
      text: 'CDN', style: { fontFamily: 'Calibri', fontSizePt: 18 } });
    const bytes = handle.save();
    handle.dispose();
    const restoreGlobals = withBrowserFontGlobals();
    const fetch = spyOn(globalThis, 'fetch').mockImplementation((input) => String(input).includes('cdn.jsdelivr.net')
      ? Promise.resolve(new Response(null, { status: 404 })) : fontAssetResponse(input));
    const register = spyOn(PptxRenderer.prototype, 'registerFont');
    const encoded = spyOn(OffscreenCanvas.prototype, 'convertToBlob');
    const pair = createInProcessPair();
    createPresentationSessionHost(pair.host, { initWasm: async () => {} });
    let client: Awaited<ReturnType<typeof createPresentationCommandClient>> | undefined;
    try {
      client = await createPresentationCommandClient(bytes, {
        session: 'cdn-export', incarnation: 'test', fontCdn: getFontCdnConfiguration(),
      }, pair.client);
      const base = fonts.BASE_LATIN_FONTS.find(({ weight, style }) => weight === 400 && style === 'normal')!;
      const substituted = register.mock.calls.some(([family, , , data]) => family === 'Calibri' && data.byteLength === base.byteLength);
      await expect(client.renderPng!({ slide: 0 })).rejects.toMatchObject({ code: 'font-unavailable' });
      expect(substituted).toBe(false);
      expect(encoded).not.toHaveBeenCalled();
    } finally {
      await client?.dispose();
      pair.client.close();
      fetch.mockRestore();
      register.mockRestore();
      encoded.mockRestore();
      restoreGlobals();
    }
  }), 30_000);

  test('PPTX PNG rejects an implicit CDN failure before encoding', async () => withUninstalledPacks(async () => {
    stubCanvas();
    const render = spyOn(OffscreenCanvas.prototype, 'convertToBlob');
    const handle = openPresentation(fixture, { clientId: 7304 });
    const restoreGlobals = withBrowserFontGlobals();
    const fetchAsset = globalThis.fetch;
    globalThis.fetch = ((input: RequestInfo | URL) => String(input).includes('cdn.jsdelivr.net')
      ? Promise.resolve(new Response(null, { status: 404 })) : fontAssetResponse(input)) as typeof fetch;
    try {
      await expect(handle.renderPng({ slide: 0 })).rejects.toMatchObject({
        code: 'font-unavailable', message: expect.stringContaining('HTTP 404'),
      });
      expect(render).not.toHaveBeenCalled();
    } finally {
      globalThis.fetch = fetchAsset;
      restoreGlobals();
      render.mockRestore();
      handle.dispose();
    }
  }));

  test('PPTX PNG registers Base Latin in headless mode without the Latin pack', async () => withUninstalledPacks(async () => {
    stubCanvas();
    const register = spyOn(PptxRenderer.prototype, 'registerFont');
    const handle = openPresentation(fixture, { clientId: 7305 });
    try {
      const { slideId } = handle.insertSlide(0);
      handle.addTextBox(slideId, { name: 'Base Latin', rect: { x: 100000, y: 100000, width: 3000000, height: 800000 },
        text: 'Base Latin', style: { fontFamily: 'Calibri', fontSizePt: 18 } });
      await expect(handle.renderPng({ slide: 0 })).resolves.toEqual(new TextEncoder().encode('png'));
      const base = fonts.BASE_LATIN_FONTS.find(({ weight, style }) => weight === 400 && style === 'normal')!;
      expect(register.mock.calls.some(([family, , , bytes]) => family === 'Calibri' && bytes.byteLength === base.byteLength)).toBe(true);
    } finally { register.mockRestore(); handle.dispose(); }
  }));

  test('PPTX PNG aborts while the worker waits for the page font CDN', async () => withUninstalledPacks(async () => {
    stubCanvas();
    const handle = openPresentation(fixture, { clientId: 7308 });
    const restoreGlobals = withBrowserFontGlobals(false);
    let requested!: () => void;
    const started = new Promise<void>((resolve) => { requested = resolve; });
    const channel = createWorkerFontPacks(({ pack }) => { if (pack === 'base') requested(); });
    setWorkerFontAssets(channel.source);
    const controller = new AbortController();
    const reason = new Error('Export cancelled');
    const result = handle.renderPng({ slide: 0 }, { signal: controller.signal }).catch((error: unknown) => error);
    try {
      await started;
      controller.abort(reason);
      const nextTurn = new Promise((resolve) => setTimeout(() => resolve('still waiting for the page'), 0));
      expect(await Promise.race([result, nextTurn])).toBe(reason);
    } finally {
      channel.dispose();
      setWorkerFontAssets(async () => undefined);
      await result;
      restoreGlobals();
      handle.dispose();
    }
  }));

  test('PPTX PNG refreshes the worker export renderer when the page CDN changes', async () => {
    stubCanvas();
    const handle = openPresentation(fixture, { clientId: 7306 });
    const previous = getFontCdnConfiguration();
    setFontCdnConfiguration(undefined);
    const fetchAsset = globalThis.fetch;
    const requests: string[] = [];
    globalThis.fetch = ((input: RequestInfo | URL) => {
      requests.push(String(input));
      return fontAssetResponse(input);
    }) as typeof fetch;
    const restoreGlobals = withBrowserFontGlobals(false);
    let pageCdn = { baseUrl: 'https://pptx-png-first.example/', timeoutMs: 5000 };
    const baseRequests: string[] = [];
    const channel = createWorkerFontPacks((message) => {
      if (message.pack === 'base') baseRequests.push(pageCdn.baseUrl);
      page.receive(message);
    });
    const page = createFontPackResponder(channel.receive, async () => undefined, () => pageCdn);
    setWorkerFontAssets(channel.source);
    const installed = spyOn(fonts, 'bundledFaceInstalled').mockImplementation(workerFonts.bundledFaceInstalled);
    const load = spyOn(fonts, 'loadBundledFontBytes').mockImplementation(workerFonts.loadBundledFontBytes);
    const register = spyOn(PptxRenderer.prototype, 'registerFont');
    try {
      const { slideId } = handle.insertSlide(0);
      handle.addTextBox(slideId, { name: 'CDN', rect: { x: 100000, y: 100000, width: 3000000, height: 800000 },
        text: 'CDN', style: { fontFamily: 'Calibri', fontSizePt: 18 } });
      await handle.renderPng({ slide: 0 });
      expect(requests.some((url) => url.startsWith(pageCdn.baseUrl))).toBe(true);
      const count = register.mock.calls.length;
      await handle.renderPng({ slide: 0 });
      expect(register).toHaveBeenCalledTimes(count);
      pageCdn = { ...pageCdn, baseUrl: 'https://pptx-png-second.example/' };
      await handle.renderPng({ slide: 0 });
      expect(requests.some((url) => url.startsWith(pageCdn.baseUrl))).toBe(true);
      expect(register.mock.calls.length).toBeGreaterThan(count);
      expect(baseRequests).toEqual(['https://pptx-png-first.example/', 'https://pptx-png-first.example/', pageCdn.baseUrl]);
    } finally {
      channel.dispose();
      setWorkerFontAssets(async () => undefined);
      installed.mockRestore();
      load.mockRestore();
      register.mockRestore();
      restoreGlobals();
      setFontCdnConfiguration(previous);
      globalThis.fetch = fetchAsset;
      handle.dispose();
    }
  });
});
