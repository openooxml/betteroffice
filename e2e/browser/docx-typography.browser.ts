import { test, expect } from 'playwright/test';
import { createServer } from 'node:http';
import type { AddressInfo } from 'node:net';
import { readFile } from 'node:fs/promises';
import { extname, resolve, sep } from 'node:path';
import { assertTypographyTable, inspectTypographyTable, typographyRows } from '../fixtures/docx-typography';
import { pinnedEvaluate } from './pinned-evaluate';

const root = resolve(import.meta.dirname, '../..');
const types: Record<string, string> = {
  '.js': 'text/javascript', '.mjs': 'text/javascript', '.wasm': 'application/wasm',
};

test('Word typography matches measurement and canvas fallback capabilities', async ({ page, browserName }) => {
  const server = createServer(async (request, response) => {
    const path = new URL(request.url ?? '/', 'http://x').pathname;
    try {
      if (path === '/') return response.writeHead(200, { 'content-type': 'text/html' }).end('<!doctype html><title>Typography</title>');
      const engine = /^\/__engine\/(.+)$/.exec(path);
      const asset = /^\/packages\/(docx|fonts|fonts-intl)\/((?:dist|assets)\/.+)$/.exec(path);
      const directory = engine
        ? resolve(root, 'packages/docx/src/wasm/generated/edit')
        : asset ? resolve(root, 'packages', asset[1]) : undefined;
      const relative = engine?.[1] ?? asset?.[2];
      if (!directory || !relative) return response.writeHead(404).end();
      const file = resolve(directory, relative);
      if (!file.startsWith(directory + sep)) return response.writeHead(404).end();
      response.writeHead(200, { 'content-type': types[extname(file)] ?? 'application/octet-stream' })
        .end(await readFile(file));
    } catch { response.writeHead(404).end(); }
  }).listen(0, '127.0.0.1');
  await new Promise(ready => server.once('listening', ready));
  try {
    await page.goto(`http://127.0.0.1:${(server.address() as AddressInfo).port}/`);
    const fonts = Object.fromEntries(await Promise.all([
      ['carlito', 'packages/fonts/assets/Carlito-Regular.ttf'],
      ['arabic', 'packages/fonts/assets/NotoNaskhArabic-Regular.ttf'],
      ['devanagari', 'packages/fonts-intl/assets/NotoSansDevanagari-Regular.ttf.zst'],
    ].map(async ([name, path]) => [name, [...await readFile(resolve(root, path))]])));
    const rows = await pinnedEvaluate(page, inspectTypographyTable, {
      moduleUrl: '/__engine/docx_edit.js', rows: typographyRows, fonts,
    });
    assertTypographyTable(rows);

    const capabilities = await page.evaluate(async () => {
      const data = await (await fetch('/packages/fonts/assets/Carlito-Regular.ttf')).arrayBuffer();
      const face = await new FontFace('Carlito', data).load();
      document.fonts.add(face);
      const ctx = document.createElement('canvas').getContext('2d')!;
      ctx.font = '48px Carlito';
      const textRendering = 'textRendering' in CanvasRenderingContext2D.prototype;
      const fontKerning = 'fontKerning' in CanvasRenderingContext2D.prototype;
      const width = (text: string, rendering: CanvasTextRendering, kerning: CanvasFontKerning) => {
        ctx.textRendering = rendering;
        ctx.fontKerning = kerning;
        return ctx.measureText(text).width;
      };
      const ligatureAuto = width('fi ti', 'auto', 'none');
      const ligatureSpeed = width('fi ti', 'optimizeSpeed', 'none');
      const plainAuto = width('mnop', 'auto', 'none');
      const plainSpeed = width('mnop', 'optimizeSpeed', 'none');
      const avAuto = width('AV', 'auto', 'normal');
      const avSpeed = width('AV', 'optimizeSpeed', 'normal');
      const avSeparate = width('A', 'auto', 'none') + width('V', 'auto', 'none');
      const avUnkerned = width('AV', 'auto', 'none');
      ctx.textRendering = 'auto';
      ctx.fontKerning = 'auto';
      const { drawPrimitive } = await import(`${location.origin}/packages/docx/dist/layout/render/index.mjs`);
      const states: Array<{ rendering: string; kerning: string; width: number }> = [];
      const fill = ctx.fillText.bind(ctx);
      ctx.fillText = (text, x, y) => {
        states.push({ rendering: ctx.textRendering, kerning: ctx.fontKerning, width: ctx.measureText(text).width });
        fill(text, x, y);
      };
      for (const path of ['text', 'glyphRun', 'modernEffects']) {
        for (const kern of [false, true]) {
          const fontFeatures = kern ? ['kern'] : undefined;
          const run = path === 'glyphRun'
            ? { kind: 'glyphRun', fontId: 0, size: 48, fallbackFont: '48px Carlito',
                text: 'AV', color: '#000000', fontFeatures, glyphs: [{ id: 1, x: 0, y: 50, cluster: 0 }] }
            : { kind: 'text', text: 'AV', x: 0, baselineY: 50, width: 100, font: '48px Carlito',
                color: '#000000', fontFeatures, modernEffects: path === 'modernEffects' ? {} : undefined };
          await drawPrimitive(ctx, run);
        }
      }
      return { textRendering, fontKerning, ligatureAuto, ligatureSpeed, plainAuto, plainSpeed,
        avAuto, avSpeed, avSeparate, avUnkerned, states,
        restored: { rendering: ctx.textRendering, kerning: ctx.fontKerning } };
    });
    test.info().annotations.push({ type: `${browserName} canvas typography`, description: JSON.stringify(capabilities) });
    expect(capabilities.textRendering).toBe(true);
    expect(capabilities.fontKerning).toBe(true);
    expect(Math.abs(capabilities.ligatureSpeed - capabilities.ligatureAuto)).toBeGreaterThan(0.01);
    expect(capabilities.plainSpeed).toBeCloseTo(capabilities.plainAuto, 5);
    expect(capabilities.avAuto).toBeLessThan(capabilities.avSeparate);
    const preservesKerning = capabilities.avAuto < capabilities.avUnkerned
      && Math.abs(capabilities.avSpeed - capabilities.avAuto) < 0.001;
    expect(capabilities.states).toHaveLength(6);
    capabilities.states.forEach((state, index) => {
      const kern = index % 2 === 1;
      expect(state.kerning).toBe(kern ? 'normal' : 'none');
      expect(state.rendering).toBe(!kern || preservesKerning ? 'optimizeSpeed' : 'auto');
      expect(state.width).toBeCloseTo(kern ? capabilities.avAuto : capabilities.avUnkerned, 5);
    });
    if (preservesKerning) expect(capabilities.avSpeed).toBeLessThan(capabilities.avSeparate);
    expect(capabilities.restored).toEqual({ rendering: 'auto', kerning: 'auto' });
  } finally {
    server.close();
  }
});
