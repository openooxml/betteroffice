import { test, expect } from 'playwright/test';
import { createServer } from 'node:http';
import type { AddressInfo } from 'node:net';
import { readFile } from 'node:fs/promises';
import { extname, resolve, sep } from 'node:path';

const root = resolve(import.meta.dirname, '../..');
const types: Record<string, string> = { '.js': 'text/javascript', '.mjs': 'text/javascript', '.wasm': 'application/wasm', '.json': 'application/json' };
const fontPackages = ['fonts', 'fonts-latin', 'fonts-intl', 'fonts-cjk'];
const fontEntry = (name: string, entry = 'index') => `/npm/@betteroffice/${name}/dist/${entry}.js`;
const imports = Object.fromEntries(fontPackages.map((name) => [`@betteroffice/${name}`, fontEntry(name)]));
/** A worker has no import map, so the served modules name the font packages by URL, as a bundler would resolve them. */
const resolveFontSpecifiers = (code: string) => code.replace(
  /(["'])@betteroffice\/(fonts(?:-latin|-intl|-cjk)?)(?:\/(progressive|cdn))?\1/g,
  (_match, quote: string, name: string, entry?: string) => `${quote}${fontEntry(name, entry)}${quote}`,
);

test('a slide exports through the canvas with the registered faces, in the page and in a worker session', async ({ page }) => {
  const server = createServer(async (request, response) => {
    const path = new URL(request.url ?? '/', 'http://x').pathname;
    try {
      if (path === '/') {
        return response.writeHead(200, { 'content-type': 'text/html' })
          .end(`<!doctype html><title>png</title><script type="importmap">${JSON.stringify({ imports })}</script>`);
      }
      if (path === '/__deck') return response.writeHead(200).end(await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')));
      if (path === '/__font') return response.writeHead(200).end(await readFile(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')));
      const fonts = /^\/npm\/@betteroffice\/([a-z-]+)\/((?:dist|assets)\/.+)$/.exec(path);
      const directory = fonts ? resolve(root, 'packages', fonts[1]) : resolve(root, 'packages/pptx/dist');
      if (fonts && !fontPackages.includes(fonts[1])) return response.writeHead(404).end();
      const file = resolve(directory, fonts ? fonts[2] : path.replace('/__pptx/', ''));
      if (!file.startsWith(directory + sep)) return response.writeHead(404).end();
      const body = await readFile(file);
      response.writeHead(200, { 'content-type': types[extname(file)] ?? 'application/octet-stream' })
        .end(/\.m?js$/.test(file) ? resolveFontSpecifiers(body.toString('utf8')) : body);
    } catch { response.writeHead(404).end(); }
  }).listen(0, '127.0.0.1');
  await new Promise((ready) => server.once('listening', ready));
  try {
    await page.goto(`http://127.0.0.1:${(server.address() as AddressInfo).port}/`);
    const result = await page.evaluate(async () => {
      const pptx = await import(`${location.origin}/__pptx/index.js`);
      await pptx.initWasm();
      const deck = new Uint8Array(await (await fetch('/__deck')).arrayBuffer());
      const font = new Uint8Array(await (await fetch('/__font')).arrayBuffer());
      const fonts = [{ family: 'Liberation Sans', bytes: font }];
      const inspect = async (png: Uint8Array) => {
        const bitmap = await createImageBitmap(new Blob([new Uint8Array(png)], { type: 'image/png' }));
        const canvas = new OffscreenCanvas(bitmap.width, bitmap.height);
        const ctx = canvas.getContext('2d')!;
        ctx.drawImage(bitmap, 0, 0);
        const data = ctx.getImageData(0, 0, bitmap.width, bitmap.height).data;
        let colored = 0;
        let mask = '';
        for (let index = 0; index < data.length; index += 4) {
          const ink = data[index] < 200 || data[index + 1] < 200 || data[index + 2] < 200;
          if (ink) colored++;
          if (index % 64 === 0) mask += ink ? '1' : '0';
        }
        return { width: bitmap.width, height: bitmap.height, colored, mask };
      };
      const handle = pptx.openPresentation(deck, { clientId: 1, fonts });
      const layout = handle.layoutSlide(0);
      const shown = await inspect(await handle.renderPng({ slide: 0 }, { scale: 2 }));
      const byId = await inspect(await handle.renderPng({ slide: handle.snapshot().slides[0].id }, { scale: 2 }));
      handle.dispose();
      const worker = await import(`${location.origin}/__pptx/worker.js`);
      const session = await worker.openPresentationSession(deck, { clientId: 2, fonts });
      let fromWorker;
      try { fromWorker = await inspect(await session.call.renderPng({ slide: 0 }, { scale: 2 })); }
      finally { await session.dispose(); }
      return { expected: [Math.ceil(layout.width * 2), Math.ceil(layout.height * 2)], shown, byId, fromWorker };
    });
    expect([result.shown.width, result.shown.height]).toEqual(result.expected);
    expect(result.shown.colored, 'the slide paints').toBeGreaterThan(1000);
    expect(result.byId).toEqual(result.shown);
    expect(result.fromWorker).toEqual(result.shown);
  } finally { server.close(); }
});
