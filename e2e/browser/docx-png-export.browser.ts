import { test, expect } from 'playwright/test';
import { createServer } from 'node:http';
import type { AddressInfo } from 'node:net';
import { readFile } from 'node:fs/promises';
import { extname, resolve, sep } from 'node:path';

const root = resolve(import.meta.dirname, '../..');
const types: Record<string, string> = {
  '.js': 'text/javascript', '.mjs': 'text/javascript', '.wasm': 'application/wasm', '.json': 'application/json',
};
const packages = ['docx', 'fonts', 'fonts-latin', 'fonts-intl', 'fonts-cjk'];
const cdn = (name: string) => `/npm/@betteroffice/${name}/dist/index.js`;
const imports = Object.fromEntries(['fonts', 'fonts-latin', 'fonts-intl', 'fonts-cjk'].map((name) => [`@betteroffice/${name}`, cdn(name)]));

test('a page exports through the canvas with the session edits, and leaves the session as it was', async ({ page }) => {
  const server = createServer(async (request, response) => {
    const path = new URL(request.url ?? '/', 'http://x').pathname;
    try {
      if (path === '/') {
        return response.writeHead(200, { 'content-type': 'text/html' })
          .end(`<!doctype html><title>png</title><script type="importmap">${JSON.stringify({ imports })}</script>`);
      }
      if (path === '/__docx') return response.writeHead(200).end(await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.docx')));
      const match = /^\/npm\/@betteroffice\/([a-z-]+)\/((?:dist|assets)\/.+)$/.exec(path);
      if (!match || !packages.includes(match[1])) return response.writeHead(404).end();
      const directory = resolve(root, 'packages', match[1]);
      const file = resolve(directory, match[2]);
      if (!file.startsWith(directory + sep)) return response.writeHead(404).end();
      response.writeHead(200, { 'content-type': types[extname(file)] ?? 'application/octet-stream' }).end(await readFile(file));
    } catch { response.writeHead(404).end(); }
  }).listen(0, '127.0.0.1');
  await new Promise((ready) => server.once('listening', ready));
  const warnings: string[] = [];
  page.on('console', (message) => { if (message.type() === 'warning') warnings.push(message.text()); });
  try {
    await page.goto(`http://127.0.0.1:${(server.address() as AddressInfo).port}/`);
    const result = await page.evaluate(async () => {
      const docx = await import(`${location.origin}/npm/@betteroffice/docx/dist/yrs/index.mjs`);
      const bytes = new Uint8Array(await (await fetch('/__docx')).arrayBuffer());
      const inspect = async (png: Uint8Array) => {
        const bitmap = await createImageBitmap(new Blob([new Uint8Array(png)], { type: 'image/png' }));
        const canvas = new OffscreenCanvas(bitmap.width, bitmap.height);
        const ctx = canvas.getContext('2d')!;
        ctx.drawImage(bitmap, 0, 0);
        const data = ctx.getImageData(0, 0, bitmap.width, bitmap.height).data;
        let ink = 0;
        let mask = '';
        for (let index = 0; index < data.length; index += 4) {
          const dark = data[index] < 200 || data[index + 1] < 200 || data[index + 2] < 200;
          if (dark) ink++;
          if (index % 256 === 0) mask += dark ? '1' : '0';
        }
        return { width: bitmap.width, height: bitmap.height, ink, mask, corner: [...data.slice(0, 4)] };
      };
      const session = await docx.createYrsSession();
      try {
        session.openDocx(bytes, true);
        const version = session.version();
        const one = await inspect(await session.renderPng({ page: 0 }));
        const again = await inspect(await session.renderPng({ page: 0 }));
        const double = await inspect(await session.renderPng({ page: 0 }, { scale: 2 }));
        const unchanged = session.version() === version;
        session.insertText({ story: 'body', paraId: session.paragraphs('body')[2].paraId, offset: 0 }, 'Inserted words that move the first line of this paragraph. ');
        const edited = await inspect(await session.renderPng({ page: 0 }));
        return { one, again, double, unchanged, edited };
      } finally { session.destroy(); }
    });
    expect(result.one.ink, 'the page paints').toBeGreaterThan(1000);
    expect(warnings.filter((text) => text.includes('synthetic text metrics')), 'every run has real metrics').toEqual([]);
    expect(result.one.corner, 'white behind the page').toEqual([255, 255, 255, 255]);
    expect(result.again).toEqual(result.one);
    expect(Math.abs(result.double.width - 2 * result.one.width)).toBeLessThanOrEqual(1);
    expect(result.unchanged).toBe(true);
    expect(result.edited.mask, 'the export shows the edit').not.toEqual(result.one.mask);
  } finally { server.close(); }
});
