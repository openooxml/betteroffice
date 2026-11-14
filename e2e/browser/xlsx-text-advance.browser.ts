import { test, expect } from 'playwright/test';
import { createServer } from 'node:http';
import type { AddressInfo } from 'node:net';
import { readFile } from 'node:fs/promises';
import { extname, resolve } from 'node:path';

const root = resolve(import.meta.dirname, '../..');
const types: Record<string, string> = { '.js': 'text/javascript', '.mjs': 'text/javascript', '.wasm': 'application/wasm' };

test('cell text paints as wide as its face says, without whole-pixel advances', async ({ page }) => {
  const server = createServer(async (request, response) => {
    const path = new URL(request.url ?? '/', 'http://x').pathname;
    try {
      if (path === '/') return response.writeHead(200, { 'content-type': 'text/html' }).end('<!doctype html><title>advance</title>');
      if (path === '/__font') return response.writeHead(200).end(await readFile(resolve(root, 'crates/xlsx-raster/assets/Carlito-Regular.ttf')));
      const file = path.replace('/__xlsx/', '');
      response.writeHead(200, { 'content-type': types[extname(file)] ?? 'application/octet-stream' })
        .end(await readFile(resolve(root, 'packages/xlsx/dist', file)));
    } catch { response.writeHead(404).end(); }
  }).listen(0, '127.0.0.1');
  await new Promise((ready) => server.once('listening', ready));
  try {
    await page.goto(`http://127.0.0.1:${(server.address() as AddressInfo).port}/`);
    const drawn = await page.evaluate(async () => {
      const xlsx = await import(`${location.origin}/__xlsx/index.js`);
      const font = await (await fetch('/__font')).arrayBuffer();
      document.fonts.add(await new FontFace('Calibri', font).load());
      const texts = ['$1,504,750.00', 'Grand total', 'Quarterly Budget Report', 'Line item 10'];
      const list = {
        width: 400, height: 40 * texts.length,
        commands: texts.map((text, row) => ({ op: 'text', x: 4, y: 28 + row * 40, text, fontSize: 11, color: '#000000', fontFamily: 'Calibri' })),
      };
      const canvas = document.createElement('canvas');
      canvas.width = 400; canvas.height = list.height;
      const seen: { text: string; width: number }[] = [];
      const fillText = CanvasRenderingContext2D.prototype.fillText;
      CanvasRenderingContext2D.prototype.fillText = function (text: string, x: number, y: number) {
        seen.push({ text, width: this.measureText(text).width });
        fillText.call(this, text, x, y);
      };
      try { xlsx.paintDisplayList(canvas.getContext('2d'), list, 1); }
      finally { CanvasRenderingContext2D.prototype.fillText = fillText; }
      const exact = document.createElement('canvas').getContext('2d')! as CanvasRenderingContext2D & { textRendering: string };
      exact.font = `${11 * 96 / 72 * 64}px "Calibri", sans-serif`;
      exact.textRendering = 'geometricPrecision';
      return seen.map(({ text, width }) => ({ text, width, exact: exact.measureText(text).width / 64 }));
    });
    expect(drawn).toHaveLength(4);
    for (const { text, width, exact } of drawn) {
      expect(Math.abs(width - exact) / exact, `${text}: painted ${width}px, face ${exact}px`).toBeLessThan(0.005);
    }
  } finally { server.close(); }
});
