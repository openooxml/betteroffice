import { test, expect } from 'playwright/test';
import { createServer } from 'node:http';
import type { AddressInfo } from 'node:net';
import { readFile } from 'node:fs/promises';
import { extname, resolve, sep } from 'node:path';
import { PICTURE_CASES, pictureDocument, pictureOutcomeFailures, type PictureResult } from '../docx-png-pictures';

const root = resolve(import.meta.dirname, '../..');
const types: Record<string, string> = {
  '.js': 'text/javascript', '.mjs': 'text/javascript', '.wasm': 'application/wasm', '.json': 'application/json',
};
const packages = ['docx', 'fonts', 'fonts-latin', 'fonts-intl', 'fonts-cjk'];
const cdn = (name: string) => `/npm/@betteroffice/${name}/dist/index.js`;
const imports = Object.fromEntries(['fonts', 'fonts-latin', 'fonts-intl', 'fonts-cjk'].map((name) => [`@betteroffice/${name}`, cdn(name)]));

test('a docx PNG follows the picture outcome table shared with the native engine', async ({ page }) => {
  const documents = await Promise.all(PICTURE_CASES.map(({ part }) => pictureDocument(part)));
  const server = createServer(async (request, response) => {
    const path = new URL(request.url ?? '/', 'http://x').pathname;
    try {
      if (path === '/') {
        return response.writeHead(200, { 'content-type': 'text/html' })
          .end(`<!doctype html><title>png</title><script type="importmap">${JSON.stringify({ imports })}</script>`);
      }
      const document = /^\/__case\/(\d+)$/.exec(path);
      if (document) return response.writeHead(200).end(documents[Number(document[1])]);
      const match = /^\/npm\/@betteroffice\/([a-z-]+)\/((?:dist|assets)\/.+)$/.exec(path);
      if (!match || !packages.includes(match[1])) return response.writeHead(404).end();
      const directory = resolve(root, 'packages', match[1]);
      const file = resolve(directory, match[2]);
      if (!file.startsWith(directory + sep)) return response.writeHead(404).end();
      response.writeHead(200, { 'content-type': types[extname(file)] ?? 'application/octet-stream' }).end(await readFile(file));
    } catch { response.writeHead(404).end(); }
  }).listen(0, '127.0.0.1');
  await new Promise((ready) => server.once('listening', ready));
  try {
    await page.goto(`http://127.0.0.1:${(server.address() as AddressInfo).port}/`);
    const results = await page.evaluate(async (cases) => {
      const docx = await import(`${location.origin}/npm/@betteroffice/docx/dist/yrs/index.mjs`);
      const out: Record<string, unknown> = {};
      for (const [index, { name, strictPictures }] of cases.entries()) {
        const bytes = new Uint8Array(await (await fetch(`/__case/${index}`)).arrayBuffer());
        const session = await docx.createYrsSession();
        const diagnostics: { code: string; pictures: { id: string; reason: string }[] }[] = [];
        try {
          session.openDocx(bytes, true);
          const png: Uint8Array = await session.renderPng({ page: 0 }, {
            strictPictures,
            onDiagnostic: ({ code, pictures }: { code: string; pictures: { id: string; reason: string }[] }) => diagnostics.push({ code, pictures: pictures.map(({ id, reason }) => ({ id, reason })) }),
          });
          const digest = new Uint8Array(await crypto.subtle.digest('SHA-256', new Uint8Array(png)));
          out[name] = { png: Array.from(digest, (byte) => byte.toString(16).padStart(2, '0')).join(''), diagnostics };
        } catch (error) {
          out[name] = { code: (error as { code?: string }).code, message: String((error as Error).message), diagnostics };
        } finally { session.destroy(); }
      }
      return out;
    }, PICTURE_CASES);
    expect(pictureOutcomeFailures(results as Record<string, PictureResult>)).toEqual([]);
  } finally { server.close(); }
});
