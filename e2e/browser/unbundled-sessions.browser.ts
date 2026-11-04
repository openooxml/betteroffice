import { test, expect, type Page } from 'playwright/test';
import { createServer, type Server } from 'node:http';
import type { AddressInfo } from 'node:net';
import { readFile } from 'node:fs/promises';
import { extname, resolve, sep } from 'node:path';

type Format = 'xlsx' | 'pptx' | 'docx';

const root = resolve(import.meta.dirname, '../..');
const types: Record<string, string> = {
  '.js': 'text/javascript', '.mjs': 'text/javascript', '.wasm': 'application/wasm', '.json': 'application/json',
};
const packages = ['xlsx', 'pptx', 'docx', 'fonts', 'fonts-latin', 'fonts-intl', 'fonts-cjk'];
const fixtures: Record<Format, string> = {
  xlsx: 'crates/ooxml-opc/tests/fixtures/sample.xlsx',
  pptx: 'apps/demo/public/betteroffice-demo.pptx',
  docx: 'apps/demo/public/betteroffice-demo.docx',
};
const workerEntries: Record<Format, string> = { xlsx: 'dist/worker.js', pptx: 'dist/worker.js', docx: 'dist/worker.mjs' };
const extension: Record<Format, string> = { xlsx: 'js', pptx: 'js', docx: 'mjs' };
const cdn = (name: string, file: string) => `/npm/@betteroffice/${name}/${file}`;

function importMap(mapInternal: boolean): string {
  const imports: Record<string, string> = {
    '@betteroffice/fonts': cdn('fonts', 'dist/index.js'),
    '@betteroffice/fonts-latin': cdn('fonts-latin', 'dist/index.js'),
    '@betteroffice/fonts-intl': cdn('fonts-intl', 'dist/index.js'),
    '@betteroffice/fonts-cjk': cdn('fonts-cjk', 'dist/index.js'),
  };
  for (const format of Object.keys(workerEntries) as Format[]) imports[`@betteroffice/${format}/worker`] = cdn(format, workerEntries[format]);
  const scopes: Record<string, Record<string, string>> = {};
  if (mapInternal) {
    for (const format of Object.keys(workerEntries) as Format[]) {
      scopes[`/npm/@betteroffice/${format}/`] = {
        '#session-worker': cdn(format, `dist/sessionWorker.browser.${extension[format]}`),
        '#session-scope': cdn(format, `dist/sessionScope.browser.${extension[format]}`),
      };
    }
  }
  return JSON.stringify({ imports, scopes });
}

let server: Server;
let origin: string;

test.beforeAll(async () => {
  server = createServer(async (request, response) => {
    const url = new URL(request.url ?? '/', 'http://localhost');
    try {
      if (url.pathname === '/') {
        const map = importMap(url.searchParams.get('internal') === '1');
        return response.writeHead(200, { 'content-type': 'text/html' })
          .end(`<!doctype html><title>unbundled</title><script type="importmap">${map}</script>`);
      }
      const fixture = /^\/__fixture\/(xlsx|pptx|docx)$/.exec(url.pathname)?.[1] as Format | undefined;
      if (fixture) return response.writeHead(200).end(await readFile(resolve(root, fixtures[fixture])));
      const match = /^\/npm\/@betteroffice\/([a-z-]+)\/((?:dist|assets)\/.+)$/.exec(url.pathname);
      if (!match || !packages.includes(match[1])) return response.writeHead(404).end();
      const directory = resolve(root, 'packages', match[1]);
      const path = resolve(directory, match[2]);
      if (!path.startsWith(directory + sep)) return response.writeHead(404).end();
      response.writeHead(200, { 'content-type': types[extname(path)] ?? 'application/octet-stream' }).end(await readFile(path));
    } catch {
      response.writeHead(404).end();
    }
  }).listen(0, '127.0.0.1');
  await new Promise((ready) => server.once('listening', ready));
  origin = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
});

test.afterAll(() => new Promise<void>((done) => server.close(() => done())));

type Outcome = { ok: true; count: number } | { ok: false; error: string };

function open(page: Page, format: Format): Promise<Outcome> {
  return page.evaluate(async (format): Promise<Outcome> => {
    try {
      const bytes = new Uint8Array(await (await fetch(`/__fixture/${format}`)).arrayBuffer());
      const specifier = `@betteroffice/${format}/worker`;
      const module = await import(specifier);
      if (format === 'xlsx') {
        const session = await module.openWorkbookSession(bytes);
        try { return { ok: true, count: (await session.call.sheets()).length }; } finally { await session.dispose(); }
      }
      if (format === 'pptx') {
        const session = await module.openPresentationSession(bytes);
        try {
          const content = await session.call.readContent();
          return content.ok ? { ok: true, count: content.stories.length } : { ok: false, error: JSON.stringify(content) };
        } finally { await session.dispose(); }
      }
      const client = new module.DocxSourceWorkerClient();
      try {
        await client.openSource({ bytes });
        const read = await client.documentRead({ kind: 'readParagraphs', request: { view: 'accepted' } });
        return read.value.ok ? { ok: true, count: read.value.paragraphs.length } : { ok: false, error: JSON.stringify(read.value) };
      } finally { client.destroy(); }
    } catch (error) {
      return { ok: false, error: `${(error as Error)?.name}: ${(error as Error)?.message ?? String(error)}` };
    }
  }, format);
}

for (const format of ['xlsx', 'pptx', 'docx'] as const) {
  test(`${format} worker session opens from raw dist with only a package import map`, async ({ page }) => {
    await page.goto(`${origin}/`);
    const outcome = await open(page, format);
    expect(outcome).toMatchObject({ ok: true });
    expect((outcome as { count: number }).count).toBeGreaterThan(0);
  });

  test(`${format} worker session opens from raw dist with the internal imports also mapped`, async ({ page }) => {
    await page.goto(`${origin}/?internal=1`);
    const outcome = await open(page, format);
    expect(outcome).toMatchObject({ ok: true });
    expect((outcome as { count: number }).count).toBeGreaterThan(0);
  });
}
