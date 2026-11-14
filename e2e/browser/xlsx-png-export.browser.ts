import { test, expect, type Page } from 'playwright/test';
import { createServer } from 'node:http';
import type { AddressInfo } from 'node:net';
import { readFile } from 'node:fs/promises';
import { extname, resolve, sep } from 'node:path';
import JSZip from 'jszip';
import { PNG_BROKEN_FACE, PNG_FONT_CASES, PNG_FONT_FACES, PNG_FONT_TARGET, pngFontOutcomeFailures, pngFontsWorkbook, type PngFontResult } from '../xlsx-png-fonts';

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
const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

async function workbook(): Promise<Buffer> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `<workbook xmlns="${main}" xmlns:r="${rel}"><sheets><sheet name="Report" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${rel}/styles" Target="styles.xml"/></Relationships>`);
  zip.file('xl/styles.xml', `<styleSheet xmlns="${main}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="11"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="3"><xf/><xf fillId="2" applyFill="1"/><xf fontId="1" applyFont="1"/></cellXfs></styleSheet>`);
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${main}"><sheetData><row r="1"><c r="A1" s="2" t="inlineStr"><is><t>Total</t></is></c></row><row r="2"><c r="B2" s="1"><v>7</v></c></row><row r="9"><c r="D9"><v>1</v></c></row></sheetData></worksheet>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

interface Export { width: number; height: number; redAtB2: boolean; inkInA1: number; ink: string; error?: string }

async function served<T>(page: Page, fixtures: Record<string, Buffer>, run: () => Promise<T>): Promise<T> {
  const server = createServer(async (request, response) => {
    const path = new URL(request.url ?? '/', 'http://x').pathname;
    try {
      if (path === '/') {
        return response.writeHead(200, { 'content-type': 'text/html' })
          .end(`<!doctype html><title>png</title><script type="importmap">${JSON.stringify({ imports })}</script>`);
      }
      const fixture = fixtures[path];
      if (fixture) return response.writeHead(200).end(fixture);
      const fonts = /^\/npm\/@betteroffice\/([a-z-]+)\/((?:dist|assets)\/.+)$/.exec(path);
      if (fonts) {
        const directory = resolve(root, 'packages', fonts[1]);
        const file = resolve(directory, fonts[2]);
        if (!fontPackages.includes(fonts[1]) || !file.startsWith(directory + sep)) return response.writeHead(404).end();
        const body = await readFile(file);
        return response.writeHead(200, { 'content-type': types[extname(file)] ?? 'application/octet-stream' })
          .end(/\.m?js$/.test(file) ? resolveFontSpecifiers(body.toString('utf8')) : body);
      }
      const file = path.replace('/__xlsx/', '');
      let body: Buffer | string = await readFile(resolve(root, 'packages/xlsx/dist', file));
      if (/\.m?js$/.test(file)) {
        body = resolveFontSpecifiers(body.toString('utf8').replace(/(["'])#session-(worker|scope)\1/g, (_, quote: string, name: string) =>
          `${quote}/__xlsx/session${name[0].toUpperCase()}${name.slice(1)}.browser.js${quote}`));
      }
      response.writeHead(200, { 'content-type': types[extname(file)] ?? 'application/octet-stream' }).end(body);
    } catch { response.writeHead(404).end(); }
  }).listen(0, '127.0.0.1');
  await new Promise((ready) => server.once('listening', ready));
  try {
    await page.goto(`http://127.0.0.1:${(server.address() as AddressInfo).port}/`);
    return await run();
  } finally { server.close(); }
}

async function exports(page: Page): Promise<Record<string, Export>> {
  const fixtures = {
    '/__fixture/doc': await workbook(),
    '/__fixture/font': await readFile(resolve(root, 'packages/fonts/assets/ComicRelief-Regular.ttf')),
  };
  return served(page, fixtures, () => evaluate(page));
}

function evaluate(page: Page): Promise<Record<string, Export>> {
  return page.evaluate(async () => {
    const xlsx = await import(`${location.origin}/__xlsx/index.js`);
    await xlsx.initWasm();
    const bytes = new Uint8Array(await (await fetch('/__fixture/doc')).arrayBuffer());
    const handle = xlsx.openWorkbook(bytes);
    const a1 = handle.cellRect(0, 0, 0);
    const b2 = handle.cellRect(0, 1, 1);
    const scale = 2;
    const inspect = async (png: Uint8Array) => {
      const bitmap = await createImageBitmap(new Blob([new Uint8Array(png)], { type: 'image/png' }));
      const canvas = new OffscreenCanvas(bitmap.width, bitmap.height);
      const ctx = canvas.getContext('2d')!;
      ctx.drawImage(bitmap, 0, 0);
      const at = (x: number, y: number) => ctx.getImageData(Math.floor(x), Math.floor(y), 1, 1).data;
      const red = at((b2.x - a1.x + b2.w / 2) * scale, (b2.y - a1.y + b2.h / 2) * scale);
      const cell = ctx.getImageData(0, 0, Math.ceil(a1.w * scale), Math.ceil(a1.h * scale)).data;
      let ink = 0;
      let mask = '';
      for (let index = 0; index < cell.length; index += 4) {
        const dark = cell[index] < 128 && cell[index + 1] < 128 && cell[index + 2] < 128;
        if (dark) ink++;
        mask += dark ? '1' : '0';
      }
      return { width: bitmap.width, height: bitmap.height, redAtB2: red[0] > 200 && red[1] < 60 && red[2] < 60, inkInA1: ink, ink: mask };
    };
    const out: Record<string, unknown> = {};
    const run = async (name: string, render: () => Promise<Uint8Array>) => {
      try { out[name] = await inspect(await render()); }
      catch (error) { out[name] = { error: `${(error as { code?: string }).code ?? ''} ${(error as Error).message}` }; }
    };
    await run('thread range', () => handle.renderPng({ sheet: 0, range: 'A1:D9' }, { scale }));
    await run('thread used', () => handle.renderPng({ sheet: 'Report', range: 'used' }, { scale }));
    const face = new Uint8Array(await (await fetch('/__fixture/font')).arrayBuffer());
    out.fonts = handle.renderPngFonts({ sheet: 0, range: 'A1:D9' });
    await run('thread range with a supplied face', () => handle.renderPng({ sheet: 0, range: 'A1:D9' }, { scale, fonts: [{ family: 'Arial', bytes: face }] }));
    const expected = { width: Math.ceil((handle.cellRect(0, 8, 3).x + handle.cellRect(0, 8, 3).w - a1.x) * scale), height: Math.ceil((handle.cellRect(0, 8, 3).y + handle.cellRect(0, 8, 3).h - a1.y) * scale) };
    handle.dispose();
    const worker = await import(`${location.origin}/__xlsx/worker.js`);
    const session = await worker.openWorkbookSession(bytes);
    try {
      await run('worker range', () => session.call.renderPng({ sheet: 0, range: 'A1:D9' }, { scale }));
      await run('worker used', () => session.call.renderPng({ sheet: 0, range: 'used' }, { scale }));
    } finally { await session.dispose(); }
    return { ...out, expected } as never;
  });
}

test('an xlsx range exports through the canvas, in the page and in a worker session', async ({ page }) => {
  const result = await exports(page) as unknown as Record<string, Export> & { expected: { width: number; height: number } };
  expect((result as unknown as { fonts: unknown }).fonts).toEqual([{ family: 'Arial', bold: false, italic: false }, { family: 'Calibri', bold: false, italic: false }]);
  for (const name of ['thread range', 'thread used', 'worker range', 'worker used', 'thread range with a supplied face']) {
    const png = result[name];
    expect(png.error, name).toBeUndefined();
    expect(Math.abs(png.width - result.expected.width), `${name}: width ${png.width} vs ${result.expected.width}`).toBeLessThanOrEqual(1);
    expect(Math.abs(png.height - result.expected.height), `${name}: height ${png.height} vs ${result.expected.height}`).toBeLessThanOrEqual(1);
    expect(png.redAtB2, `${name}: B2 keeps its red fill`).toBe(true);
    expect(png.inkInA1, `${name}: A1 paints its text`).toBeGreaterThan(20);
  }
  expect(result['worker range']).toEqual(result['thread range']);
  expect(result['thread range with a supplied face'].ink, 'the supplied face, not the runtime font, paints A1').not.toBe(result['thread range'].ink);
});

test('an xlsx PNG follows the supplied-fonts outcome table shared with the native engine', async ({ page }) => {
  const fixtures: Record<string, Buffer> = { '/__fixture/fonts-doc': Buffer.from(await pngFontsWorkbook()) };
  for (const [name, path] of Object.entries(PNG_FONT_FACES)) fixtures[`/__fixture/face/${name}`] = await readFile(resolve(root, path));
  const results = await served(page, fixtures, () => page.evaluate(async ({ cases, target, broken }) => {
    const xlsx = await import(`${location.origin}/__xlsx/index.js`);
    await xlsx.initWasm();
    const handle = xlsx.openWorkbook(new Uint8Array(await (await fetch('/__fixture/fonts-doc')).arrayBuffer()));
    const face = async (name: string) => name === 'broken'
      ? new Uint8Array(broken)
      : new Uint8Array(await (await fetch(`/__fixture/face/${name}`)).arrayBuffer());
    const out: Record<string, { png: string } | { code?: string; message: string }> = {};
    try {
      for (const { name, fonts } of cases) {
        try {
          const options = fonts ? { fonts: await Promise.all(fonts.map(async ({ family, face: faceName }) => ({ family, bytes: await face(faceName) }))) } : {};
          const png: Uint8Array = await handle.renderPng(target, options);
          let binary = '';
          for (const byte of png) binary += String.fromCharCode(byte);
          out[name] = { png: btoa(binary) };
        } catch (error) {
          out[name] = { code: (error as { code?: string }).code, message: error instanceof Error ? error.message : String(error) };
        }
      }
    } finally { handle.dispose(); }
    return out;
  }, { cases: PNG_FONT_CASES, target: PNG_FONT_TARGET, broken: PNG_BROKEN_FACE }));
  expect(pngFontOutcomeFailures(results as Record<string, PngFontResult>)).toEqual([]);
});
