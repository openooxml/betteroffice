import { test, expect, type Page } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { extname, resolve } from 'node:path';
import JSZip from 'jszip';

const root = resolve(import.meta.dirname, '../..');
const types: Record<string, string> = { '.js': 'text/javascript', '.mjs': 'text/javascript', '.wasm': 'application/wasm' };
const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

async function workbook(): Promise<Buffer> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `<workbook xmlns="${main}" xmlns:r="${rel}"><sheets><sheet name="Report" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${rel}/styles" Target="styles.xml"/></Relationships>`);
  zip.file('xl/styles.xml', `<styleSheet xmlns="${main}"><fonts count="1"><font><sz val="11"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf/><xf fillId="2" applyFill="1"/></cellXfs></styleSheet>`);
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${main}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Total</t></is></c></row><row r="2"><c r="B2" s="1"><v>7</v></c></row><row r="9"><c r="D9"><v>1</v></c></row></sheetData></worksheet>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

interface Export { width: number; height: number; redAtB2: boolean; inkInA1: number; error?: string }

async function exports(page: Page): Promise<Record<string, Export>> {
  const doc = await workbook();
  await page.route('**/__xlsx/**', async (route) => {
    const path = new URL(route.request().url()).pathname.replace('/__xlsx/', '');
    await route.fulfill({ body: await readFile(resolve(root, 'packages/xlsx/dist', path)), contentType: types[extname(path)] ?? 'application/octet-stream' });
  });
  await page.route('**/__fixture/doc', async (route) => route.fulfill({ body: doc }));
  await page.route('**/__page', async (route) => route.fulfill({ contentType: 'text/html', body: `<!doctype html><script type="importmap">${JSON.stringify({ imports: {
    '#session-worker': '/__xlsx/sessionWorker.browser.js', '#session-scope': '/__xlsx/sessionScope.browser.js',
  } })}</script><title>png</title>` }));
  await page.goto('/__page');
  return page.evaluate(async () => {
    const xlsx = await import(`${location.origin}/__xlsx/index.js`);
    await xlsx.initWasm();
    const bytes = new Uint8Array(await (await fetch('/__fixture/doc')).arrayBuffer());
    const handle = xlsx.openWorkbook(bytes);
    const a1 = handle.cellRect(0, 0, 0);
    const b2 = handle.cellRect(0, 1, 1);
    const scale = 2;
    const inspect = async (png: Uint8Array) => {
      const bitmap = await createImageBitmap(new Blob([png], { type: 'image/png' }));
      const canvas = new OffscreenCanvas(bitmap.width, bitmap.height);
      const ctx = canvas.getContext('2d')!;
      ctx.drawImage(bitmap, 0, 0);
      const at = (x: number, y: number) => ctx.getImageData(Math.floor(x), Math.floor(y), 1, 1).data;
      const red = at((b2.x - a1.x + b2.w / 2) * scale, (b2.y - a1.y + b2.h / 2) * scale);
      const cell = ctx.getImageData(0, 0, Math.ceil(a1.w * scale), Math.ceil(a1.h * scale)).data;
      let ink = 0;
      for (let index = 0; index < cell.length; index += 4) if (cell[index] < 128 && cell[index + 1] < 128 && cell[index + 2] < 128) ink++;
      return { width: bitmap.width, height: bitmap.height, redAtB2: red[0] > 200 && red[1] < 60 && red[2] < 60, inkInA1: ink };
    };
    const out: Record<string, unknown> = {};
    const run = async (name: string, render: () => Promise<Uint8Array>) => {
      try { out[name] = await inspect(await render()); }
      catch (error) { out[name] = { error: `${(error as { code?: string }).code ?? ''} ${(error as Error).message}` }; }
    };
    await run('thread range', () => handle.renderPng({ sheet: 0, range: 'A1:D9' }, { scale }));
    await run('thread used', () => handle.renderPng({ sheet: 'Report', range: 'used' }, { scale }));
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
  for (const name of ['thread range', 'thread used', 'worker range', 'worker used']) {
    const png = result[name];
    expect(png.error, name).toBeUndefined();
    expect(Math.abs(png.width - result.expected.width), `${name}: width ${png.width} vs ${result.expected.width}`).toBeLessThanOrEqual(1);
    expect(Math.abs(png.height - result.expected.height), `${name}: height ${png.height} vs ${result.expected.height}`).toBeLessThanOrEqual(1);
    expect(png.redAtB2, `${name}: B2 keeps its red fill`).toBe(true);
    expect(png.inkInA1, `${name}: A1 paints its text`).toBeGreaterThan(20);
  }
  expect(result['worker range']).toEqual(result['thread range']);
});
