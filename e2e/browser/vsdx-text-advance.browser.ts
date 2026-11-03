import { test, expect } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { extname, resolve } from 'node:path';

const root = resolve(import.meta.dirname, '../..');
const types: Record<string, string> = { '.js': 'text/javascript', '.wasm': 'application/wasm' };

test('text painted with a registered face advances as far as the layout says', async ({ page }) => {
  await page.route('**/__vsdx/**', async (route) => {
    const path = new URL(route.request().url()).pathname.replace('/__vsdx/', '');
    const body = await readFile(resolve(root, 'packages/vsdx/dist', path));
    await route.fulfill({ body, contentType: types[extname(path)] ?? 'application/octet-stream' });
  });
  await page.route('**/__fixture/doc', async (route) =>
    route.fulfill({ body: await readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx')) }));
  await page.route('**/__fixture/font', async (route) =>
    route.fulfill({ body: await readFile(resolve(root, 'packages/fonts/assets/LiberationSans-Regular.ttf')) }));
  await page.goto('/');
  const drawn = await page.evaluate(async () => {
    const vsdx = await import(`${location.origin}/__vsdx/index.js`);
    await vsdx.initWasm();
    const doc = new Uint8Array(await (await fetch('/__fixture/doc')).arrayBuffer());
    const font = new Uint8Array(await (await fetch('/__fixture/font')).arrayBuffer());
    const faces = [false, true].map((bold) => ({ family: 'Arial', bold, bytes: font }));
    const diagram = vsdx.openDiagram(doc, { clientId: 1, fonts: faces });
    for (const bold of [false, true]) {
      document.fonts.add(await new FontFace('Arial', font.slice().buffer, { weight: bold ? '700' : '400' }).load());
    }
    const list = diagram.layoutPage(0);
    const canvas = document.createElement('canvas');
    document.body.append(canvas);
    const dpr = vsdx.sizeCanvasForPage(canvas, list, 1, 1);
    const seen: { text: string; advance: number }[] = [];
    const fillText = CanvasRenderingContext2D.prototype.fillText;
    CanvasRenderingContext2D.prototype.fillText = function (text: string, x: number, y: number) {
      seen.push({ text, advance: this.measureText(text).width * Math.hypot(this.getTransform().a, this.getTransform().b) });
      fillText.call(this, text, x, y);
    };
    await vsdx.paintPage(canvas.getContext('2d'), list, dpr, 1);
    CanvasRenderingContext2D.prototype.fillText = fillText;
    const lines = new Map<string, number>();
    const walk = (items: any[]) => {
      for (const item of items) {
        if (item.kind === 'group') walk(item.primitives);
        if (item.kind !== 'textBox') continue;
        const text = item.paragraphs.flatMap((paragraph: any) => paragraph.runs.map((run: any) => run.text)).join('');
        for (const line of item.lines) lines.set(new TextDecoder().decode(new TextEncoder().encode(text).slice(line.start, line.end)), line.width * dpr);
      }
    };
    walk(list.primitives);
    return seen.map((entry) => ({ ...entry, expected: lines.get(entry.text) ?? 0, measure: list.paintTransform.a }));
  });
  expect(drawn.length).toBeGreaterThan(0);
  for (const { text, advance, expected, measure } of drawn) {
    const layout = expected * measure;
    expect(layout, text).toBeGreaterThan(0);
    expect(Math.abs(advance - layout) / layout, `${text}: drew ${advance}, layout ${layout}`).toBeLessThan(0.1);
  }
});

test('a placeholder label is drawn at label size inside the page', async ({ page }) => {
  await page.route('**/__vsdx/**', async (route) => {
    const path = new URL(route.request().url()).pathname.replace('/__vsdx/', '');
    const body = await readFile(resolve(root, 'packages/vsdx/dist', path));
    await route.fulfill({ body, contentType: types[extname(path)] ?? 'application/octet-stream' });
  });
  await page.route('**/__fixture/doc', async (route) =>
    route.fulfill({ body: await readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/group-master-text.vsdx')) }));
  await page.goto('/');
  const drawn = await page.evaluate(async () => {
    const vsdx = await import(`${location.origin}/__vsdx/index.js`);
    await vsdx.initWasm();
    const doc = new Uint8Array(await (await fetch('/__fixture/doc')).arrayBuffer());
    const diagram = vsdx.openDiagram(doc, { clientId: 1 });
    const list = diagram.layoutPage(0);
    const reasons = new Set<string>();
    const walk = (items: any[]) => { for (const item of items) { if (item.kind === 'group') walk(item.primitives); if (item.kind === 'placeholder' && item.reason) reasons.add(item.reason); } };
    walk(list.primitives);
    const canvas = document.createElement('canvas');
    document.body.append(canvas);
    const dpr = vsdx.sizeCanvasForPage(canvas, list, 1, 1);
    const seen: { advance: number; x: number; y: number }[] = [];
    const fillText = CanvasRenderingContext2D.prototype.fillText;
    CanvasRenderingContext2D.prototype.fillText = function (text: string, x: number, y: number) {
      if (reasons.has(text)) {
        const matrix = this.getTransform();
        const at = matrix.transformPoint(new DOMPoint(x, y));
        seen.push({ advance: this.measureText(text).width * Math.hypot(matrix.a, matrix.b), x: at.x, y: at.y });
      }
      fillText.call(this, text, x, y);
    };
    await vsdx.paintPage(canvas.getContext('2d'), list, dpr, 1);
    CanvasRenderingContext2D.prototype.fillText = fillText;
    return { reasons: [...reasons], seen, width: canvas.width, height: canvas.height };
  });
  expect(drawn.reasons.length).toBeGreaterThan(0);
  expect(drawn.seen).toHaveLength(drawn.reasons.length);
  for (const label of drawn.seen) {
    expect(label.advance).toBeGreaterThan(40);
    expect(label.advance).toBeLessThan(600);
    expect(label.x).toBeGreaterThanOrEqual(0);
    expect(label.x).toBeLessThan(drawn.width);
    expect(label.y).toBeGreaterThanOrEqual(0);
    expect(label.y).toBeLessThan(drawn.height);
  }
});

test('a placeholder is a thin dashed outline, not a filled block', async ({ page }) => {
  await page.route('**/__vsdx/**', async (route) => {
    const path = new URL(route.request().url()).pathname.replace('/__vsdx/', '');
    const body = await readFile(resolve(root, 'packages/vsdx/dist', path));
    await route.fulfill({ body, contentType: types[extname(path)] ?? 'application/octet-stream' });
  });
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const vsdx = await import(`${location.origin}/__vsdx/index.js`);
    const list = { contractVersion: 7, width: 384, height: 288, printWidth: 384, printHeight: 288,
      paintTransform: { a: 96, b: 0, c: 0, d: -96, e: 0, f: 288 },
      primitives: [{ kind: 'placeholder', id: 'p', zOrder: 0, x: 1, y: 1, width: 2, height: 1, reason: '' }] };
    const canvas = document.createElement('canvas');
    document.body.append(canvas);
    const dpr = vsdx.sizeCanvasForPage(canvas, list, 1, 1);
    const ctx = canvas.getContext('2d');
    if (!ctx) throw new Error('2d canvas context unavailable');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    await vsdx.paintPage(ctx, list, dpr, 1);
    const data = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
    const gray = (x: number, y: number) => { const i = (y * canvas.width + x) * 4; return data[i + 3] === 0 ? 255 : Math.min(data[i], data[i + 1], data[i + 2]); };
    const inked = [];
    for (let y = 92; y <= 100; y++) if ([150, 153, 156].some((x) => gray(x, y) < 200)) inked.push(y);
    const row = inked[0] ?? 96;
    const marks = [];
    for (let x = 98; x < 286; x++) marks.push(gray(x, row) < 200);
    let runs = 0;
    for (let i = 1; i < marks.length; i++) if (marks[i] !== marks[i - 1]) runs++;
    return { center: gray(192, 144), inked, runs, ink: marks.filter(Boolean).length, total: marks.length };
  });
  expect(result.center).toBe(255);
  expect(result.inked.length).toBeGreaterThan(0);
  expect(result.inked.length).toBeLessThanOrEqual(2);
  expect(result.runs).toBeGreaterThanOrEqual(40);
  expect(result.ink).toBeGreaterThan(result.total / 3);
});
