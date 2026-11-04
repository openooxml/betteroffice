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
