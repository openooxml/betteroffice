import { beforeAll, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { initWasm, openDiagram } from '@betteroffice/vsdx';
import type { GeometryPathCommand, PagePrimitive } from '@betteroffice/vsdx';
import { standardShapes } from './shapeLibrary';

const root = resolve(import.meta.dir, '../../../../..');
let fixture: Uint8Array;
beforeAll(async () => {
  await initWasm(await readFile(resolve(root, 'packages/vsdx/src/wasm/generated/vsdx_wasm_bg.wasm')));
  fixture = await readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/foundation.vsdx'));
});

function paths(primitives: PagePrimitive[]): unknown[] {
  return primitives.flatMap((primitive): unknown[] => primitive.kind === 'shape' ? [primitive.path] : primitive.kind === 'group' ? paths(primitive.primitives) : []);
}

type Point = { x: number; y: number };
function renderedPoints(path: GeometryPathCommand[]): Point[] {
  const result: Point[] = [];
  let current = { x: 0, y: 0 };
  let start = current;
  for (const command of path) {
    const end = { x: Number(command.x) + 0.5, y: 0.5 - Number(command.y) };
    if (command.type === 'move') { current = end; start = end; result.push(end); }
    else if (command.type === 'line') { result.push(end); current = end; }
    else if (command.type === 'close') { result.push(start); current = start; }
    else if (command.type === 'cubic') {
      const c1 = { x: Number(command.cp1x) + 0.5, y: 0.5 - Number(command.cp1y) };
      const c2 = { x: Number(command.cp2x) + 0.5, y: 0.5 - Number(command.cp2y) };
      for (let step = 1; step <= 64; step++) {
        const t = step / 64, u = 1 - t;
        result.push({ x: u ** 3 * current.x + 3 * u * u * t * c1.x + 3 * u * t * t * c2.x + t ** 3 * end.x, y: u ** 3 * current.y + 3 * u * u * t * c1.y + 3 * u * t * t * c2.y + t ** 3 * end.y });
      }
      current = end;
    } else throw new Error(`Unexpected geometry command ${command.type}`);
  }
  return result;
}
function previewPoints(path: string): Point[] {
  const tokens = path.trim().split(/\s+/);
  const result: Point[] = [];
  let current = { x: 0, y: 0 }, start = current;
  while (tokens.length) {
    const command = tokens.shift();
    const number = () => Number(tokens.shift());
    if (command === 'M' || command === 'L') {
      current = { x: number(), y: number() };
      if (command === 'M') start = current;
      result.push(current);
    } else if (command === 'Z') { current = start; result.push(start); }
    else if (command === 'A') {
      const rx = number(), ry = number(), rotation = number(), large = number(), sweep = number();
      expect(rotation).toBe(0);
      const end = { x: number(), y: number() };
      const dx = (current.x - end.x) / 2, dy = (current.y - end.y) / 2;
      const factor = (large === sweep ? -1 : 1) * Math.sqrt(Math.max(0, (rx * rx * ry * ry - rx * rx * dy * dy - ry * ry * dx * dx) / (rx * rx * dy * dy + ry * ry * dx * dx)));
      const center = { x: (current.x + end.x) / 2 + factor * rx * dy / ry, y: (current.y + end.y) / 2 - factor * ry * dx / rx };
      const angle = Math.atan2((current.y - center.y) / ry, (current.x - center.x) / rx);
      let delta = Math.atan2((end.y - center.y) / ry, (end.x - center.x) / rx) - angle;
      if (sweep && delta < 0) delta += 2 * Math.PI;
      if (!sweep && delta > 0) delta -= 2 * Math.PI;
      for (let step = 1; step <= 128; step++) {
        const at = angle + delta * step / 128;
        result.push({ x: center.x + rx * Math.cos(at), y: center.y + ry * Math.sin(at) });
      }
      current = end;
    } else throw new Error(`Unexpected preview command ${command}`);
  }
  return result;
}
function outlineDistance(points: Point[], outline: Point[]): number {
  return Math.max(...points.map((point) => Math.min(...outline.slice(1).map((end, index) => {
    const start = outline[index];
    const dx = end.x - start.x, dy = end.y - start.y;
    const length = dx * dx + dy * dy;
    const t = length ? Math.max(0, Math.min(1, ((point.x - start.x) * dx + (point.y - start.y) * dy) / length)) : 0;
    return Math.hypot(point.x - start.x - t * dx, point.y - start.y - t * dy);
  }))));
}

for (const shape of standardShapes) {
  test(`${shape.id} keeps its geometry through collaboration and save`, () => {
    const diagram = openDiagram(fixture, { clientId: 501 });
    const peer = openDiagram(fixture, { clientId: 502 });
    try {
      const receipt = diagram.addShape('page:1', shape.draft(0, 0, 1, 1));
      const added = diagram.snapshot().pages[0].shapes.find((item) => item.id === receipt.shapeId)!;
      expect(added.cells.filter((cell) => cell.locator.section === 'Geometry').every((cell) => Boolean(cell.rowType))).toBe(true);
      const frame = diagram.layoutPage(0);
      const primitive = frame.primitives.find((item) => item.kind === 'shape' && item.id === `visio/pages/page1.xml:${added.sourceId}`);
      expect(primitive?.kind).toBe('shape');
      if (primitive?.kind !== 'shape') throw new Error('Missing inserted geometry');
      const rendered = renderedPoints(primitive.path);
      const preview = previewPoints(shape.preview);
      expect(outlineDistance(rendered, preview)).toBeLessThan(0.001);
      expect(outlineDistance(preview, rendered)).toBeLessThan(0.001);
      const draftHeight = Number(added.cells.find((cell) => cell.name === 'Height')!.value);
      diagram.resizeShape('page:1', added.id, '4', '5');
      const resized = diagram.layoutPage(0).primitives.find((item) => item.kind === 'shape' && item.id === primitive.id);
      if (resized?.kind !== 'shape') throw new Error('Missing resized geometry');
      const scaled = renderedPoints(resized.path).map((point) => ({ x: (point.x - 0.5) / 4 + 0.5, y: (point.y - 0.5) / 5 * draftHeight + 0.5 }));
      expect(outlineDistance(scaled, preview)).toBeLessThan(0.001);
      expect(outlineDistance(preview, scaled)).toBeLessThan(0.001);
      expect(frame.primitives.some((primitive) => primitive.kind === 'shape' && Boolean(primitive.fill || primitive.stroke))).toBe(true);
      const geometry = paths(diagram.layoutPage(0).primitives);
      expect(geometry.length).toBeGreaterThan(0);
      expect((geometry[0] as unknown[]).length).toBeGreaterThan(1);
      peer.applyUpdate(diagram.encodeDiff(peer.encodeStateVector()));
      expect(paths(peer.layoutPage(0).primitives)).toEqual(geometry);
      const reopened = openDiagram(diagram.save(), { clientId: 503 });
      try {
        expect(paths(reopened.layoutPage(0).primitives)).toEqual(geometry);
      } finally { reopened.dispose(); }
    } finally { diagram.dispose(); peer.dispose(); }
  });
}

test('inserts the whole gallery on one page and reopens the saved package', () => {
  const diagram = openDiagram(fixture, { clientId: 601 });
  try {
    standardShapes.forEach((shape, index) => {
      diagram.addShape('page:1', shape.draft(index % 8, Math.floor(index / 8), 1, 1));
    });
    const geometry = paths(diagram.layoutPage(0).primitives);
    expect(geometry.length).toBeGreaterThanOrEqual(standardShapes.length);
    const reopened = openDiagram(diagram.save(), { clientId: 602 });
    try {
      expect(paths(reopened.layoutPage(0).primitives)).toEqual(geometry);
    } finally { reopened.dispose(); }
  } finally { diagram.dispose(); }
});
