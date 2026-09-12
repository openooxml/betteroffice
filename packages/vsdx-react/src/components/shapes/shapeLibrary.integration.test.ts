import { beforeAll, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { initWasm, openDiagram } from '@betteroffice/vsdx';
import type { PagePrimitive } from '@betteroffice/vsdx';
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

for (const shape of standardShapes) {
  test(`${shape.id} keeps its geometry through collaboration and save`, () => {
    const diagram = openDiagram(fixture, { clientId: 501 });
    const peer = openDiagram(fixture, { clientId: 502 });
    try {
      const receipt = diagram.addShape('page:1', shape.draft(2, 3, 4, 5));
      const added = diagram.snapshot().pages[0].shapes.find((item) => item.id === receipt.shapeId)!;
      expect(added.cells.filter((cell) => cell.locator.section === 'Geometry').every((cell) => Boolean(cell.rowType))).toBe(true);
      const frame = diagram.layoutPage(0);
      expect(frame.primitives.some((primitive) => primitive.kind === 'shape' && Boolean(primitive.fill || primitive.stroke))).toBe(true);
      const geometry = paths(frame.primitives);
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
