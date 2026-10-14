import { beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { initWasm, openDiagram } from '@betteroffice/vsdx';
import type { ShapeSnapshot } from '@betteroffice/vsdx';
import { standardShapeById } from './components/shapes/shapeLibrary';
import { connectionPointsForShape, connectorDraft, connectorGlue, connectorRouteFromFrame, routeConnector } from './connector';
import type { ConnectionPoint } from './connector';

const root = resolve(import.meta.dir, '../../..');
let foundation: Uint8Array;

beforeAll(async () => {
  await initWasm(await readFile(resolve(root, 'packages/vsdx/src/wasm/generated/vsdx_wasm_bg.wasm')));
  foundation = await readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/foundation.vsdx'));
});

function pinOf(shape: ShapeSnapshot, name: string): number {
  const cell = shape.cells.find((item) => item.name === name)!;
  return Number(cell.value ?? cell.formula);
}

function snapPoint(shapes: ShapeSnapshot[], shapeId: string, side: ConnectionPoint['side']): ConnectionPoint {
  return connectionPointsForShape(shapes.find((shape) => shape.id === shapeId)!).find((point) => point.side === side)!;
}

describe('connector glue through the UI draft helpers', () => {
  test('a dynamic connector follows both shapes when they move', () => {
    const diagram = openDiagram(foundation, { clientId: 7101 });
    try {
      const rectangle = standardShapeById('rectangle')!;
      const from = diagram.addShape('page:1', rectangle.draft(2, 2, 1, 1));
      const to = diagram.addShape('page:1', rectangle.draft(5, 2, 1, 1));
      const page = diagram.snapshot().pages[0];
      const fromCentre = snapPoint(page.shapes, from.shapeId, 'centre');
      const toCentre = snapPoint(page.shapes, to.shapeId, 'centre');
      const receipt = diagram.addConnector('page:1', connectorDraft(fromCentre, toCentre), connectorGlue(from.shapeId, fromCentre), connectorGlue(to.shapeId, toCentre));
      const part = diagram.snapshot().pages[0].sourcePartPath;
      const connector = diagram.snapshot().pages[0].shapes.find((shape) => shape.id === receipt.shapeId)!;
      const route = connectorRouteFromFrame(diagram.layoutPage(0), part, connector.sourceId)!;
      expect(route).toEqual(routeConnector(fromCentre, toCentre));
      diagram.moveShape('page:1', from.shapeId, '3', '3');
      const moved = connectorRouteFromFrame(diagram.layoutPage(0), part, connector.sourceId)!;
      expect(moved[0]).toEqual({ x: 3, y: 3 });
      expect(moved).not.toEqual(route);
      diagram.moveShape('page:1', to.shapeId, '7', '4');
      const movedBoth = connectorRouteFromFrame(diagram.layoutPage(0), part, connector.sourceId)!;
      expect(movedBoth[movedBoth.length - 1]).toEqual({ x: 7, y: 4 });
    } finally { diagram.dispose(); }
  });

  test('the engine route and the overlay preview agree on a diagonal connector', () => {
    const diagram = openDiagram(foundation, { clientId: 7104 });
    try {
      const rectangle = standardShapeById('rectangle')!;
      const from = diagram.addShape('page:1', rectangle.draft(2, 2, 1, 1));
      const to = diagram.addShape('page:1', rectangle.draft(5, 5, 1, 1));
      const shapes = () => diagram.snapshot().pages[0].shapes;
      const fromCentre = snapPoint(shapes(), from.shapeId, 'centre');
      const toCentre = snapPoint(shapes(), to.shapeId, 'centre');
      const receipt = diagram.addConnector('page:1', connectorDraft(fromCentre, toCentre), connectorGlue(from.shapeId, fromCentre), connectorGlue(to.shapeId, toCentre));
      const part = diagram.snapshot().pages[0].sourcePartPath;
      const connector = shapes().find((shape) => shape.id === receipt.shapeId)!;
      const route = connectorRouteFromFrame(diagram.layoutPage(0), part, connector.sourceId)!;
      expect(route).toHaveLength(3);
      expect(route).toEqual(routeConnector(fromCentre, toCentre));
    } finally { diagram.dispose(); }
  });

  test('a point-glued connector stays pinned to its outline side while the shape moves', () => {
    const diagram = openDiagram(foundation, { clientId: 7102 });
    try {
      const rectangle = standardShapeById('rectangle')!;
      const from = diagram.addShape('page:1', rectangle.draft(2, 2, 1, 1));
      const to = diagram.addShape('page:1', rectangle.draft(5, 2, 1, 1));
      const shapes = () => diagram.snapshot().pages[0].shapes;
      const north = snapPoint(shapes(), from.shapeId, 'north');
      const south = snapPoint(shapes(), to.shapeId, 'south');
      const receipt = diagram.addConnector('page:1', connectorDraft(north, south), connectorGlue(from.shapeId, north), connectorGlue(to.shapeId, south));
      const part = diagram.snapshot().pages[0].sourcePartPath;
      const connector = shapes().find((shape) => shape.id === receipt.shapeId)!;
      const before = connectorRouteFromFrame(diagram.layoutPage(0), part, connector.sourceId)!;
      expect(before[0]).toEqual({ x: north.x, y: north.y });
      diagram.moveShape('page:1', to.shapeId, '7', '4');
      const after = connectorRouteFromFrame(diagram.layoutPage(0), part, connector.sourceId)!;
      const target = shapes().find((shape) => shape.id === to.shapeId)!;
      expect(pinOf(target, 'PinX')).toBe(7);
      expect(pinOf(target, 'PinY')).toBe(4);
      const movedSouth = snapPoint(shapes(), to.shapeId, 'south');
      expect(movedSouth.y).toBeLessThan(4);
      expect(after[after.length - 1].x).toBeCloseTo(movedSouth.x, 9);
      expect(after[after.length - 1].y).toBeCloseTo(movedSouth.y, 9);
      expect(after[0]).toEqual(before[0]);
    } finally { diagram.dispose(); }
  });

  test('a created connector renders, saves, and reloads identically with a fresh shape id', async () => {
    const diagram = openDiagram(foundation, { clientId: 7103 });
    const rectangle = standardShapeById('rectangle')!;
    const from = diagram.addShape('page:1', rectangle.draft(2, 2, 1, 1));
    const to = diagram.addShape('page:1', rectangle.draft(5, 4, 1, 1));
    const shapes = diagram.snapshot().pages[0].shapes;
    const fromCentre = snapPoint(shapes, from.shapeId, 'centre');
    const toCentre = snapPoint(shapes, to.shapeId, 'centre');
    const receipt = diagram.addConnector('page:1', connectorDraft(fromCentre, toCentre), connectorGlue(from.shapeId, fromCentre), connectorGlue(to.shapeId, toCentre));
    const live = diagram.layoutPage(0);
    const part = diagram.snapshot().pages[0].sourcePartPath;
    const connector = diagram.snapshot().pages[0].shapes.find((shape) => shape.id === receipt.shapeId)!;
    expect(live.primitives.find((item) => item.id === `${part}:${connector.sourceId}`)?.kind).toBe('shape');
    const saved = diagram.save();
    diagram.dispose();

    const reopened = openDiagram(saved, { clientId: 7104 });
    try {
      expect(reopened.layoutPage(0)).toEqual(live);
      const savedConnector = reopened.snapshot().pages[0].shapes.find((shape) => shape.name === 'Dynamic connector')!;
      expect(savedConnector.id).not.toBe(receipt.shapeId);
      expect(savedConnector.cells).toEqual(expect.arrayContaining([
        expect.objectContaining({ name: 'ShapeRouteStyle', formula: '1' }),
        expect.objectContaining({ name: 'EndArrow', formula: '4' }),
      ]));
      reopened.deleteShape('page:1', savedConnector.id);
      const deleted = reopened.save();
      const archive = await JSZip.loadAsync(deleted);
      const pageXml = await archive.file('visio/pages/page1.xml')!.async('text');
      expect(pageXml).not.toMatch(new RegExp(`(?:FromSheet|ToSheet)=["']${savedConnector.sourceId}["']`));
    } finally { reopened.dispose(); }

    const [original, result] = await Promise.all([JSZip.loadAsync(foundation), JSZip.loadAsync(saved)]);
    expect(Object.keys(result.files).sort()).toEqual(Object.keys(original.files).sort());
    await Promise.all(Object.keys(original.files).filter((path) => path !== 'visio/pages/page1.xml').map(async (path) => {
      expect(await result.file(path)!.async('uint8array')).toEqual(await original.file(path)!.async('uint8array'));
    }));
  });

  test('an outline drop on a row-less imported shape still draws through dynamic glue', () => {
    const diagram = openDiagram(foundation, { clientId: 7107 });
    try {
      const plain = diagram.addShape('page:1', { name: 'Imported', cells: [
        { locator: { cellName: 'Width' }, formula: '2' },
        { locator: { cellName: 'Height' }, formula: '1' },
        { locator: { cellName: 'PinX' }, formula: '2' },
        { locator: { cellName: 'PinY' }, formula: '2' },
        { locator: { cellName: 'LocPinX' }, formula: 'Width*0.5' },
        { locator: { cellName: 'LocPinY' }, formula: 'Height*0.5' },
      ] });
      const imported = diagram.snapshot().pages[0].shapes.find((shape) => shape.id === plain.shapeId)!;
      const east = connectionPointsForShape(imported).find((point) => point.side === 'east')!;
      expect(east.toCell).toBeUndefined();
      const rectangle = standardShapeById('rectangle')!;
      const to = diagram.addShape('page:1', rectangle.draft(5, 2, 1, 1));
      const shapes = diagram.snapshot().pages[0].shapes;
      const target = shapes.find((shape) => shape.id === to.shapeId)!;
      const south = connectionPointsForShape(target).find((point) => point.side === 'south')!;
      expect(south.toCell).toBe('Connections.X3');
      const receipt = diagram.addConnector('page:1', connectorDraft(east, south), connectorGlue(imported.id, east), connectorGlue(to.shapeId, south));
      const part = diagram.snapshot().pages[0].sourcePartPath;
      const connector = diagram.snapshot().pages[0].shapes.find((shape) => shape.id === receipt.shapeId)!;
      const live = diagram.layoutPage(0);
      expect(live.primitives.find((item) => item.id === `${part}:${connector.sourceId}`)?.kind).toBe('shape');
      const route = connectorRouteFromFrame(live, part, connector.sourceId)!;
      expect(route).not.toBeNull();
      expect(route.length).toBeGreaterThanOrEqual(2);
      expect(route[route.length - 1]).toEqual({ x: south.x, y: south.y });
    } finally { diagram.dispose(); }
  });

  test('deleting a glued shape leaves no dangling glue behind', async () => {
    const diagram = openDiagram(foundation, { clientId: 7105 });
    const rectangle = standardShapeById('rectangle')!;
    const from = diagram.addShape('page:1', rectangle.draft(2, 2, 1, 1));
    const to = diagram.addShape('page:1', rectangle.draft(5, 2, 1, 1));
    const shapes = diagram.snapshot().pages[0].shapes;
    const fromCentre = snapPoint(shapes, from.shapeId, 'centre');
    const toCentre = snapPoint(shapes, to.shapeId, 'centre');
    diagram.addConnector('page:1', connectorDraft(fromCentre, toCentre), connectorGlue(from.shapeId, fromCentre), connectorGlue(to.shapeId, toCentre));
    const beforeDelete = await (await JSZip.loadAsync(diagram.save())).file('visio/pages/page1.xml')!.async('text');
    diagram.deleteShape('page:1', from.shapeId);
    const saved = diagram.save();
    diagram.dispose();
    const archive = await JSZip.loadAsync(saved);
    const pageXml = await archive.file('visio/pages/page1.xml')!.async('text');
    expect((pageXml.match(/<Connect /g) ?? []).length).toBe((beforeDelete.match(/<Connect /g) ?? []).length - 1);
    const reopened = openDiagram(saved, { clientId: 7106 });
    try {
      expect(reopened.snapshot().pages[0].shapes).not.toContainEqual(expect.objectContaining({ id: from.shapeId }));
    } finally { reopened.dispose(); }
  });
});
