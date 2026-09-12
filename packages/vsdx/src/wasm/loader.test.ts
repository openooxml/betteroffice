import { beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { VsdxDocument, VsdxRenderer } from './generated/vsdx_wasm.js';
import { initWasm, openDiagram } from '../index';
import type { FormulaShapeDraft } from '../types';

const root = resolve(import.meta.dir, '../../../..');
let foundation: Uint8Array;
let nestedGroups: Uint8Array;
let groupedGlue: Uint8Array;
let textAccounting: Uint8Array;
let demo: Uint8Array;

beforeAll(async () => {
  const [wasm, foundationBytes, nestedGroupsBytes, groupedGlueBytes, textAccountingBytes, demoBytes] = await Promise.all([
    readFile(resolve(import.meta.dir, 'generated/vsdx_wasm_bg.wasm')),
    readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/foundation.vsdx')),
    readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/nested-groups.vsdx')),
    readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/grouped-glue.vsdx')),
    readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/text-accounting.vsdx')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.vsdx')),
  ]);
  await initWasm(wasm);
  foundation = foundationBytes;
  nestedGroups = nestedGroupsBytes;
  groupedGlue = groupedGlueBytes;
  textAccounting = textAccountingBytes;
  demo = demoBytes;
});

describe('VSDX wasm boundary', () => {
  test('opens and disposes a diagram', () => {
    const diagram = openDiagram(foundation, { clientId: 9001 });
    expect(diagram.snapshot().pages).toHaveLength(1);
    diagram.dispose();
    expect(() => diagram.snapshot()).toThrow('diagram handle is disposed');
  });

  test('accepts only v4 display lists', () => {
    const diagram = openDiagram(foundation, { clientId: 9002 });
    expect(diagram.layoutPage(0).contractVersion).toBe(4);

    const layoutPageJson = VsdxRenderer.prototype.layoutPageJson;
    VsdxRenderer.prototype.layoutPageJson = () => JSON.stringify({ contractVersion: 2 });
    try {
      expect(() => diagram.layoutPage(0)).toThrow('unsupported VSDX display-list contract version 2');
    } finally {
      VsdxRenderer.prototype.layoutPageJson = layoutPageJson;
      diagram.dispose();
    }
  });

  test('frees the document when applying the initial update fails', () => {
    const applyUpdateJson = VsdxDocument.prototype.applyUpdateJson;
    const docFree = VsdxDocument.prototype.free;
    const freed: string[] = [];
    VsdxDocument.prototype.applyUpdateJson = () => { throw new Error('boom'); };
    VsdxDocument.prototype.free = function (...args) { freed.push('doc'); return docFree.apply(this, args); };
    try {
      expect(() => openDiagram(foundation, { clientId: 9050, initialUpdate: new Uint8Array([1]) })).toThrow('boom');
      expect(freed).toEqual(['doc']);
    } finally {
      VsdxDocument.prototype.applyUpdateJson = applyUpdateJson;
      VsdxDocument.prototype.free = docFree;
    }
  });

  test('frees the document and renderer when font registration fails', () => {
    const registerFont = VsdxRenderer.prototype.registerFont;
    const rendererFree = VsdxRenderer.prototype.free;
    const docFree = VsdxDocument.prototype.free;
    const freed: string[] = [];
    VsdxRenderer.prototype.registerFont = () => { throw new Error('font boom'); };
    VsdxRenderer.prototype.free = function (...args) { freed.push('renderer'); return rendererFree.apply(this, args); };
    VsdxDocument.prototype.free = function (...args) { freed.push('doc'); return docFree.apply(this, args); };
    try {
      expect(() => openDiagram(foundation, { clientId: 9051, fonts: [{ family: 'Test', bytes: new Uint8Array([0]) }] })).toThrow('font boom');
      expect(freed).toEqual(['renderer', 'doc']);
    } finally {
      VsdxRenderer.prototype.registerFont = registerFont;
      VsdxRenderer.prototype.free = rendererFree;
      VsdxDocument.prototype.free = docFree;
    }
  });

  test('returns no hit before layout and for misses', () => {
    const diagram = openDiagram(foundation, { clientId: 9003 });
    expect(diagram.hitTest(0, 0)).toBeNull();
    diagram.layoutPage(0);
    expect(diagram.hitTest(-1, -1)).toBeNull();
    diagram.dispose();
  });

  test('decodes wasm text diagnostic categories using wire casing', () => {
    const diagram = openDiagram(textAccounting, { clientId: 9011 });
    const diagnostics = diagram.layoutPage(0).primitives.flatMap(primitive => primitive.kind === 'textBox' ? primitive.paragraphs.flatMap(paragraph => paragraph.runs.flatMap(run => run.diagnostics)) : []);
    expect(diagnostics).toContainEqual(expect.objectContaining({ category: 'fidelity', code: 'unregistered-font' }));
    diagram.dispose();
  });

  test('delivers local and remote frames', () => {
    const diagram = openDiagram(foundation, { clientId: 9004 });
    const drainUpdateEvent = VsdxDocument.prototype.drainUpdateEvent;
    const frames = [Uint8Array.of(0, 9), Uint8Array.of(1, 8), new Uint8Array()];
    const received: Array<{ update: Uint8Array; origin: string }> = [];
    VsdxDocument.prototype.drainUpdateEvent = () => frames.shift() ?? new Uint8Array();
    try {
      const unsubscribe = diagram.onUpdate((update, origin) => received.push({ update, origin }));
      diagram.setCellFormula('page:1', 'page:1:shape:1', { cellName: 'Both' }, '3');
      expect(received.map(({ origin }) => origin)).toEqual(['local', 'remote']);
      expect(received.map(({ update }) => [...update])).toEqual([[9], [8]]);
      unsubscribe();
    } finally {
      VsdxDocument.prototype.drainUpdateEvent = drainUpdateEvent;
    }

    diagram.dispose();
  });

  test('delivers a full-state resync after a genuine observation overflow', () => {
    const diagram = openDiagram(foundation, { clientId: 9008 });
    const drainUpdateEvent = VsdxDocument.prototype.drainUpdateEvent;
    const resyncs: Uint8Array[] = [];
    const updates: Uint8Array[] = [];
    const unsubscribeUpdate = diagram.onUpdate(update => updates.push(update));
    const unsubscribeResync = diagram.onResync(({ update }) => resyncs.push(update));
    VsdxDocument.prototype.drainUpdateEvent = () => new Uint8Array();
    try {
      for (let index = 0; index <= 1024; index++) diagram.setCellFormula('page:1', 'page:1:shape:1', { cellName: 'Both' }, String(index));
    } finally {
      VsdxDocument.prototype.drainUpdateEvent = drainUpdateEvent;
    }
    diagram.setCellFormula('page:1', 'page:1:shape:1', { cellName: 'Both' }, '1025');
    expect(updates).toHaveLength(0);
    expect(resyncs).toHaveLength(1);
    const recovered = openDiagram(foundation, { clientId: 9009 });
    expect(recovered.applyUpdate(resyncs[0]).pages[0].shapes[0].cells.find(cell => cell.name === 'Both')?.formula).toBe('1025');
    recovered.dispose();
    unsubscribeUpdate();
    unsubscribeResync();
    diagram.dispose();
  });

  test('returns committed media and plain missing-media errors', () => {
    const diagram = openDiagram(nestedGroups, { clientId: 9005 });
    expect([...diagram.mediaBytes('visio/media/image1.png')]).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
    expect(() => diagram.mediaBytes('visio/media/missing.png')).toThrow('invalid diagram state: media part was not found');
    diagram.dispose();
  });

  test('refuses guarded cell edits through the TypeScript surface', () => {
    const diagram = openDiagram(foundation, { clientId: 9006 });
    diagram.setCellFormula('page:1', 'page:1:shape:1', { cellName: 'FOnly' }, 'GUARD(1)');
    expect(() => diagram.setCellFormula('page:1', 'page:1:shape:1', { cellName: 'FOnly' }, '2')).toThrow('GUARD protects the requested cell');
    diagram.dispose();
  });

  test('persists an edit through save and reopen while preserving untouched parts', async () => {
    const pageId = 'page:1';
    const shapeId = 'page:1:shape:1';
    const editedPart = 'visio/pages/page1.xml';
    const diagram = openDiagram(demo, { clientId: 9012 });
    diagram.moveShape(pageId, shapeId, '3.25', '4.5');
    diagram.setCellFormula(pageId, shapeId, { cellName: 'FillForegnd' }, '"#336699"');
    const saved = diagram.save();
    diagram.dispose();

    // Microsoft Visio reopening is not possible in this environment; reopen through the production WASM boundary instead.
    const reopened = openDiagram(saved, { clientId: 9013 });
    const cells = reopened.snapshot().pages[0].shapes[0].cells;
    expect(cells.find(cell => cell.name === 'PinX')?.formula).toBe('3.25');
    expect(cells.find(cell => cell.name === 'PinY')?.formula).toBe('4.5');
    expect(cells.find(cell => cell.name === 'FillForegnd')?.formula).toBe('"#336699"');
    reopened.dispose();

    const [original, result] = await Promise.all([JSZip.loadAsync(demo), JSZip.loadAsync(saved)]);
    expect(Object.keys(result.files).sort()).toEqual(Object.keys(original.files).sort());
    await Promise.all(Object.keys(original.files).filter(path => path !== editedPart).map(async path => {
      expect(await result.file(path)!.async('uint8array')).toEqual(await original.file(path)!.async('uint8array'));
    }));
  });

  test('persists an added shape through the public WASM boundary', async () => {
    const pageId = 'page:1';
    const editedPart = 'visio/pages/page1.xml';
    const diagram = openDiagram(foundation, { clientId: 9016 });
    const receipt = diagram.addShape(pageId, { name: 'Added rectangle', cells: [
      { locator: { cellName: 'PinX' }, formula: '7' },
      { locator: { cellName: 'Width' }, formula: '2' },
      { locator: { section: 'User', rowIndex: 0, cellName: 'Value' }, formula: '3' },
    ] });
    const live = diagram.snapshot();
    expect(live.pages[0].shapes.find(shape => shape.id === receipt.shapeId)).toEqual(expect.objectContaining({ sourceId: 2 }));
    const saved = diagram.save();
    diagram.dispose();

    const reopened = openDiagram(saved, { clientId: 9017 });
    const added = reopened.snapshot().pages[0].shapes.find(shape => shape.id === 'page:1:shape:2');
    expect(added).toEqual(expect.objectContaining({ id: 'page:1:shape:2', name: 'Added rectangle' }));
    expect(added?.cells).toEqual(expect.arrayContaining([
      expect.objectContaining({ name: 'PinX', formula: '7' }),
      expect.objectContaining({ name: 'Width', formula: '2' }),
      expect.objectContaining({ name: 'Value', formula: '3', locator: expect.objectContaining({ section: 'User', row: { index: 0 } }) }),
    ]));
    expect(reopened.snapshot().pages[0].shapes.find(shape => shape.id === 'page:1:shape:2')).toEqual(expect.objectContaining({ sourceId: 2 }));
    reopened.dispose();
    await expectUntouchedParts(foundation, saved, editedPart);
  });

  test('persists deletion and removes dependent Connect records through deleteShapeJson', async () => {
    const pageId = 'page:1';
    const shapeId = 'page:1:shape:1';
    const editedPart = 'visio/pages/page1.xml';
    const diagram = openDiagram(foundation, { clientId: 9018 });
    diagram.deleteShape(pageId, shapeId);
    const saved = diagram.save();
    diagram.dispose();

    const reopened = openDiagram(saved, { clientId: 9019 });
    expect(reopened.snapshot().pages[0].shapes).not.toContainEqual(expect.objectContaining({ id: shapeId }));
    reopened.dispose();
    const archive = await JSZip.loadAsync(saved);
    const pageXml = await archive.file(editedPart)!.async('text');
    expect(pageXml).not.toContain("FromSheet='1'");
    expect(pageXml).not.toContain("ToSheet='1'");
    await expectUntouchedParts(foundation, saved, editedPart);
  });

  test('persists shape order through save and reopen', async () => {
    const pageId = 'page:1';
    const editedPart = 'visio/pages/page1.xml';
    const diagram = openDiagram(demo, { clientId: 9020 });
    const originalOrder = diagram.snapshot().pages[0].shapes.map(shape => shape.id);
    diagram.reorderShape(pageId, originalOrder[2], 0);
    const saved = diagram.save();
    diagram.dispose();

    const reopened = openDiagram(saved, { clientId: 9021 });
    expect(reopened.snapshot().pages[0].shapes.map(shape => shape.id)).toEqual([originalOrder[2], originalOrder[0], originalOrder[1]]);
    reopened.dispose();
    await expectUntouchedParts(demo, saved, editedPart);
  });

  test('persists page order through save and reopen', async () => {
    const diagram = openDiagram(demo, { clientId: 9042 });
    const originalPages = diagram.snapshot().pages;
    diagram.reorderPage(originalPages[1].id, 0);
    const snapshot = diagram.snapshot();
    const expected = snapshot.pages[0];
    const expectedPrimitiveIds = expected.shapes.map(shape => `${expected.sourcePartPath}:${shape.sourceId}`);
    const live = diagram.layoutPage(0).primitives.map(primitive => primitive.id);
    expect(live).toEqual(expect.arrayContaining(expectedPrimitiveIds));
    const saved = diagram.save();
    diagram.dispose();

    const reopened = openDiagram(saved, { clientId: 9043 });
    expect(reopened.snapshot()).toEqual(snapshot);
    expect(reopened.layoutPage(0).primitives.map(primitive => primitive.id)).toEqual(live);
    reopened.dispose();
    await expectUntouchedParts(demo, saved, ['visio/pages/pages.xml', 'visio/pages/_rels/pages.xml.rels']);
  });

  test('persists page reordering with shape additions and deletions', async () => {
    const diagram = openDiagram(demo, { clientId: 9044 });
    const [firstPage, secondPage] = diagram.snapshot().pages;
    const removed = firstPage.shapes[0].id;
    const added = diagram.addShape(firstPage.id, { name: 'Added', cells: [] }).shapeId;
    diagram.deleteShape(firstPage.id, removed);
    diagram.reorderPage(secondPage.id, 0);
    const saved = diagram.save();
    diagram.dispose();

    const reopened = openDiagram(saved, { clientId: 9045 });
    expect(reopened.snapshot().pages.map(page => page.id)).toEqual([secondPage.id, firstPage.id]);
    expect(reopened.snapshot().pages[1].shapes).not.toContainEqual(expect.objectContaining({ id: removed }));
    expect(reopened.snapshot().pages[1].shapes).toContainEqual(expect.objectContaining({ name: 'Added' }));
    expect(added).not.toBe(removed);
    reopened.dispose();
    await expectUntouchedParts(demo, saved, ['visio/pages/pages.xml', 'visio/pages/_rels/pages.xml.rels', 'visio/pages/page1.xml']);
  });

  test('keeps added-shape cells in the live render projection', () => {
    const diagram = openDiagram(foundation, { clientId: 9046 });
    diagram.addShape('page:1', { cells: [
      { locator: { cellName: 'PinX' }, formula: '2' },
      { locator: { cellName: 'PinY' }, formula: '2' },
      { locator: { cellName: 'Width' }, formula: '1' },
      { locator: { cellName: 'Height' }, formula: '1' },
      { locator: { cellName: 'LocPinX' }, formula: '0' },
      { locator: { cellName: 'LocPinY' }, formula: '0' },
      { locator: { section: 'Geometry', rowIndex: 0, cellName: 'X' }, formula: '0' },
      { locator: { section: 'Geometry', rowIndex: 0, cellName: 'Y' }, formula: '0' },
      { locator: { section: 'Geometry', rowIndex: 1, cellName: 'X' }, formula: '1' },
      { locator: { section: 'Geometry', rowIndex: 1, cellName: 'Y' }, formula: '1' },
    ] });
    expect(diagram.layoutPage(0).primitives).toContainEqual(expect.objectContaining({
      id: 'visio/pages/page1.xml:2',
      x: 2,
      y: 2,
      width: 1,
      height: 1,
    }));
    diagram.dispose();
  });

  test('persists an added shape at its requested order through save and reopen', async () => {
    const pageId = 'page:1';
    const editedPart = 'visio/pages/page1.xml';
    const diagram = openDiagram(foundation, { clientId: 9022 });
    const original = diagram.snapshot().pages[0].shapes[0].id;
    const added = diagram.addShape(pageId, { cells: [] }).shapeId;
    diagram.reorderShape(pageId, added, 0);
    const saved = diagram.save();
    diagram.dispose();
    const reopened = openDiagram(saved, { clientId: 9023 });
    expect(reopened.snapshot().pages[0].shapes.map(shape => shape.id)).toEqual(['page:1:shape:2', original]);
    reopened.dispose();
    await expectUntouchedParts(foundation, saved, editedPart);
  });

  test('persists a colliding added shape after deleting the original', async () => {
    const pageId = 'page:1';
    const editedPart = 'visio/pages/page1.xml';
    const diagram = openDiagram(foundation, { clientId: 9034 });
    const original = diagram.snapshot().pages[0].shapes[0].id;
    const added = diagram.addShape(pageId, { name: 'Collision', cells: [] }).shapeId;
    diagram.deleteShape(pageId, original);
    const saved = diagram.save();
    diagram.dispose();

    const reopened = openDiagram(saved, { clientId: 9035 });
    expect(reopened.snapshot().pages[0].shapes).toEqual([
      expect.objectContaining({ id: 'page:1:shape:2', name: 'Collision' }),
    ]);
    expect(reopened.snapshot().pages[0].shapes).not.toContainEqual(expect.objectContaining({ id: original }));
    expect(added).not.toBe(original);
    reopened.dispose();
    await expectUntouchedParts(foundation, saved, editedPart);
  });

  test('reorders a colliding added shape without trapping', async () => {
    const pageId = 'page:1';
    const editedPart = 'visio/pages/page1.xml';
    const diagram = openDiagram(foundation, { clientId: 9036 });
    const original = diagram.snapshot().pages[0].shapes[0].id;
    const added = diagram.addShape(pageId, { cells: [] }).shapeId;
    expect(() => diagram.reorderShape(pageId, added, 0)).not.toThrow();
    const saved = diagram.save();
    diagram.dispose();

    const reopened = openDiagram(saved, { clientId: 9037 });
    expect(reopened.snapshot().pages[0].shapes.map(shape => shape.id)).toEqual(['page:1:shape:2', original]);
    reopened.dispose();
    await expectUntouchedParts(foundation, saved, editedPart);
  });

  test('persists moving an original behind a non-colliding added shape', async () => {
    const pageId = 'page:1';
    const editedPart = 'visio/pages/page1.xml';
    const diagram = openDiagram(foundation, { clientId: 9038 });
    const original = diagram.snapshot().pages[0].shapes[0].id;
    diagram.addShape(pageId, { cells: [] });
    diagram.reorderShape(pageId, original, 1);
    const saved = diagram.save();
    diagram.dispose();

    const reopened = openDiagram(saved, { clientId: 9039 });
    expect(reopened.snapshot().pages[0].shapes.map(shape => shape.id)).toEqual(['page:1:shape:2', original]);
    reopened.dispose();
    await expectUntouchedParts(foundation, saved, editedPart);
  });

  test('persists adding, reordering, and deleting in one session', async () => {
    const pageId = 'page:1';
    const editedPart = 'visio/pages/page1.xml';
    const diagram = openDiagram(foundation, { clientId: 9040 });
    const original = diagram.snapshot().pages[0].shapes[0].id;
    const added = diagram.addShape(pageId, { name: 'Combined', cells: [] }).shapeId;
    diagram.reorderShape(pageId, added, 0);
    diagram.deleteShape(pageId, original);
    const saved = diagram.save();
    diagram.dispose();

    const reopened = openDiagram(saved, { clientId: 9041 });
    expect(reopened.snapshot().pages[0].shapes).toEqual([
      expect.objectContaining({ id: 'page:1:shape:2', name: 'Combined' }),
    ]);
    reopened.dispose();
    await expectUntouchedParts(foundation, saved, editedPart);
  });

  test('persists edits to an added shape through save and reopen', async () => {
    const pageId = 'page:1';
    const editedPart = 'visio/pages/page1.xml';
    const diagram = openDiagram(foundation, { clientId: 9024 });
    const added = diagram.addShape(pageId, { cells: [{ locator: { cellName: 'Width' }, formula: '2' }] }).shapeId;
    diagram.setCellFormula(pageId, added, { cellName: 'Width' }, '3');
    const saved = diagram.save();
    diagram.dispose();
    const reopened = openDiagram(saved, { clientId: 9025 });
    expect(reopened.snapshot().pages[0].shapes.find(shape => shape.id === 'page:1:shape:2')?.cells).toContainEqual(expect.objectContaining({ name: 'Width', formula: '3' }));
    reopened.dispose();
    await expectUntouchedParts(foundation, saved, editedPart);
  });

  test('persists deleting a group child while retaining its group and siblings', async () => {
    const pageId = 'page:1';
    const editedPart = 'visio/pages/page1.xml';
    const diagram = openDiagram(nestedGroups, { clientId: 9026 });
    const group = findGroupWithSiblings(diagram.snapshot().pages[0].shapes)!;
    const [removed, sibling] = group.children;
    expect(sibling).toBeDefined();
    diagram.deleteShape(pageId, removed.id);
    const saved = diagram.save();
    diagram.dispose();
    const reopened = openDiagram(saved, { clientId: 9027 });
    const savedGroup = flattenShapes(reopened.snapshot().pages[0].shapes).find(shape => shape.id === group.id)!;
    expect(savedGroup.children.map(shape => shape.id)).not.toContain(removed.id);
    expect(savedGroup.children.map(shape => shape.id)).toContain(sibling.id);
    reopened.dispose();
    await expectUntouchedParts(nestedGroups, saved, editedPart);
  });

  test('persists reordering shapes within a group through save and reopen', async () => {
    const pageId = 'page:1';
    const editedPart = 'visio/pages/page1.xml';
    const diagram = openDiagram(nestedGroups, { clientId: 9028 });
    const group = findGroupWithSiblings(diagram.snapshot().pages[0].shapes)!;
    const originalOrder = group.children.map(shape => shape.id);
    const [second] = originalOrder.slice(1);
    expect(second).toBeDefined();
    const expectedOrder = [second, ...originalOrder.filter(shapeId => shapeId !== second)];
    diagram.reorderShape(pageId, second, 0);
    expect(flattenShapes(diagram.snapshot().pages[0].shapes).find(shape => shape.id === group.id)!.children.map(shape => shape.id)).toEqual(expectedOrder);
    const saved = diagram.save();
    diagram.dispose();
    const reopened = openDiagram(saved, { clientId: 9029 });
    expect(flattenShapes(reopened.snapshot().pages[0].shapes).find(shape => shape.id === group.id)!.children.map(shape => shape.id)).toEqual(expectedOrder);
    reopened.dispose();
    await expectUntouchedParts(nestedGroups, saved, editedPart);
  });

  test('persists deleting a group and removes connects to its subtree', async () => {
    const pageId = 'page:1';
    const editedPart = 'visio/pages/page1.xml';
    const diagram = openDiagram(groupedGlue, { clientId: 9030 });
    const group = diagram.snapshot().pages[0].shapes.find(shape => shape.children.length > 0)!;
    const deletedIds = flattenShapes([group]).map(shape => shape.id.split(':shape:').pop()!);
    diagram.deleteShape(pageId, group.id);
    const saved = diagram.save();
    diagram.dispose();
    const reopened = openDiagram(saved, { clientId: 9031 });
    expect(flattenShapes(reopened.snapshot().pages[0].shapes)).not.toContainEqual(expect.objectContaining({ id: group.id }));
    reopened.dispose();
    const archive = await JSZip.loadAsync(saved);
    const pageXml = await archive.file(editedPart)!.async('text');
    for (const id of deletedIds) expect(pageXml).not.toMatch(new RegExp(`(?:FromSheet|ToSheet)=["']${id}["']`));
    await expectUntouchedParts(groupedGlue, saved, editedPart);
  });

  test('preserves a LockDelete shape through a save round trip', async () => {
    const editedPart = 'visio/pages/page1.xml';
    const diagram = openDiagram(foundation, { clientId: 9032 });
    diagram.addShape('page:1', { cells: [{ locator: { cellName: 'LockDelete' }, formula: '1' }] });
    const saved = diagram.save();
    diagram.dispose();
    const reopened = openDiagram(saved, { clientId: 9033 });
    expect(reopened.snapshot().pages[0].shapes).toContainEqual(expect.objectContaining({ id: 'page:1:shape:2' }));
    reopened.dispose();
    await expectUntouchedParts(foundation, saved, editedPart);
  });

  test('aborts a guarded move batch without changing the save bytes', () => {
    const pageId = 'page:1';
    const shapeId = 'page:1:shape:1';
    const prepared = openDiagram(demo, { clientId: 9014 });
    prepared.setCellFormula(pageId, shapeId, { cellName: 'PinY' }, 'GUARD(1)');
    const original = prepared.save();
    prepared.dispose();

    const diagram = openDiagram(original, { clientId: 9015 });
    expect(() => diagram.moveShape(pageId, shapeId, '3.25', '4.5')).toThrow('GUARD protects the requested cell');
    expect(diagram.save()).toEqual(original);
    diagram.dispose();
  });

  test('does not change the snapshot when applyUpdate rejects malformed bytes', () => {
    const diagram = openDiagram(foundation, { clientId: 9010 });
    const before = diagram.snapshot();
    expect(() => diagram.applyUpdate(Uint8Array.of(0))).toThrow('invalid yrs update');
    expect(diagram.snapshot()).toEqual(before);
    diagram.dispose();
  });

  test('keeps shape drafts formula-only at the type boundary', () => {
    const draft: FormulaShapeDraft = { cells: [] };
    expect(draft.cells).toEqual([]);
    // @ts-expect-error Shape cells accept formulas, never cached values.
    const invalid: FormulaShapeDraft = { cells: [{ locator: { cellName: 'Width' }, formula: '1', value: '1' }] };
    expect(invalid).toBeDefined();
  });

  test('adds a shape from the declared cell locator shape', () => {
    const diagram = openDiagram(foundation, { clientId: 9012 });
    const draft: FormulaShapeDraft = {
      name: 'Added',
      cells: [
        { locator: { cellName: 'Width' }, formula: '1' },
        { locator: { section: 'Geometry', rowIndex: 0, cellName: 'X' }, formula: '2' },
      ],
    };
    const receipt = diagram.addShape('page:1', draft);
    const added = diagram.snapshot().pages[0].shapes.find(shape => shape.id === receipt.shapeId);
    expect(added?.cells.find(cell => cell.name === 'Width')?.formula).toBe('1');
    const x = added?.cells.find(cell => cell.name === 'X');
    expect(x?.formula).toBe('2');
    expect(x?.locator).toEqual(expect.objectContaining({ section: 'Geometry', row: { index: 0 }, cellName: 'X' }));
    diagram.dispose();
  });

  test('does not reenter update listeners before the outer call unwinds', () => {
    const diagram = openDiagram(foundation, { clientId: 9007 });
    const sequence: string[] = [];
    let nested = false;
    const unsubscribe = diagram.onUpdate(() => {
      sequence.push('start');
      if (!nested) {
        nested = true;
        diagram.setCellFormula('page:1', 'page:1:shape:1', { cellName: 'FOnly' }, '5');
      }
      sequence.push('end');
    });
    diagram.setCellFormula('page:1', 'page:1:shape:1', { cellName: 'FOnly' }, '4');
    expect(sequence).toEqual(['start', 'end', 'start', 'end']);
    unsubscribe();
    diagram.dispose();
  });
});

async function expectUntouchedParts(originalBytes: Uint8Array, saved: Uint8Array, editedParts: string | string[]): Promise<void> {
  const [original, result] = await Promise.all([JSZip.loadAsync(originalBytes), JSZip.loadAsync(saved)]);
  expect(Object.keys(result.files).sort()).toEqual(Object.keys(original.files).sort());
  const changed = new Set(typeof editedParts === 'string' ? [editedParts] : editedParts);
  await Promise.all(Object.keys(original.files).filter(path => !changed.has(path)).map(async path => {
    expect(await result.file(path)!.async('uint8array')).toEqual(await original.file(path)!.async('uint8array'));
  }));
}

function flattenShapes<T extends { children: T[] }>(shapes: T[]): T[] {
  return shapes.flatMap(shape => [shape, ...flattenShapes(shape.children)]);
}

function findGroupWithSiblings<T extends { children: T[] }>(shapes: T[]): T | undefined {
  return flattenShapes(shapes).find(shape => shape.children.length > 1);
}

test('hit testing returns a shape ID accepted by editing commands', () => {
  const diagram = openDiagram(demo, { clientId: 9080 });
  try {
    const page = diagram.snapshot().pages[0];
    const frame = diagram.layoutPage(0);
    let hit = null as ReturnType<typeof diagram.hitTest>;
    for (let y = 12; y < frame.height && !hit; y += 24) for (let x = 12; x < frame.width && !hit; x += 24) hit = diagram.hitTest(x, y);
    expect(hit).not.toBeNull();
    expect(flattenShapes(page.shapes).some(shape => shape.id === hit!.shapeId)).toBe(true);
    diagram.setCellFormula(page.id, hit!.shapeId, { cellName: 'PinX' }, '0.25');
    expect(flattenShapes(diagram.snapshot().pages[0].shapes).find(shape => shape.id === hit!.shapeId)?.cells.find(cell => cell.name === 'PinX')?.formula).toBe('0.25');
  } finally { diagram.dispose(); }
});
