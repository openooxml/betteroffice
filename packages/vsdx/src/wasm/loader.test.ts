import { beforeAll, describe, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { VsdxDocument, VsdxRenderer } from './generated/vsdx_wasm.js';
import { initWasm, openDiagram } from '../index';
import type { FormulaShapeDraft } from '../types';

const root = resolve(import.meta.dir, '../../../..');
let foundation: Uint8Array;
let nestedGroups: Uint8Array;

beforeAll(async () => {
  const [wasm, foundationBytes, nestedGroupsBytes] = await Promise.all([
    readFile(resolve(import.meta.dir, 'generated/vsdx_wasm_bg.wasm')),
    readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/foundation.vsdx')),
    readFile(resolve(root, 'crates/vsdx-parse/tests/fixtures/nested-groups.vsdx')),
  ]);
  await initWasm(wasm);
  foundation = foundationBytes;
  nestedGroups = nestedGroupsBytes;
});

describe('VSDX wasm boundary', () => {
  test('opens and disposes a diagram', () => {
    const diagram = openDiagram(foundation, { clientId: 9001 });
    expect(diagram.snapshot().pages).toHaveLength(1);
    diagram.dispose();
    expect(() => diagram.snapshot()).toThrow('diagram handle is disposed');
  });

  test('accepts only v3 display lists', () => {
    const diagram = openDiagram(foundation, { clientId: 9002 });
    expect(diagram.layoutPage(0).contractVersion).toBe(3);

    const layoutPageJson = VsdxRenderer.prototype.layoutPageJson;
    VsdxRenderer.prototype.layoutPageJson = () => JSON.stringify({ contractVersion: 2 });
    try {
      expect(() => diagram.layoutPage(0)).toThrow('unsupported VSDX display-list contract version 2');
    } finally {
      VsdxRenderer.prototype.layoutPageJson = layoutPageJson;
      diagram.dispose();
    }
  });

  test('returns no hit before layout and for misses', () => {
    const diagram = openDiagram(foundation, { clientId: 9003 });
    expect(diagram.hitTest(0, 0)).toBeNull();
    diagram.layoutPage(0);
    expect(diagram.hitTest(-1, -1)).toBeNull();
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
      diagram.setCellFormula('page:1', 'page:1:shape:1', { cellName: 'FOnly' }, '3');
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
      for (let index = 0; index <= 1024; index++) diagram.setCellFormula('page:1', 'page:1:shape:1', { cellName: 'FOnly' }, String(index));
    } finally {
      VsdxDocument.prototype.drainUpdateEvent = drainUpdateEvent;
    }
    diagram.setCellFormula('page:1', 'page:1:shape:1', { cellName: 'FOnly' }, '1025');
    expect(updates).toHaveLength(0);
    expect(resyncs).toHaveLength(1);
    const recovered = openDiagram(foundation, { clientId: 9009 });
    expect(recovered.applyUpdate(resyncs[0]).pages[0].shapes[0].cells.find(cell => cell.name === 'FOnly')?.formula).toBe('1025');
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

  test('does not change the snapshot when applyUpdate rejects malformed bytes', () => {
    const diagram = openDiagram(foundation, { clientId: 9010 });
    const before = diagram.snapshot();
    expect(() => diagram.applyUpdate(Uint8Array.of(0))).toThrow('invalid yrs update');
    expect(diagram.snapshot()).toEqual(before);
    diagram.dispose();
  });

  test('keeps shape drafts formula-only at the type boundary', () => {
    const draft: FormulaShapeDraft = { sourceId: 2, cells: [] };
    expect(draft.sourceId).toBe(2);
    // @ts-expect-error Shape cells accept formulas, never cached values.
    const invalid: FormulaShapeDraft = { sourceId: 2, cells: [{ locator: {}, name: 'Width', value: '1' }] };
    expect(invalid).toBeDefined();
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
