import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import type { CollaborationReplica } from '../collaboration/types';
import { XlsxDocument } from './generated/xlsx_wasm.js';
import { initWasm, openWorkbook } from './loader';
import type { WorkbookHandle, WorkbookUpdateOrigin } from './loader';

const FIXTURE = resolve(import.meta.dir, '../../test-fixtures/sample.xlsx');
const CHART_FIXTURE = resolve(import.meta.dir, '../../test-fixtures/charts.xlsx');
const WASM = resolve(import.meta.dir, './generated/xlsx_wasm_bg.wasm');

function sampleBytes(): Uint8Array {
  return new Uint8Array(readFileSync(FIXTURE));
}

function chartBytes(): Uint8Array {
  return new Uint8Array(readFileSync(CHART_FIXTURE));
}

function collaborative(clientId: number): WorkbookHandle {
  return openWorkbook(sampleBytes(), { collaborative: true, clientId });
}

function collaborativeCharts(clientId: number): WorkbookHandle {
  return openWorkbook(chartBytes(), { collaborative: true, clientId });
}

function requireReplica(replica: CollaborationReplica): CollaborationReplica {
  return replica;
}

describe('wasm collaboration', () => {
  beforeAll(() => initWasm(new Uint8Array(readFileSync(WASM))));

  it('opens explicit and generated replicas and validates client IDs', () => {
    expect(() => openWorkbook(sampleBytes(), { clientId: 1 })).toThrow(
      'clientId requires collaborative mode'
    );
    for (const clientId of [0, -1, 1.5, Number.MAX_SAFE_INTEGER + 1, Infinity, NaN]) {
      expect(() => openWorkbook(sampleBytes(), { collaborative: true, clientId })).toThrow(
        'clientId must be a nonzero safe integer'
      );
    }
    expect(() =>
      openWorkbook(Uint8Array.of(1, 2, 3), { collaborative: true, clientId: 77 })
    ).toThrow();

    const explicit = collaborative(Number.MAX_SAFE_INTEGER);
    const generated = openWorkbook(sampleBytes(), { collaborative: true });
    const standalone = openWorkbook(sampleBytes());
    try {
      expect(explicit.clientId).toBe(Number.MAX_SAFE_INTEGER);
      expect(explicit.sheetInfo().sheetIds).toEqual(['sheet:0', 'sheet:1', 'sheet:2']);
      expect(generated.clientId).toBeGreaterThan(0);
      expect(generated.clientId).toBeLessThanOrEqual(Number.MAX_SAFE_INTEGER);
      expect(requireReplica(explicit)).toBe(explicit);
      expect(standalone.encodeStateVector()).toBeInstanceOf(Uint8Array);
      expect(standalone.encodeStateAsUpdate()).toBeInstanceOf(Uint8Array);
      expect(() => standalone.applyUpdate(Uint8Array.of(0, 0))).toThrow(
        'remote updates require a collaborative workbook'
      );
    } finally {
      explicit.dispose();
      generated.dispose();
      standalone.dispose();
    }
  });

  it('validates client IDs at the generated wasm boundary before casting', () => {
    for (const clientId of [0, -1, 1.5, Number.MAX_SAFE_INTEGER + 1, Infinity, NaN]) {
      expect(() => XlsxDocument.openCollaborative(sampleBytes(), clientId)).toThrow(
        'client ID must be a nonzero integer no greater than Number.MAX_SAFE_INTEGER'
      );
    }
  });

  it('uses all 53 random client ID bits', () => {
    const crypto = globalThis.crypto;
    const original = crypto.getRandomValues;
    crypto.getRandomValues = ((array: Uint32Array) => {
      array[0] = 0x10_0000;
      array[1] = 7;
      return array;
    }) as typeof crypto.getRandomValues;
    let handle: WorkbookHandle | undefined;
    try {
      handle = openWorkbook(sampleBytes(), { collaborative: true });
      expect(handle.clientId).toBe(0x10_0000 * 0x1_0000_0000 + 7);
      expect(handle.clientId).toBeGreaterThan(0xffff_ffff);
    } finally {
      handle?.dispose();
      crypto.getRandomValues = original;
    }
  });

  it('encodes vectors and diffs and converges concurrent edits', () => {
    const left = collaborative(1001);
    const right = collaborative(1002);
    try {
      const baseline = left.encodeStateVector();
      expect([...right.encodeStateVector()]).toEqual([...baseline]);
      expect([...left.encodeStateAsUpdate()]).toEqual([...right.encodeStateAsUpdate()]);
      expect([...left.encodeStateAsUpdate(right.encodeStateVector())]).toEqual([0, 0]);

      left.editCell(0, 19, 0, 'left');
      right.editCell(0, 19, 1, 'right');
      const leftUpdate = left.encodeStateAsUpdate(baseline);
      const rightUpdate = right.encodeStateAsUpdate(baseline);
      left.applyUpdate(rightUpdate);
      right.applyUpdate(leftUpdate);
      expect(left.cell(0, 19, 0).input).toBe('left');
      expect(right.cell(0, 19, 0).input).toBe('left');
      expect(left.cell(0, 19, 1).input).toBe('right');
      expect(right.cell(0, 19, 1).input).toBe('right');
      expect([...left.encodeStateVector()]).toEqual([...right.encodeStateVector()]);

      const leftBefore = left.encodeStateVector();
      const rightBefore = right.encodeStateVector();
      left.editCell(0, 19, 2, 'left-wins-or-loses');
      right.editCell(0, 19, 2, 'right-wins-or-loses');
      const concurrentLeft = left.encodeStateAsUpdate(rightBefore);
      const concurrentRight = right.encodeStateAsUpdate(leftBefore);
      left.applyUpdate(concurrentRight);
      right.applyUpdate(concurrentLeft);
      expect(left.cell(0, 19, 2).input).toBe(right.cell(0, 19, 2).input);
      expect([...left.encodeStateVector()]).toEqual([...right.encodeStateVector()]);
    } finally {
      left.dispose();
      right.dispose();
    }
  });

  it('round-trips formatting and collaborative undo between handles', () => {
    const left = collaborative(1101);
    const right = collaborative(1102);
    try {
      const beforeDisplay = right.displayList({ x: 0, y: 0, width: 500, height: 120 });
      left.patchRangeStyle(0, 'A1:B2', {
        bold: true,
        fillColor: '#ffcc00',
        textColor: '#123456',
      });
      left.setNumberFormat(0, 'A1:B2', { type: 'custom', pattern: '0.0000' });
      const captured = left.captureFormat(0, 'A1');
      left.applyFormat(0, 'C1', captured);

      const update = left.encodeStateAsUpdate(right.encodeStateVector());
      expect(right.applyUpdate(update).applied).toBe(true);
      expect(right.selectionFormatting(0, 'A1:B2')).toMatchObject({
        bold: true,
        fillColor: '#ffcc00',
        textColor: '#123456',
        numberFormat: 'custom',
        numberFormatPattern: '0.0000',
      });
      expect(right.selectionFormatting(0, 'C1')).toEqual(left.selectionFormatting(0, 'A1'));
      const afterDisplay = right.displayList({ x: 0, y: 0, width: 500, height: 120 });
      expect(afterDisplay).not.toEqual(beforeDisplay);
      expect(
        afterDisplay.commands.find((command) => command.op === 'text' && command.text === 'Q1')
      ).toMatchObject({ op: 'text', fontSize: 11, bold: true });
      expect(left.historyState()).toMatchObject({ canUndo: true, undoDepth: 3 });

      const rightBeforeUndo = right.encodeStateVector();
      expect(left.undo().applied).toBe(true);
      expect(right.applyUpdate(left.encodeStateAsUpdate(rightBeforeUndo)).applied).toBe(true);
      expect(right.selectionFormatting(0, 'C1').bold).toBe(false);
      expect([...left.encodeStateVector()]).toEqual([...right.encodeStateVector()]);
    } finally {
      left.dispose();
      right.dispose();
    }
  });

  it('delivers owned local and remote bytes and isolates listener exceptions', () => {
    const left = collaborative(2001);
    const right = collaborative(2002);
    const mutated: Uint8Array[] = [];
    const received: Array<{ update: Uint8Array; origin: WorkbookUpdateOrigin }> = [];
    const remote: Array<{ update: Uint8Array; origin: WorkbookUpdateOrigin }> = [];
    try {
      left.onUpdate((update) => {
        mutated.push(update);
        update.fill(0);
        throw new Error('listener failure');
      });
      left.onUpdate((update, origin) => received.push({ update, origin }));
      right.onUpdate((update, origin) => remote.push({ update, origin }));

      expect(() => left.editCell(0, 20, 0, 'observed')).not.toThrow();
      expect(received).toHaveLength(1);
      expect(received[0].origin).toBe('local');
      expect(mutated[0]).not.toBe(received[0].update);
      expect([...mutated[0]].every((byte) => byte === 0)).toBe(true);
      expect([...received[0].update].some((byte) => byte !== 0)).toBe(true);

      right.applyUpdate(received[0].update);
      expect(remote).toHaveLength(1);
      expect(remote[0].origin).toBe('remote');
      expect(remote[0].update).not.toBe(received[0].update);
      expect([...remote[0].update]).toEqual([...received[0].update]);
      expect(right.cell(0, 20, 0).input).toBe('observed');
    } finally {
      left.dispose();
      right.dispose();
    }
  });

  // the wasm event arrives as `[origin, ...update]`, so a listener that reaches
  // for `.buffer` must not find the origin tag riding along — and the shape must
  // not depend on how many other listeners happen to be subscribed.
  it('delivers exact update buffers regardless of subscriber count', () => {
    for (const extraSubscriber of [false, true]) {
      const source = collaborative(extraSubscriber ? 2102 : 2101);
      const peer = collaborative(extraSubscriber ? 2104 : 2103);
      const received: Uint8Array[] = [];
      try {
        source.onUpdate((update, origin) => {
          if (origin === 'local') received.push(update);
        });
        if (extraSubscriber) source.onUpdate(() => {});

        source.editCell(0, 22, 0, 'exact buffer');
        expect(received).toHaveLength(1);
        const update = received[0];
        expect(update.byteOffset).toBe(0);
        expect(update.buffer.byteLength).toBe(update.byteLength);

        peer.applyUpdate(new Uint8Array(update.buffer));
        expect(peer.cell(0, 22, 0).input).toBe('exact buffer');
      } finally {
        source.dispose();
        peer.dispose();
      }
    }
  });

  it('unsubscribes idempotently', () => {
    const handle = collaborative(3001);
    let calls = 0;
    try {
      const unsubscribe = handle.onUpdate(() => {
        calls += 1;
      });
      unsubscribe();
      unsubscribe();
      handle.editCell(0, 21, 0, 'not observed');
      expect(calls).toBe(0);
    } finally {
      handle.dispose();
    }
  });

  it('keeps duplicate listener registrations independent', () => {
    const handle = collaborative(3002);
    let calls = 0;
    const listener = () => {
      calls += 1;
    };
    try {
      const unsubscribeFirst = handle.onUpdate(listener);
      const unsubscribeSecond = handle.onUpdate(listener);
      unsubscribeFirst();
      handle.editCell(0, 21, 1, 'observed once');
      expect(calls).toBe(1);
      unsubscribeSecond();
    } finally {
      handle.dispose();
    }
  });

  it('exposes safe polling on the generated API', () => {
    const doc = XlsxDocument.openCollaborative(sampleBytes(), 3003);
    try {
      doc.startUpdateObservation();
      doc.editCellJson(JSON.stringify({ sheet: 0, row: 21, col: 2, input: 'polled' }));

      const event = doc.drainUpdateEvent();
      expect(event[0]).toBe(0);
      expect(event.byteLength).toBeGreaterThan(1);
      expect(doc.cellJson(JSON.stringify({ sheet: 0, row: 21, col: 2 }))).toContain('polled');
      expect(doc.drainUpdateEvent()).toEqual(new Uint8Array());
      doc.clearUpdateObservation();
    } finally {
      doc.free();
    }
  });

  it('allows reentrant reads and edits after the raw wasm call returns', () => {
    const handle = collaborative(4001);
    const reads: string[] = [];
    let events = 0;
    try {
      handle.onUpdate(() => {
        events += 1;
        reads.push(handle.cell(0, 22, 0).input);
        if (events === 1) handle.editCell(0, 22, 1, 'nested');
      });
      handle.editCell(0, 22, 0, 'outer');

      expect(events).toBe(2);
      expect(reads).toEqual(['outer', 'outer']);
      expect(handle.cell(0, 22, 1).input).toBe('nested');
    } finally {
      handle.dispose();
    }
  });

  it('ignores duplicate updates and rolls back invalid updates', () => {
    const source = collaborative(5001);
    const target = collaborative(5002);
    let remoteEvents = 0;
    try {
      target.onUpdate((_update, origin) => {
        if (origin === 'remote') remoteEvents += 1;
      });
      source.editCell(0, 23, 0, 'once');
      const update = source.encodeStateAsUpdate(target.encodeStateVector());
      const applied = target.applyUpdate(update);
      expect(applied.applied).toBe(true);
      expect(applied.sheetInfo.sheetNames).toEqual(['Budget', 'Summary', 'Styled']);
      expect(target.applyUpdate(update).applied).toBe(false);
      expect(remoteEvents).toBe(1);

      const state = target.encodeStateAsUpdate();
      const cell = target.cell(0, 23, 0);
      expect(() => target.applyUpdate(Uint8Array.of(0xff))).toThrow('invalid Yrs v1 update');
      expect([...target.encodeStateAsUpdate()]).toEqual([...state]);
      expect(target.cell(0, 23, 0)).toEqual(cell);
      expect(remoteEvents).toBe(1);
    } finally {
      source.dispose();
      target.dispose();
    }
  });

  it('rejects local and remote structural changes without mutation', () => {
    const target = collaborative(6001);
    const structuralSource = openWorkbook(sampleBytes());
    try {
      const state = target.encodeStateAsUpdate();
      const insertRows = [{ type: 'insertRows', sheet: 0, at: 0, count: 1 }];
      expect(() => target.applyOps(insertRows)).toThrow(
        'structural operations are unavailable in collaborative mode'
      );
      expect([...target.encodeStateAsUpdate()]).toEqual([...state]);

      expect(structuralSource.applyOps(insertRows).applied).toBe(true);
      const update = structuralSource.encodeStateAsUpdate(target.encodeStateVector());
      expect(() => target.applyUpdate(update)).toThrow(
        'remote update changes the frozen workbook structure'
      );
      expect([...target.encodeStateAsUpdate()]).toEqual([...state]);
    } finally {
      target.dispose();
      structuralSource.dispose();
    }
  });

  it('converges disjoint collaborative merges through the public handle and reopens the saved XLSX', () => {
    const left = collaborative(6003);
    const right = collaborative(6004);
    const first = { start: { row: 29, col: 0 }, end: { row: 30, col: 1 } };
    const second = { start: { row: 31, col: 3 }, end: { row: 32, col: 4 } };
    try {
      const baseline = left.encodeStateVector();
      expect(left.applyOps([{ type: 'mergeCells', sheet: 0, range: first }]).applied).toBe(true);
      expect(right.applyOps([{ type: 'mergeCells', sheet: 0, range: second }]).applied).toBe(true);
      const leftUpdate = left.encodeStateAsUpdate(baseline);
      const rightUpdate = right.encodeStateAsUpdate(baseline);
      expect(left.applyUpdate(rightUpdate).applied).toBe(true);
      expect(right.applyUpdate(leftUpdate).applied).toBe(true);
      expect(left.mergedRanges(0, 'A30:E33')).toEqual([first, second]);
      expect(right.mergedRanges(0, 'A30:E33')).toEqual([first, second]);

      expect(left.undo().applied).toBe(true);
      expect(right.applyUpdate(left.encodeStateAsUpdate(right.encodeStateVector())).applied).toBe(true);
      expect(left.mergedRanges(0, 'A30:E33')).toEqual([second]);
      expect(right.mergedRanges(0, 'A30:E33')).toEqual([second]);

      const reopened = openWorkbook(right.save());
      try {
        expect(reopened.mergedRanges(0, 'A30:E33')).toEqual([second]);
      } finally {
        reopened.dispose();
      }
    } finally {
      left.dispose();
      right.dispose();
    }
  });

  it('projects overlapping concurrent merges the same way at both public handles', () => {
    const left = collaborative(6005);
    const right = collaborative(6006);
    const first = { start: { row: 39, col: 0 }, end: { row: 41, col: 2 } };
    const second = { start: { row: 40, col: 1 }, end: { row: 42, col: 3 } };
    try {
      const baseline = left.encodeStateVector();
      left.applyOps([{ type: 'mergeCells', sheet: 0, range: first }]);
      right.applyOps([{ type: 'mergeCells', sheet: 0, range: second }]);
      const leftUpdate = left.encodeStateAsUpdate(baseline);
      const rightUpdate = right.encodeStateAsUpdate(baseline);
      right.applyUpdate(leftUpdate);
      left.applyUpdate(rightUpdate);
      const projected = left.mergedRanges(0, 'A40:D43');
      expect(right.mergedRanges(0, 'A40:D43')).toEqual(projected);
      expect(projected).toHaveLength(1);
      expect([first, second]).toContainEqual(projected[0]);
    } finally {
      left.dispose();
      right.dispose();
    }
  });

  it('authors disjoint hyperlinks through the public handle, converges, undoes, and reopens XLSX', () => {
    const left = collaborative(6011);
    const right = collaborative(6012);
    const first = { range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }, location: 'Empty!A1' };
    const second = { range: { start: { row: 1, col: 1 }, end: { row: 1, col: 1 } }, location: 'Data!B1' };
    const viewport = { x: 0, y: 0, width: 300, height: 100 };
    try {
      const baseline = left.encodeStateVector();
      expect(left.setHyperlink(0, first).applied).toBe(true);
      expect(right.setHyperlink(0, second).applied).toBe(true);
      const leftUpdate = left.encodeStateAsUpdate(baseline);
      const rightUpdate = right.encodeStateAsUpdate(baseline);
      expect(left.applyUpdate(rightUpdate).applied).toBe(true);
      expect(right.applyUpdate(leftUpdate).applied).toBe(true);
      expect(left.displayList(viewport).hyperlinks?.map(link => link.location)).toEqual(['Empty!A1', 'Data!B1']);
      expect(right.displayList(viewport).hyperlinks).toEqual(left.displayList(viewport).hyperlinks);

      expect(left.undo().applied).toBe(true);
      expect(right.applyUpdate(left.encodeStateAsUpdate(right.encodeStateVector())).applied).toBe(true);
      expect(right.displayList(viewport).hyperlinks?.map(link => link.location)).toEqual(['Data!B1']);
      const reopened = openWorkbook(right.save());
      try {
        expect(reopened.displayList(viewport).hyperlinks?.map(link => link.location)).toEqual(['Data!B1']);
      } finally {
        reopened.dispose();
      }
      expect(right.removeHyperlink(0, second.range).applied).toBe(true);
      expect(left.applyUpdate(right.encodeStateAsUpdate(left.encodeStateVector())).applied).toBe(true);
      expect(left.displayList(viewport).hyperlinks ?? []).toEqual([]);
    } finally {
      left.dispose();
      right.dispose();
    }
  });

  it('drags a chart in a collaborative session and converges the peer', () => {
    const source = collaborativeCharts(6101);
    const target = collaborativeCharts(6102);
    const viewport = { x: 0, y: 0, width: 800, height: 800 };
    try {
      const before = (source.displayList(viewport).charts ?? [])[0];
      expect(source.moveChart(0, before.id, 24, 12).applied).toBe(true);

      const moved = (source.displayList(viewport).charts ?? [])[0];
      expect(moved.rect.x).toBeCloseTo(before.rect.x + 24, 1);
      expect(moved.rect.y).toBeCloseTo(before.rect.y + 12, 1);

      const update = source.encodeStateAsUpdate(target.encodeStateVector());
      expect(target.applyUpdate(update).applied).toBe(true);
      expect((target.displayList(viewport).charts ?? [])[0].rect).toEqual(moved.rect);
    } finally {
      source.dispose();
      target.dispose();
    }
  });

  it('supports reentrant disposal and rejects all later operations consistently', () => {
    const handle = collaborative(7001);
    let calls = 0;
    const unsubscribe = handle.onUpdate(() => {
      calls += 1;
      handle.dispose();
    });

    expect(() => handle.editCell(0, 24, 0, 'dispose')).not.toThrow();
    expect(calls).toBe(1);
    expect(() => handle.dispose()).not.toThrow();
    expect(() => handle.dispose()).not.toThrow();
    expect(() => unsubscribe()).not.toThrow();
    expect(() => handle.clientId).toThrow('workbook handle is disposed');
    expect(() => handle.sheetInfo()).toThrow('workbook handle is disposed');
    expect(() => handle.encodeStateVector()).toThrow('workbook handle is disposed');
    expect(() => handle.editCell(0, 0, 0, 'later')).toThrow('workbook handle is disposed');
    expect(() => handle.onUpdate(() => {})).toThrow('workbook handle is disposed');
  });
});
