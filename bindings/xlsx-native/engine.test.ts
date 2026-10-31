import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, spyOn, test } from 'bun:test';
import JSZip from 'jszip';
import initWasm, { XlsxDocument as WasmDocument, XlsxSnapshotBuilder as WasmSnapshotBuilder, decodeFontFace as wasmDecodeFontFace, displayListCellLimit as wasmDisplayListCellLimit } from '../../packages/xlsx/src/wasm/generated/xlsx_wasm.js';
import { expectFontFace, fontCases, fontErrorCases } from '../font-fixtures';

const native = require('./index.js');

const fixture = readFileSync(new URL('../../crates/ooxml-opc/tests/fixtures/sample.xlsx', import.meta.url));
const context = JSON.stringify({ nowSerial: 45000.75, randSeed: 42 });
const viewport = JSON.stringify({ x: 0, y: 0, width: 300, height: 120 });
const withoutVersion = (json: string) => json.replace(/"version":"[0-9a-f]{32}-\d+"/g, '"version":""');

beforeAll(async () => {
  await initWasm({
    module_or_path: readFileSync(new URL('../../packages/xlsx/src/wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)),
  });
});

type EngineDocument = {
  sheetInfoJson(): string;
  displayListJson(viewportJson: string): string;
  editCellJson(args: string): string;
  setCalculationContextJson(context: string): void;
  encodeStateVector(): Uint8Array;
  encodeStateAsUpdate(): Uint8Array;
  encodeDiff(stateVector: Uint8Array): Uint8Array;
  recoveryCheckpoint(throughSeq: number): Uint8Array;
  renderPng(viewportJson: string): Uint8Array;
  renderRangePng(args: string): Uint8Array;
  startUpdateObservation(): void;
  drainUpdateEvent(): Uint8Array;
  undoJson(): string;
  readCellsJson(request: string): string;
  exportStructuredJson(options: string): string;
  saveBytes(): Uint8Array;
  free(): void;
};

function expectPlainBytes(bytes: Uint8Array) {
  expect(bytes).toBeInstanceOf(Uint8Array);
  expect(Buffer.isBuffer(bytes)).toBe(false);
  expect(Object.getPrototypeOf(bytes)).toBe(Uint8Array.prototype);
  expect(bytes.byteOffset).toBe(0);
  expect(bytes.buffer.byteLength).toBe(bytes.byteLength);
}

function caughtValue(call: () => unknown): unknown {
  try {
    call();
  } catch (error) {
    return error;
  }
  throw new Error('Expected the engine call to throw');
}

function expectThrownParity(wasmCall: () => unknown, nativeCall: () => unknown): unknown {
  const expected = caughtValue(wasmCall);
  const actual = caughtValue(nativeCall);
  expect(typeof actual).toBe(typeof expected);
  expect(actual instanceof Error).toBe(expected instanceof Error);
  expect(actual instanceof Error ? actual.name : undefined)
    .toBe(expected instanceof Error ? expected.name : undefined);
  expect(actual instanceof Error ? actual.message : String(actual))
    .toBe(expected instanceof Error ? expected.message : String(expected));
  expect(actual instanceof Error && 'code' in actual).toBe(expected instanceof Error && 'code' in expected);
  return actual;
}

function runScript(doc: EngineDocument) {
  const sheetInfo = doc.sheetInfoJson();
  const displayList = doc.displayListJson(viewport);
  doc.startUpdateObservation();
  const edited = doc.editCellJson(JSON.stringify({ sheet: 0, row: 0, col: 0, input: '42' }));
  const drained = doc.drainUpdateEvent();
  const stateVector = doc.encodeStateVector();
  const undone = doc.undoJson();
  const readCells = doc.readCellsJson(JSON.stringify({ ranges: [] }));
  const exported = doc.exportStructuredJson(JSON.stringify({}));
  const saved = doc.saveBytes();
  return { sheetInfo, displayList, edited, drained, stateVector, undone, readCells, exported, saved };
}

describe('xlsx native engine parity', () => {
  test('linked font decoder returns known answers and matching native and wasm errors', () => {
    expect(typeof wasmDecodeFontFace).toBe('function');
    expect(typeof native.decodeFontFace).toBe('function');
    for (const entry of fontCases()) {
      const expected = wasmDecodeFontFace(entry.bytes, entry.codec, entry.dict, entry.maxOutput);
      const actual = native.decodeFontFace(entry.bytes, entry.codec, entry.dict, entry.maxOutput);
      expectFontFace(expected, entry);
      expectFontFace(actual, entry);
      expectPlainBytes(expected.font);
      expectPlainBytes(actual.font);
      expect(actual.font).toEqual(expected.font);
      expect(actual.payload).toEqual(expected.payload);
      expect('payload' in expected).toBe(entry.payload !== undefined);
      if (expected.payload && actual.payload) {
        expectPlainBytes(expected.payload);
        expectPlainBytes(actual.payload);
      }
    }
    for (const entry of fontErrorCases()) {
      const error = expectThrownParity(
        () => wasmDecodeFontFace(entry.bytes, entry.codec, entry.dict, entry.maxOutput),
        () => native.decodeFontFace(entry.bytes, entry.codec, entry.dict, entry.maxOutput),
      );
      expect(error).toBeInstanceOf(Error);
      expect(error).toMatchObject({ name: 'Error', message: entry.message });
    }
  });

  test('native byte results are plain Uint8Arrays and drained slices are independent', () => {
    const doc = native.XlsxDocument.openWithCalculationJson(new Uint8Array(fixture), context) as EngineDocument;
    try {
      doc.startUpdateObservation();
      doc.editCellJson(JSON.stringify({ sheet: 0, row: 0, col: 0, input: '42' }));
      const drained = doc.drainUpdateEvent();
      for (const bytes of [
        drained,
        doc.encodeStateVector(),
        doc.encodeStateAsUpdate(),
        doc.encodeDiff(new Uint8Array([0])),
        doc.recoveryCheckpoint(0),
        doc.renderPng(viewport),
        doc.renderRangePng(JSON.stringify({ range: 'A1:B2' })),
        doc.saveBytes(),
      ]) expectPlainBytes(bytes);
      expect(drained.byteLength).toBeGreaterThan(0);
      const copied = drained.slice();
      const first = drained[0];
      copied[0] ^= 0xff;
      expect(drained[0]).toBe(first);
      expect(copied[0]).not.toBe(drained[0]);
    } finally {
      doc.free();
    }
  });

  test('native saved bytes can be transferred', () => {
    const doc = native.XlsxDocument.openWithCalculationJson(new Uint8Array(fixture), context) as EngineDocument;
    try {
      const bytes = doc.saveBytes();
      const expected = bytes.slice();
      const transferred = structuredClone(bytes, { transfer: [bytes.buffer as ArrayBuffer] });
      expect(bytes.byteLength).toBe(0);
      expect(transferred).toEqual(expected);
      expectPlainBytes(transferred);
    } finally {
      doc.free();
    }
  });

  test('message failures throw the same string on both engines', () => {
    const error = expectThrownParity(
      () => WasmDocument.openCollaborative(new Uint8Array(), 0),
      () => native.XlsxDocument.openCollaborative(new Uint8Array(), 0),
    );
    expect(typeof error).toBe('string');
  });

  test('invalid recovery sequences throw the same TypeError on both engines', () => {
    const error = expectThrownParity(
      () => WasmDocument.openRecoveryCheckpoint(new Uint8Array(), -1),
      () => native.XlsxDocument.openRecoveryCheckpoint(new Uint8Array(), -1),
    );
    expect(error).toBeInstanceOf(TypeError);
  });

  test('invalid recovery checkpoints throw the same Error on both engines', () => {
    const error = expectThrownParity(
      () => WasmDocument.openRecoveryCheckpoint(new Uint8Array(), 0),
      () => native.XlsxDocument.openRecoveryCheckpoint(new Uint8Array(), 0),
    );
    expect(error).toBeInstanceOf(Error);
    expect(error).not.toBeInstanceOf(TypeError);
  });

  test.each(['Date.now', 'getTimezoneOffset'])('a throwing %s fails the edit on both engines and leaves the native document untouched', (source) => {
    const failure = new Error('clock callback failed');
    const request = JSON.stringify({
      ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1' } }],
    });
    for (const Document of [WasmDocument, native.XlsxDocument]) {
      const doc = Document.openWithCalculationJson(new Uint8Array(fixture), context) as EngineDocument;
      try {
        const before = doc.readCellsJson(request);
        const clock = source === 'Date.now'
          ? spyOn(Date, 'now')
          : spyOn(Date.prototype, 'getTimezoneOffset');
        let caught: unknown;
        try {
          clock.mockImplementationOnce(() => { throw failure; });
          caught = caughtValue(() => doc.editCellJson(JSON.stringify({
            sheet: 0, row: 0, col: 0, input: 'changed',
          })));
        } finally {
          clock.mockRestore();
        }
        expect(caught).toBe(failure);
        if (Document !== WasmDocument) expect(doc.readCellsJson(request)).toBe(before);
      } finally {
        if (Document !== WasmDocument) doc.free();
      }
    }
  });

  test.each(['fill', 'detach'])('clock callbacks can %s open inputs without changing results', (operation) => {
    const results = [];
    const request = JSON.stringify({
      ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1:C3' } }],
    });
    for (const Document of [WasmDocument, native.XlsxDocument]) {
      const bytes = new Uint8Array(fixture);
      let mutated = false;
      const clock = spyOn(Date, 'now').mockImplementationOnce(() => {
        if (operation === 'fill') bytes.fill(0);
        else structuredClone(bytes, { transfer: [bytes.buffer] });
        mutated = true;
        return Date.UTC(2026, 0, 1);
      });
      let doc: EngineDocument | undefined;
      try {
        doc = Document.open(bytes) as EngineDocument;
        expect(mutated).toBe(true);
        if (operation === 'fill') expect(bytes.every((byte) => byte === 0)).toBe(true);
        else expect(bytes.byteLength).toBe(0);
        results.push({
          sheetInfo: doc.sheetInfoJson(),
          displayList: doc.displayListJson(viewport),
          cells: JSON.parse(doc.readCellsJson(request)).ranges,
        });
      } finally {
        try { doc?.free(); }
        finally { clock.mockRestore(); }
      }
    }
    expect(results[1]).toEqual(results[0]);
  });

  test.each([-120, 300, -330])('native calculation uses the JS clock and timezone offset %s', (offset) => {
    const ms = Date.UTC(2026, 0, 1, 23, 30);
    const clock = spyOn(Date, 'now').mockReturnValue(ms);
    const timezone = spyOn(Date.prototype, 'getTimezoneOffset').mockReturnValue(offset);
    let seed: EngineDocument | undefined;
    let wasmDoc: EngineDocument | undefined;
    let nativeDoc: EngineDocument | undefined;
    try {
      seed = native.XlsxDocument.openWithCalculationJson(new Uint8Array(fixture), context) as EngineDocument;
      seed.editCellJson(JSON.stringify({ sheet: 0, row: 30, col: 0, input: '=NOW()' }));
      seed.editCellJson(JSON.stringify({ sheet: 0, row: 30, col: 1, input: '=TODAY()' }));
      const bytes = seed.saveBytes();
      wasmDoc = WasmDocument.open(bytes);
      nativeDoc = native.XlsxDocument.open(bytes) as EngineDocument;
      const request = JSON.stringify({
        ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A31:B31' } }],
      });
      const nowSerial = (ms - offset * 60_000) / 86_400_000 + 25_569;
      for (const doc of [wasmDoc, nativeDoc]) {
        expect(JSON.parse(doc.readCellsJson(request)).ranges[0].cells[0].map((cell: { value: unknown }) => cell.value))
          .toEqual([{ kind: 'number', value: nowSerial }, { kind: 'number', value: Math.floor(nowSerial) }]);
        doc.setCalculationContextJson(context);
        doc.editCellJson(JSON.stringify({ sheet: 0, row: 39, col: 25, input: '1' }));
      }
      clock.mockReturnValue(ms + 86_400_000);
      for (const doc of [wasmDoc, nativeDoc]) {
        doc.setCalculationContextJson('null');
        doc.editCellJson(JSON.stringify({ sheet: 0, row: 39, col: 25, input: '2' }));
        expect(JSON.parse(doc.readCellsJson(request)).ranges[0].cells[0].map((cell: { value: unknown }) => cell.value))
          .toEqual([{ kind: 'number', value: nowSerial + 1 }, { kind: 'number', value: Math.floor(nowSerial) + 1 }]);
      }
    } finally {
      seed?.free();
      wasmDoc?.free();
      nativeDoc?.free();
      timezone.mockRestore();
      clock.mockRestore();
    }
  });

  test('native opens report progress through the shared JS hook', async () => {
    const zip = await JSZip.loadAsync(fixture);
    const rows = ['<row r="1"><c r="A1"><v>1</v></c></row>'];
    for (let row = 2; row <= 20_000; row++) {
      rows.push(`<row r="${row}"><c r="A${row}"><f>A${row - 1}+1</f></c></row>`);
    }
    zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows.join('')}</sheetData></worksheet>`);
    const bytes = await zip.generateAsync({ type: 'uint8array' });
    const host = globalThis as typeof globalThis & { __betterofficeEngineProgress?: () => void };
    const previous = host.__betterofficeEngineProgress;
    try {
      for (const Document of [WasmDocument, native.XlsxDocument]) {
        let ticks = 0;
        const input = bytes.slice();
        host.__betterofficeEngineProgress = () => { ticks++; input.fill(0); };
        const doc = Document.open(input);
        try {
          expect(ticks).toBeGreaterThanOrEqual(4);
          expect(input.every((byte) => byte === 0)).toBe(true);
        } finally {
          doc.free();
        }
      }
    } finally {
      if (previous === undefined) delete host.__betterofficeEngineProgress;
      else host.__betterofficeEngineProgress = previous;
    }
  });

  test('wasm and native engines return identical results', () => {
    const wasmDoc = WasmDocument.openWithCalculationJson(new Uint8Array(fixture), context) as EngineDocument;
    const nativeDoc = native.XlsxDocument.openWithCalculationJson(new Uint8Array(fixture), context) as EngineDocument;
    try {
      const wasmResult = runScript(wasmDoc);
      const nativeResult = runScript(nativeDoc);
      expect(nativeResult.sheetInfo).toBe(wasmResult.sheetInfo);
      expect(nativeResult.displayList).toBe(wasmResult.displayList);
      expect(withoutVersion(nativeResult.edited)).toBe(withoutVersion(wasmResult.edited));
      expect(nativeResult.drained[0]).toBe(wasmResult.drained[0]);
      expect(withoutVersion(nativeResult.undone)).toBe(withoutVersion(wasmResult.undone));
      expect(withoutVersion(nativeResult.readCells)).toBe(withoutVersion(wasmResult.readCells));
      expect(withoutVersion(nativeResult.exported)).toBe(withoutVersion(wasmResult.exported));
      expect(new Uint8Array(nativeResult.saved)).toEqual(new Uint8Array(wasmResult.saved));
    } finally {
      wasmDoc.free();
      nativeDoc.free();
    }
  });

  test('collaborative documents with the same client id encode identical updates', () => {
    const wasmDoc = WasmDocument.openCollaborative(new Uint8Array(fixture), 4321) as EngineDocument;
    const nativeDoc = native.XlsxDocument.openCollaborative(new Uint8Array(fixture), 4321) as EngineDocument;
    try {
      const script = (doc: EngineDocument) => {
        doc.startUpdateObservation();
        doc.editCellJson(JSON.stringify({ sheet: 0, row: 0, col: 0, input: '42' }));
        return { drained: doc.drainUpdateEvent(), stateVector: doc.encodeStateVector(), update: doc.encodeStateAsUpdate() };
      };
      const wasmResult = script(wasmDoc);
      const nativeResult = script(nativeDoc);
      expect(new Uint8Array(nativeResult.drained)).toEqual(new Uint8Array(wasmResult.drained));
      expect(new Uint8Array(nativeResult.stateVector)).toEqual(new Uint8Array(wasmResult.stateVector));
      expect(new Uint8Array(nativeResult.update)).toEqual(new Uint8Array(wasmResult.update));
    } finally {
      wasmDoc.free();
      nativeDoc.free();
    }
  });

  test('a drained snapshot ends with undefined on both engines', () => {
    const wasmDoc = WasmDocument.openWithCalculationJson(new Uint8Array(fixture), context) as EngineDocument & {
      beginPeerSnapshot(records: number, bytes: number): void;
      nextPeerSnapshotChunk(): Uint8Array | undefined;
    };
    const nativeDoc = native.XlsxDocument.openWithCalculationJson(new Uint8Array(fixture), context) as typeof wasmDoc;
    try {
      for (const doc of [wasmDoc, nativeDoc]) {
        doc.beginPeerSnapshot(1_000_000, 64 * 1024 * 1024);
        let chunk = doc.nextPeerSnapshotChunk();
        let chunks = 0;
        while (chunk !== undefined) {
          expectPlainBytes(chunk);
          chunks++;
          chunk = doc.nextPeerSnapshotChunk();
        }
        expect(chunks).toBeGreaterThan(0);
        expect(chunk).toBeUndefined();
      }
    } finally {
      wasmDoc.free();
      nativeDoc.free();
    }
  });

  test('transcendental formulas compute bit-identical results on both engines', () => {
    const inputs = ['0', '1E-300', '0.1', '0.5', '1', '2', '3.141592653589793', '6.283185307179586', '17.27', '1E10', '-7.25', '-0.75'];
    const functions = [
      (x: string) => `=SIN(${x})`, (x: string) => `=COS(${x})`, (x: string) => `=TAN(${x})`,
      (x: string) => `=ATAN2(${x},1.7)`, (x: string) => `=ATAN(${x})`, (x: string) => `=POWER(${x},1.37)`,
      (x: string) => `=POWER(1.37,${x})`, (x: string) => `=LN(${x})`, (x: string) => `=LOG10(${x})`,
      (x: string) => `=EXP(${x})`, (x: string) => `=SINH(${x})`, (x: string) => `=(${x})^(1/4)`,
    ];
    const read = (doc: EngineDocument) => {
      functions.forEach((formula, col) => inputs.forEach((input, row) => {
        doc.editCellJson(JSON.stringify({ sheet: 0, row: 200 + row, col, input: formula(input) }));
      }));
      return doc.readCellsJson(JSON.stringify({
        ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: `A201:L${200 + inputs.length}` } }],
      }));
    };
    const wasmDoc = WasmDocument.openWithCalculationJson(new Uint8Array(fixture), context) as EngineDocument;
    const nativeDoc = native.XlsxDocument.openWithCalculationJson(new Uint8Array(fixture), context) as EngineDocument;
    try {
      const wasmCells = read(wasmDoc);
      expect(wasmCells).toContain('"value"');
      expect(withoutVersion(read(nativeDoc))).toBe(withoutVersion(wasmCells));
    } finally {
      wasmDoc.free();
      nativeDoc.free();
    }
  });

  test('free drops the native document and later calls throw', () => {
    const doc = native.XlsxDocument.openWithCalculationJson(new Uint8Array(fixture), context) as EngineDocument;
    doc.free();
    expect(() => doc.sheetInfoJson()).toThrow('XlsxDocument was freed');
    doc.free();
    expect(() => doc.sheetInfoJson()).toThrow('XlsxDocument was freed');
  });

  test('snapshot builder errors have the wasm error shape', () => {
    const wasmBuilder = new WasmSnapshotBuilder();
    const nativeBuilder = new native.XlsxSnapshotBuilder();
    try {
      const error = expectThrownParity(
        () => wasmBuilder.push(new Uint8Array([1, 2, 3])),
        () => nativeBuilder.push(new Uint8Array([1, 2, 3])),
      );
      expect(error).toBeInstanceOf(Error);
    } finally {
      wasmBuilder.free();
      nativeBuilder.free();
    }
  });

  test('free drops the native snapshot builder and later calls throw', () => {
    const builder = new native.XlsxSnapshotBuilder();
    builder.free();
    expect(() => builder.push(new Uint8Array([1]))).toThrow('XlsxSnapshotBuilder was freed');
    expect(() => builder.finish()).toThrow('XlsxSnapshotBuilder was freed');
  });

  test('engineVersion matches the binding version shape', () => {
    expect(typeof native.engineVersion()).toBe('string');
    expect(native.engineVersion().length).toBeGreaterThan(0);
    expect(native.allocationFailure()).toBe(0);
    expect(native.displayListCellLimit()).toBeGreaterThan(0);
    expect(native.XlsxDocument.version()).toBe(WasmDocument.version());
    expect(native.displayListCellLimit()).toBe(wasmDisplayListCellLimit());
  });
});
