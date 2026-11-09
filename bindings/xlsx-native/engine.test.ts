import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { inflateSync } from 'node:zlib';
import { beforeAll, describe, expect, spyOn, test } from 'bun:test';
import JSZip from 'jszip';
import initWasm, { XlsxDocument as WasmDocument, XlsxSnapshotBuilder as WasmSnapshotBuilder, decodeFontFace as wasmDecodeFontFace, displayListCellLimit as wasmDisplayListCellLimit } from '../../packages/xlsx/src/wasm/generated/xlsx_wasm.js';
import { expectFontFace, fontCases, fontErrorCases } from '../font-fixtures';
import { cappedExtent } from '../../packages/xlsx/src/wasm/pngResolve';

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
  recalculateJson(): string;
  encodeStateVector(): Uint8Array;
  encodeStateAsUpdate(): Uint8Array;
  encodeDiff(stateVector: Uint8Array): Uint8Array;
  recoveryCheckpoint(throughSeq: number): Uint8Array;
  renderPngTarget(targetJson: string, optionsJson: string): Uint8Array;
  rangePngDisplayListJson(args: string): string;
  startUpdateObservation(): void;
  drainUpdateEvent(): Uint8Array;
  undoJson(): string;
  readCellsJson(request: string): string;
  calculationStatusJson(): string;
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

function pngPixels(bytes: Uint8Array): Uint8Array {
  const png = Buffer.from(bytes);
  expect([...png.subarray(0, 8)]).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
  expect([png[24], png[25], png[28]]).toEqual([8, 6, 0]);
  const stride = png.readUInt32BE(16) * 4;
  const height = png.readUInt32BE(20);
  const chunks: Buffer[] = [];
  for (let offset = 8; offset < png.length;) {
    const size = png.readUInt32BE(offset);
    if (png.toString('ascii', offset + 4, offset + 8) === 'IDAT') {
      chunks.push(png.subarray(offset + 8, offset + 8 + size));
    }
    offset += size + 12;
  }
  const filtered = inflateSync(Buffer.concat(chunks));
  expect(filtered.length).toBe(height * (stride + 1));
  const pixels = new Uint8Array(height * stride);
  for (let row = 0; row < height; row++) {
    const filter = filtered[row * (stride + 1)];
    expect(filter).toBeLessThanOrEqual(4);
    for (let col = 0; col < stride; col++) {
      const index = row * stride + col;
      const left = col >= 4 ? pixels[index - 4] : 0;
      const up = row > 0 ? pixels[index - stride] : 0;
      const corner = row > 0 && col >= 4 ? pixels[index - stride - 4] : 0;
      const estimate = left + up - corner;
      const distances = [Math.abs(estimate - left), Math.abs(estimate - up), Math.abs(estimate - corner)];
      const paeth = distances[0] <= distances[1] && distances[0] <= distances[2]
        ? left : distances[1] <= distances[2] ? up : corner;
      const prediction = [0, left, up, Math.floor((left + up) / 2), paeth][filter];
      pixels[index] = (filtered[row * (stride + 1) + col + 1] + prediction) & 255;
    }
  }
  return pixels;
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
  describe('renderPng fonts option', () => {
    const face = (name: string) => new Uint8Array(readFileSync(new URL(`../../crates/betteroffice-fonts/tests/fixtures/Synthetic-${name}.ttf`, import.meta.url)));
    const regular = face('Regular');
    const bold = face('Bold');
    const view = { sheet: 0, viewport: { x: 0, y: 0, width: 300, height: 80 } };

    async function openFixture() {
      const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
      const rels = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
      const zip = new JSZip();
      zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
      zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rels}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
      zip.file('xl/workbook.xml', `<workbook xmlns="${main}" xmlns:r="${rels}"><sheets><sheet name="Fonts" sheetId="1" r:id="rId1"/></sheets></workbook>`);
      zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rels}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${rels}/styles" Target="styles.xml"/></Relationships>`);
      zip.file('xl/styles.xml', `<styleSheet xmlns="${main}"><fonts count="1"><font><sz val="24"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs></styleSheet>`);
      zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${main}"><cols><col min="1" max="1" width="40" customWidth="1"/></cols><sheetData><row r="1" ht="40" customHeight="1"><c r="A1" t="inlineStr"><is><t>ABCABC</t></is></c></row></sheetData></worksheet>`);
      const entry = await import('./dist/index.js') as typeof import('../../packages/xlsx/src/index');
      await entry.initWasm();
      return entry.openWorkbook(await zip.generateAsync({ type: 'uint8array' }));
    }

    test('paints with the supplied face, for that export only', async () => {
      const workbook = await openFixture();
      try {
        const baseline = Buffer.from(await workbook.renderPng(view));
        const withRegular = Buffer.from(await workbook.renderPng(view, { fonts: [{ family: 'Calibri', bytes: regular }] }));
        const withBold = Buffer.from(await workbook.renderPng(view, { fonts: [{ family: ' calibri ', bytes: bold }] }));
        expect(withRegular.equals(baseline)).toBe(false);
        expect(withBold.equals(withRegular)).toBe(false);
        expect(Buffer.from(await workbook.renderPng(view)).equals(baseline)).toBe(true);
      } finally {
        workbook.dispose();
      }
    });

    test('a face for a family the sheet does not use changes nothing and is never loaded', async () => {
      const workbook = await openFixture();
      try {
        const baseline = Buffer.from(await workbook.renderPng(view));
        const rendered = await workbook.renderPng(view, { fonts: [{ family: 'Not Used', bytes: new Uint8Array([1, 2, 3]) }] });
        expect(Buffer.from(rendered).equals(baseline)).toBe(true);
      } finally {
        workbook.dispose();
      }
    });

    test('a supplied face that cannot paint rejects with font-unavailable', async () => {
      const workbook = await openFixture();
      try {
        const error = await workbook.renderPng(view, { fonts: [{ family: 'Calibri', bytes: new Uint8Array([1, 2, 3]) }] }).catch((failure: unknown) => failure);
        expect(error).toBeInstanceOf(Error);
        expect((error as { code?: string }).code).toBe('font-unavailable');
        expect((error as Error).message).toContain('Calibri');
      } finally {
        workbook.dispose();
      }
    });
  });

  test('the package renders installed CJK and Devanagari faces without fetching', async () => {
    const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const rels = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
    zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rels}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
    zip.file('xl/workbook.xml', `<workbook xmlns="${main}" xmlns:r="${rels}"><sheets><sheet name="Fonts" sheetId="1" r:id="rId1"/></sheets></workbook>`);
    zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rels}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${rels}/styles" Target="styles.xml"/></Relationships>`);
    zip.file('xl/styles.xml', `<styleSheet xmlns="${main}"><fonts count="1"><font><sz val="24"/><name val="Calibri"/><color rgb="FF000000"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf fontId="0" fillId="0" borderId="0" numFmtId="0" xfId="0"/></cellXfs></styleSheet>`);
    zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${main}"><cols><col min="1" max="1" width="40" customWidth="1"/></cols><sheetData><row r="1" ht="48" customHeight="1"><c r="A1" t="inlineStr"><is><t>漢字テスト</t></is></c></row><row r="2" ht="48" customHeight="1"><c r="A2" t="inlineStr"><is><t>नमस्ते</t></is></c></row></sheetData></worksheet>`);
    const source = await zip.generateAsync({ type: 'uint8array' });
    const view = { x: 0, y: 0, width: 300, height: 150 };
    const ranges = ['A1', 'A2'];
    const fetch = spyOn(globalThis, 'fetch').mockImplementation(() => {
      throw new Error('PNG font loading must not fetch');
    });
    try {
      const script = join(mkdtempSync(join(tmpdir(), 'xlsx-native-fonts-')), 'baseline.mjs');
      writeFileSync(script, `
        import { createRequire } from 'node:module';
        import { readFileSync } from 'node:fs';
        globalThis.fetch = () => { throw new Error('PNG font loading must not fetch'); };
        const { XlsxDocument } = createRequire(import.meta.url)(${JSON.stringify(fileURLToPath(new URL('./index.js', import.meta.url)))});
        const doc = XlsxDocument.open(new Uint8Array(readFileSync(0)));
        try {
          process.stdout.write(JSON.stringify({
            png: Buffer.from(doc.renderPngTarget(${JSON.stringify(JSON.stringify({ sheet: 0, viewport: view }))}, '{}')).toString('base64'),
            ranges: ${JSON.stringify(ranges)}.map(range => Buffer.from(doc.renderPngTarget(JSON.stringify({ sheet: 0, range }), JSON.stringify({ scale: 2 }))).toString('base64')),
          }));
        } finally { doc.free(); }
      `);
      const baseline = spawnSync(process.execPath, [script], { input: source, encoding: 'utf8' });
      expect(baseline.status, baseline.stderr).toBe(0);
      const withoutPacks = JSON.parse(baseline.stdout) as { png: string; ranges: string[] };
      const entry = await import('./dist/index.js') as typeof import('../../packages/xlsx/src/index');
      await entry.initWasm();
      const workbook = entry.openWorkbook(source);
      try {
        expect(Buffer.from(await workbook.renderPng({ sheet: 0, viewport: view }))).not.toEqual(Buffer.from(withoutPacks.png, 'base64'));
        for (const [row, range] of ranges.entries()) {
          const rendered = await workbook.renderPng({ sheet: 0, range }, { scale: 2 });
          const pixels = pngPixels(rendered);
          expect(Buffer.from(rendered)).not.toEqual(Buffer.from(withoutPacks.ranges[row], 'base64'));
          workbook.editCell(0, row, 0, '\u{10ffff}'.repeat(5));
          const placeholder = pngPixels(await workbook.renderPng({ sheet: 0, range }, { scale: 2 }));
          expect(pixels.length).toBe(placeholder.length);
          let ink = 0;
          let changed = 0;
          for (let i = 0; i < pixels.length; i += 4) {
            if (pixels[i] < 128 && pixels[i + 1] < 128 && pixels[i + 2] < 128 && pixels[i + 3] > 128) ink++;
            if (Math.abs(pixels[i] - placeholder[i]) > 64) changed++;
          }
          expect(ink).toBeGreaterThan(100);
          expect(changed).toBeGreaterThan(100);
        }
      } finally {
        workbook.dispose();
      }
      expect(fetch).not.toHaveBeenCalled();
    } finally {
      fetch.mockRestore();
    }
  }, 30_000);

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
        doc.renderPngTarget(JSON.stringify({ sheet: 0, viewport: JSON.parse(viewport) }), '{}'),
        doc.renderPngTarget(JSON.stringify({ sheet: 0, range: 'A1:B2' }), '{}'),
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

  test('recalculation reports and saves the same on both engines', () => {
    const wasmDoc = WasmDocument.openWithCalculationJson(new Uint8Array(fixture), context) as EngineDocument;
    const nativeDoc = native.XlsxDocument.openWithCalculationJson(new Uint8Array(fixture), context) as EngineDocument;
    try {
      expect(nativeDoc.recalculateJson()).toBe(wasmDoc.recalculateJson());
      expect(new Uint8Array(nativeDoc.saveBytes())).toEqual(new Uint8Array(wasmDoc.saveBytes()));
    } finally {
      wasmDoc.free();
      nativeDoc.free();
    }
  });

  test('a flagged workbook with a stale array-formula cache saves the same calculated values on both engines', async () => {
    const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const rels = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
    zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rels}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
    zip.file('xl/workbook.xml', `<workbook xmlns="${main}" xmlns:r="${rels}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><calcPr fullCalcOnLoad="1"/></workbook>`);
    zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rels}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
    zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${main}"><sheetData><row r="2"><c r="C2"><f t="array" ref="C2:C3">1</f><v>1</v></c></row><row r="3"><c r="C3"><v>100</v></c></row></sheetData></worksheet>`);
    const source = await zip.generateAsync({ type: 'uint8array' });
    const opens: Array<[string, (engine: typeof WasmDocument) => EngineDocument]> = [
      ['open', engine => engine.open(source) as EngineDocument],
      ['openWithCalculationJson', engine => engine.openWithCalculationJson(source, context) as EngineDocument],
    ];
    for (const [name, open] of opens) {
      const wasmDoc = open(WasmDocument);
      const nativeDoc = open(native.XlsxDocument);
      try {
        const wasmSaved = new Uint8Array(wasmDoc.saveBytes());
        const nativeSaved = new Uint8Array(nativeDoc.saveBytes());
        expect(nativeSaved, name).toEqual(wasmSaved);
        for (const saved of [wasmSaved, nativeSaved]) {
          const sheet = await (await JSZip.loadAsync(saved)).file('xl/worksheets/sheet1.xml')!.async('string');
          expect(sheet.match(/<c r="C3"[^>]*>.*?<\/c>/)?.[0], name).toContain('<v>1</v>');
        }
      } finally {
        wasmDoc.free();
        nativeDoc.free();
      }
    }
  });

  test('PNG display lists are identical on both engines', () => {
    const charts = readFileSync(new URL('../../packages/xlsx/test-fixtures/charts.xlsx', import.meta.url));
    for (const bytes of [fixture, charts]) {
      const wasmDoc = WasmDocument.openWithCalculationJson(new Uint8Array(bytes), context);
      const nativeDoc = native.XlsxDocument.openWithCalculationJson(new Uint8Array(bytes), context);
      try {
        for (const view of [viewport, JSON.stringify({ x: 12.5, y: 40.25, width: 640, height: 360 })]) {
          expect(nativeDoc.pngDisplayListJson(view)).toBe(wasmDoc.pngDisplayListJson(view));
        }
        for (const args of [{}, { range: 'A1:D8', scale: 2 }]) {
          const json = JSON.stringify(args);
          expect(nativeDoc.rangePngDisplayListJson(json)).toBe(wasmDoc.rangePngDisplayListJson(json));
        }
      } finally {
        wasmDoc.free();
        nativeDoc.free();
      }
    }
  });

  test('renderPngTarget draws both target forms, scales and caps them, and refuses bad input', () => {
    const doc = native.XlsxDocument.openWithCalculationJson(new Uint8Array(fixture), context) as EngineDocument;
    const size = (bytes: Uint8Array) => {
      const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      return [view.getUint32(16), view.getUint32(20)];
    };
    const render = (target: unknown, options: unknown = {}) => doc.renderPngTarget(JSON.stringify(target), JSON.stringify(options));
    const box = { x: 0, y: 0, width: 300, height: 120 };
    try {
      expect(size(render({ sheet: 0, viewport: box }))).toEqual([300, 120]);
      expect(size(render({ sheet: 0, viewport: box }, { scale: 2 }))).toEqual([600, 240]);
      const [width, height] = size(render({ sheet: 0, range: 'A1:D8' }));
      const [doubleWidth, doubleHeight] = size(render({ sheet: 0, range: 'a1:d8' }, { scale: 2 }));
      expect(doubleWidth).toBeGreaterThanOrEqual(2 * width - 2);
      expect(doubleWidth).toBeLessThanOrEqual(2 * width);
      expect(doubleHeight).toBeGreaterThanOrEqual(2 * height - 2);
      expect(doubleHeight).toBeLessThanOrEqual(2 * height);
      for (const target of [{ sheet: 0, viewport: box }, { sheet: 0, range: 'A1:D8' }]) {
        for (const scale of [1, 1.3, 2.7]) {
          const [cappedWidth, cappedHeight] = size(render(target, { scale, maxWidth: 37, maxHeight: 23 }));
          expect(cappedWidth).toBeLessThanOrEqual(37);
          expect(cappedWidth).toBeGreaterThanOrEqual(36);
          expect(cappedHeight).toBeLessThanOrEqual(23);
          expect(cappedHeight).toBeGreaterThanOrEqual(22);
        }
      }
      const used = JSON.parse(doc.rangePngDisplayListJson('{"scale":2}')) as { width: number; height: number };
      expect(size(render({ sheet: 0, range: 'used' }, { scale: 2 }))).toEqual([Math.ceil(used.width), Math.ceil(used.height)]);
      expect(() => render({ sheet: 99, viewport: box })).toThrow('sheet index 99 out of range');
      expect(() => render({ sheet: 0, range: 'not a range' })).toThrow('bad range: malformed address');
      expect(render({ sheet: 0, range: 'a1:b2' })).toEqual(render({ sheet: 0, range: 'A1:B2' }));
      expect(() => doc.renderPngTarget('{"sheet":0,"viewport":{"x":0,"y":0,"width":1e999,"height":1}}', '{}')).toThrow();
      expect(() => render({ sheet: 0, viewport: box }, { scale: 1, dpi: 2 })).toThrow('bad options');
      expect(() => render({ sheet: 0, viewport: box, extra: true })).toThrow('target must be');
      expect(() => render({ sheet: 0, viewport: { ...box, z: 1 } })).toThrow('target must be');
    } finally {
      doc.free();
    }
  });

  test('the canvas path caps a viewport to the same PNG size as the native engine', () => {
    const doc = native.XlsxDocument.openWithCalculationJson(new Uint8Array(fixture), context) as EngineDocument;
    try {
      for (const [extent, cap, scale] of [
        [300, 37, 1], [300, 37, 1.3], [300, 37, 2.7], [300, 100, 3], [300, 391, 1.3], [300, 390, 1.3],
        [300, 810, 2.7], [300, 811, 2.7], [120.5, 23, 1.3], [64.75, 97, 0.75], [300, 1, 2.7],
      ] as const) {
        const bytes = doc.renderPngTarget(JSON.stringify({ sheet: 0, viewport: { x: 0, y: 0, width: extent, height: 40 } }),
          JSON.stringify({ scale, maxWidth: cap }));
        const nativeWidth = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(16);
        const canvasWidth = Math.max(1, Math.ceil(Math.fround(cappedExtent(extent, cap, scale) * Math.fround(scale))));
        expect(canvasWidth, `extent ${extent}, cap ${cap}, scale ${scale}`).toBe(nativeWidth);
      }
    } finally {
      doc.free();
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

  test('both engines keep and report a call to an unknown function by name', async () => {
    const zip = await JSZip.loadAsync(fixture);
    zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">'
      + '<c r="A1"><f>IFERROR(REPRO_UNKNOWN_FUNCTION(1),0)</f><v>42</v></c>'
      + '<c r="B1"><f>IFERROR(TYPO(1),0)</f></c>'
      + '<c r="C1"><f>TYPO(1)</f></c></row></sheetData></worksheet>');
    const bytes = await zip.generateAsync({ type: 'uint8array' });
    const read = (doc: EngineDocument) => [
      doc.calculationStatusJson(),
      withoutVersion(doc.readCellsJson(JSON.stringify({
        ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1:C1' } }],
      }))),
    ];
    const wasmDoc = WasmDocument.openWithCalculationJson(bytes.slice(), context) as EngineDocument;
    const nativeDoc = native.XlsxDocument.openWithCalculationJson(bytes.slice(), context) as EngineDocument;
    try {
      const [status, cells] = read(wasmDoc);
      expect(JSON.parse(status).unevaluatedCells).toEqual([
        { address: 'A1', functions: ['REPRO_UNKNOWN_FUNCTION'] },
        { address: 'B1', functions: ['TYPO'] },
        { address: 'C1', functions: ['TYPO'] },
      ]);
      expect(read(nativeDoc)).toEqual([status, cells]);
    } finally {
      wasmDoc.free();
      nativeDoc.free();
    }
  });

  test('both engines keep and report a formula no offline engine can compute', async () => {
    const zip = await JSZip.loadAsync(fixture);
    zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">'
      + '<c r="A1" t="str"><f>IFERROR(_xlfn.WEBSERVICE("http://x"),"fallback")</f><v>live</v></c>'
      + '<c r="B1"><f>LEN(A1)*2</f><v>9</v></c></row></sheetData></worksheet>');
    const bytes = await zip.generateAsync({ type: 'uint8array' });
    const read = (doc: EngineDocument) => [
      doc.calculationStatusJson(),
      withoutVersion(doc.readCellsJson(JSON.stringify({
        ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1:B1' } }],
      }))),
    ];
    const wasmDoc = WasmDocument.openWithCalculationJson(bytes.slice(), context) as EngineDocument;
    const nativeDoc = native.XlsxDocument.openWithCalculationJson(bytes.slice(), context) as EngineDocument;
    try {
      const [status, cells] = read(wasmDoc);
      expect(JSON.parse(status).unevaluatedCells).toEqual([
        { address: 'A1', functions: ['WEBSERVICE'] },
        { address: 'B1', functions: ['WEBSERVICE'] },
      ]);
      expect(cells).toContain('"live"');
      expect(read(nativeDoc)).toEqual([status, cells]);
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

describe('control characters in formulas through applyOpsJson', () => {
  test('both engines write a new one with CHAR and carry an opened one through paste, shift and move', async () => {
    const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const rels = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
    zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rels}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
    zip.file('xl/workbook.xml', `<workbook xmlns="${main}" xmlns:r="${rels}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`);
    zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rels}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
    zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${main}"><sheetData><row r="1"><c r="D1" t="str"><f>"a&#1;b"</f><v>x</v></c></row></sheetData></worksheet>`);
    const source = await zip.generateAsync({ type: 'uint8array' });
    const refusal = 'U+0000, which a workbook file cannot store and CHAR cannot produce';
    const refused = (run: () => unknown) => {
      try {
        run();
      } catch (error) {
        return String(error instanceof Error ? error.message : error);
      }
      return '';
    };
    const added = JSON.stringify({ ops: [{ type: 'setCell', sheet: 0, at: { row: 0, col: 7 }, cell: { formula: '"new\u0001"' } }] });
    const shifted = JSON.stringify({
      ops: [
        { type: 'insertRows', sheet: 0, at: 0, count: 1 },
        { type: 'setCell', sheet: 0, at: { row: 1, col: 3 }, cell: { formula: '"a\u0001b"' } },
      ],
    });
    const saved: Uint8Array[] = [];
    for (const engine of [WasmDocument, native.XlsxDocument]) {
      const doc = engine.openWithCalculationJson(new Uint8Array(source), context);
      try {
        expect(JSON.parse(doc.applyOpsJson(added)).applied).toBe(true);
        doc.editCellJson(JSON.stringify({ sheet: 0, row: 0, col: 8, input: '="typed\u0001"' }));
        expect(refused(() => doc.editCellJson(JSON.stringify({ sheet: 0, row: 0, col: 10, input: '="nul\u0000"' })))).toContain(refusal);
        doc.editCellsJson(JSON.stringify({ sheet: 0, edits: [{ row: 0, col: 9, input: '="a\u0001b"', copyFrom: { row: 0, col: 3 } }] }));
        expect(JSON.parse(doc.applyOpsJson(shifted)).applied).toBe(true);
        doc.moveRangeJson(JSON.stringify({ sheet: 0, source: 'D2', destination: 'G2' }));
        saved.push(new Uint8Array(doc.saveBytes()));
      } finally {
        doc.free();
      }
    }
    expect(saved[1]).toEqual(saved[0]);
    const sheet = await (await JSZip.loadAsync(saved[0])).file('xl/worksheets/sheet1.xml')!.async('string');
    expect(sheet.match(/a&#1;b/g)?.length).toBe(2);
    expect(sheet).toContain('<f>(&quot;new&quot;&amp;CHAR(1))</f>');
    expect(sheet).toContain('<f>(&quot;typed&quot;&amp;CHAR(1))</f>');
    expect(sheet).not.toContain('nul');
  });
});
