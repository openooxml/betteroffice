import { expect, mock, test } from 'bun:test';
import { createEngineLoader, type EngineModules } from '../src/engine-loader';
import { initialize } from '../src/prototype';
import { tableCells } from '../src/docx-edits';

const xlsx = {} as EngineModules['xlsx'];
const pptx = {} as EngineModules['pptx'];
const docx = {} as EngineModules['docx'];
const wasm = { docx: async () => docx, xlsx: async () => xlsx, pptx: async () => pptx };

test('engine loader falls back per format and warns once across failures', async () => {
  const warn = mock((_warning: Error) => {});
  const failed = mock(async (): Promise<never> => { throw new Error('no native prebuilt'); });
  const loader = createEngineLoader({ docx: failed, xlsx: failed, pptx: failed }, wasm, warn);
  expect(await loader.load('xlsx')).toEqual({ module: xlsx, native: false });
  expect(await loader.load('pptx')).toEqual({ module: pptx, native: false });
  expect(await loader.load('xlsx')).toBe(loader.get('xlsx'));
  expect(failed).toHaveBeenCalledTimes(2);
  expect(warn).toHaveBeenCalledTimes(1);
  const warning = warn.mock.calls[0]![0];
  expect(warning).toMatchObject({ code: 'BETTEROFFICE_AGENTS_NATIVE_UNAVAILABLE' });
  expect(warning.message).toContain('xlsx, pptx');
  expect(warning.message).toContain('no native prebuilt');
  expect(warning.message).toContain('@napi-rs/canvas');
});

test('engine loader uses native lazily and caches the chosen module without warning', async () => {
  const module = {} as EngineModules['xlsx'];
  const native = { docx: mock(wasm.docx), xlsx: mock(async () => module), pptx: mock(wasm.pptx) };
  const fallback = { docx: mock(wasm.docx), xlsx: mock(wasm.xlsx), pptx: mock(wasm.pptx) };
  const warn = mock((_warning: Error) => {});
  const loader = createEngineLoader(native, fallback, warn);
  const [first, second] = await Promise.all([loader.load('xlsx'), loader.load('xlsx')]);
  expect(first).toEqual({ module, native: true });
  expect(second).toBe(first);
  expect(loader.get('xlsx')).toBe(first);
  expect(native.xlsx).toHaveBeenCalledTimes(1);
  expect(native.docx).not.toHaveBeenCalled();
  expect(native.pptx).not.toHaveBeenCalled();
  expect(fallback.xlsx).not.toHaveBeenCalled();
  expect(warn).not.toHaveBeenCalled();
});

test('engine loader keeps the first native failure reason', async () => {
  const warn = mock((_warning: Error) => {});
  const loader = createEngineLoader({ ...wasm, xlsx: async () => { throw new Error('first failure'); }, pptx: async () => { throw new Error('second failure'); } }, wasm, warn);
  await loader.load('xlsx');
  await loader.load('pptx');
  expect(warn).toHaveBeenCalledTimes(1);
  expect(warn.mock.calls[0]![0].message).toContain('first failure');
  expect(warn.mock.calls[0]![0].message).not.toContain('second failure');
});

test('engine loader exposes only an already-loaded module to synchronous helpers', async () => {
  const loader = createEngineLoader(undefined, wasm, () => {});
  expect(() => loader.get('docx')).toThrow('has not loaded');
  const engine = await loader.load('docx');
  expect(loader.get('docx')).toBe(engine);
});

test('native initialization skips wasm bytes and browser initialization preserves them', async () => {
  const init = mock(async (_input?: Uint8Array) => {});
  const bytes = new Uint8Array([1, 2, 3]);
  await initialize('xlsx', init, undefined, true);
  expect(init).toHaveBeenLastCalledWith(undefined);
  await initialize('pptx', init, bytes, true);
  expect(init).toHaveBeenLastCalledWith(undefined);
  await initialize('xlsx', init, bytes);
  expect(init).toHaveBeenLastCalledWith(bytes);
});

test('attached DOCX table reads use payloads without an engine', () => {
  const payload = { grid: [1, 1, 1], rows: [
    { trPr: { _originalFormatting: { gridBefore: 1, gridAfter: 1 } }, cells: [{ story: 'a', tcPr: { rowspan: 2, colspan: 2 } }, { story: 'b' }] },
    { trPr: { gridAfter: 2 }, cells: [{ story: 'c' }, { story: 'd' }] },
  ] };
  const cells = tableCells('body', [{ kind: 'embed', embedKind: 'table', payload }] as never);
  expect([...cells.keys()]).toEqual(['a', 'b', 'c', 'd']);
  expect(cells.get('a')).toMatchObject({ columns: 6, rows: 2, anchor: { row: 0, column: 1, rowspan: 2, colspan: 2, story: 'a' } });
  expect(cells.get('d')).toMatchObject({ columns: 6, anchor: { row: 1, column: 3, rowspan: 1, colspan: 1, story: 'd' } });
  expect([...cells.get('a')!.mergedRows]).toEqual([0, 1]);
});
