import { describe, expect, test } from 'bun:test';
import { parseSurface } from './check-engine-surface';
import {
  compareDeclarations,
  compareArity,
  compareRuntime,
  discover,
  jsArity,
  mergeDeclarations,
  mergeSurfaces,
  normalizeParamName,
  scopeAllowances,
  summaryLine,
  unreached,
  type Allowances,
} from './raw-engine-surface';

const none = (): Allowances => ({ nativeOnly: {}, wasmOnly: {}, internals: { wasm: {}, native: {} } });

class WasmDoc {
  static open(_bytes: Uint8Array) { return new WasmDoc(); }
  free() {}
  [Symbol.dispose]() {}
  sheetCount() { return 1; }
  get title() { return ''; }
}

function wasmModule(overrides: Record<string, unknown> = {}) {
  return { Doc: WasmDoc, exportJson(_bytes: Uint8Array) { return ''; }, ...overrides };
}

function nativeModule(change?: (proto: Record<string | symbol, unknown>, ctor: Record<string, unknown>) => void) {
  class NativeDoc {}
  const proto = NativeDoc.prototype as unknown as Record<string | symbol, unknown>;
  const ctor = NativeDoc as unknown as Record<string, unknown>;
  for (const key of Reflect.ownKeys(WasmDoc.prototype)) {
    if (key !== 'constructor') Object.defineProperty(proto, key, Object.getOwnPropertyDescriptor(WasmDoc.prototype, key)!);
  }
  ctor.open = WasmDoc.open;
  change?.(proto, ctor);
  return { Doc: NativeDoc, exportJson() { return ''; } };
}

const WASM_D_TS = `export class Doc {
  private constructor();
  free(): void;
  [Symbol.dispose](): void;
  static open(bytes: Uint8Array): Doc;
  sheetCount(): number;
  readonly title: string;
}
export function exportJson(bytes: Uint8Array): string;
`;
const NATIVE_D_TS = `export declare class Doc {
  free(): void
  static open(bytes: Uint8Array): Doc
  sheetCount(): number
  get title(): string
}
export declare function exportJson(bytes: Uint8Array): string
`;

describe('raw engine surface discovery', () => {
  test('finds classes, statics, getters and symbol-keyed members without a class list', () => {
    const surface = discover(wasmModule());
    expect(surface.get('Doc')).toBe('class');
    expect(surface.get('exportJson')).toBe('function');
    expect(surface.get('Doc.open')).toBe('method');
    expect(surface.get('Doc.prototype.sheetCount')).toBe('method');
    expect(surface.get('Doc.prototype.title')).toBe('getter');
    expect(surface.get('Doc.prototype.[Symbol.dispose]')).toBe('method');
  });

  test('a name exported by two merged modules must have the same shape', () => {
    const { problems } = mergeSurfaces([
      ['a', discover({ shared() {} })],
      ['b', discover({ shared: 1 })],
    ]);
    expect(problems).toEqual(['shared: function in a, value in b']);
  });
});

describe('raw engine surface compare', () => {
  test('identical surfaces pass and report figures', () => {
    const { problems, stats } = compareRuntime(discover(wasmModule()), discover(nativeModule()), none());
    expect(problems).toEqual([]);
    expect(stats).toEqual({ classes: 1, compared: 6, equal: 6, extras: 0, pending: 0 });
    expect(summaryLine('xlsx', stats)).toBe('xlsx: classes 1, members compared 6, equal 6, declared extras 0, pending 0');
  });

  test('a method missing on native is named', () => {
    const native = discover(nativeModule((proto) => { delete proto.sheetCount; }));
    expect(compareRuntime(discover(wasmModule()), native, none()).problems)
      .toEqual(['Doc.prototype.sheetCount: missing on native (method); add it to the native binding or declare it wasm-only']);
  });

  test('a method missing on wasm is named', () => {
    const native = discover(nativeModule((proto) => { proto.renderPngTarget = () => new Uint8Array(); }));
    expect(compareRuntime(discover(wasmModule()), native, none()).problems)
      .toEqual(['Doc.prototype.renderPngTarget: only on native (method); declare it in the native extras or add it to wasm']);
  });

  test('a symbol-keyed member missing on native is named', () => {
    const native = discover(nativeModule((proto) => { delete proto[Symbol.dispose]; }));
    expect(compareRuntime(discover(wasmModule()), native, none()).problems[0]).toStartWith('Doc.prototype.[Symbol.dispose]: missing on native');
  });

  test('static versus instance is a difference', () => {
    const native = discover(nativeModule((proto, ctor) => { proto.open = ctor.open; delete ctor.open; }));
    const problems = compareRuntime(discover(wasmModule()), native, none()).problems;
    expect(problems).toContain('Doc.open: missing on native (method); add it to the native binding or declare it wasm-only');
    expect(problems).toContain('Doc.prototype.open: only on native (method); declare it in the native extras or add it to wasm');
  });

  test('getter versus method is a difference', () => {
    const native = discover(nativeModule((proto) => { Object.defineProperty(proto, 'title', { value() { return ''; } }); }));
    expect(compareRuntime(discover(wasmModule()), native, none()).problems).toEqual(['Doc.prototype.title: getter on wasm, method on native']);
  });

  test('a missing class is reported once, not per member', () => {
    expect(compareRuntime(discover(wasmModule()), discover({ exportJson() {} }), none()).problems)
      .toEqual(['Doc: missing on native (class); add it to the native binding or declare it wasm-only']);
  });

  test('declared extras pass, are counted, and pending ones are counted apart', () => {
    const native = discover({ ...nativeModule((proto) => { proto.renderPngTarget = () => new Uint8Array(); }), engineVersion() { return ''; } });
    const allow = none();
    allow.nativeOnly['Doc.prototype.renderPngTarget'] = { reason: 'native raster entry point', backs: 'renderPng' };
    allow.nativeOnly.engineVersion = { reason: 'no wasm equivalent yet', pending: 'item 8' };
    const { problems, stats } = compareRuntime(discover(wasmModule()), native, allow);
    expect(problems).toEqual([]);
    expect(stats.extras).toBe(1);
    expect(stats.pending).toBe(1);
  });

  test('a declared extra that no longer exists fails', () => {
    const allow = none();
    allow.nativeOnly['Doc.prototype.renderPngTarget'] = { reason: 'native raster entry point', backs: 'renderPng' };
    expect(compareRuntime(discover(wasmModule()), discover(nativeModule()), allow).problems)
      .toEqual(['Doc.prototype.renderPngTarget: declared native-only but not present; remove the stale entry']);
  });

  test('a declared extra that wasm now has too fails', () => {
    const allow = none();
    allow.nativeOnly['Doc.prototype.sheetCount'] = { reason: 'stale', backs: 'sheetCount' };
    expect(compareRuntime(discover(wasmModule()), discover(nativeModule()), allow).problems)
      .toEqual(['Doc.prototype.sheetCount: declared native-only but present on both engines; remove the stale entry']);
  });

  test('engine internals are skipped, and a stale internal fails', () => {
    const allow = none();
    allow.internals.wasm.initSync = { reason: 'wasm-bindgen instantiation' };
    allow.internals.native.__napiBindingTarget = { reason: 'napi loader' };
    const wasm = discover(wasmModule({ initSync() {} }));
    expect(compareRuntime(wasm, discover(nativeModule()), allow).problems)
      .toEqual(['__napiBindingTarget: listed as a native engine internal but not present']);
  });
});

describe('raw engine surface declarations', () => {
  const declared = (text: string) => parseSurface(text, { paramNames: true });

  test('matching declarations pass, snake and camel names compare equal', () => {
    expect(normalizeParamName('client_id')).toBe('clientId');
    const wasm = WASM_D_TS.replace('open(bytes', 'open(source_bytes');
    const native = NATIVE_D_TS.replace('open(bytes', 'open(sourceBytes');
    expect(compareDeclarations(discover(wasmModule()), discover(nativeModule()), declared(wasm), declared(native), none())).toEqual([]);
  });

  test.each([
    ['arity', 'sheetCount(): number', 'sheetCount(index: number): number', 'Doc.prototype.sheetCount: arity 0 on wasm, 1 on native'],
    ['parameter name', 'static open(bytes: Uint8Array): Doc', 'static open(data: Uint8Array): Doc', 'Doc.open: parameter 1 name differs'],
    ['parameter type', 'static open(bytes: Uint8Array): Doc', 'static open(bytes: string): Doc', 'Doc.open: parameter 1 type differs'],
    ['optionality', 'static open(bytes: Uint8Array): Doc', 'static open(bytes?: Uint8Array): Doc', 'Doc.open: parameter 1 optionality differs'],
    ['sync versus async return', 'sheetCount(): number', 'sheetCount(): Promise<number>', 'Doc.prototype.sheetCount: return type differs'],
  ])('a changed %s is named', (_, from, to, expected) => {
    const problems = compareDeclarations(
      discover(wasmModule()), discover(nativeModule()), declared(WASM_D_TS), declared(NATIVE_D_TS.replace(from, to)), none(),
    );
    expect(problems).toHaveLength(1);
    expect(problems[0]).toStartWith(expected);
  });

  test('a runtime member without a declaration fails', () => {
    const native = discover(nativeModule((proto) => { proto.extra = () => 0; }));
    expect(compareDeclarations(discover(wasmModule()), native, declared(WASM_D_TS), declared(NATIVE_D_TS), none()))
      .toEqual(['Doc.prototype.extra: method on native at runtime but not in its declarations']);
  });

  test('a declared member absent at runtime fails', () => {
    const native = discover(nativeModule((proto) => { delete proto.sheetCount; }));
    expect(compareDeclarations(discover(wasmModule()), native, declared(WASM_D_TS), declared(NATIVE_D_TS), none()))
      .toEqual(['Doc.prototype.sheetCount: declared on native but absent at runtime']);
  });
});

describe('engine module arity', () => {
  test('JavaScript functions compare by parameter count, native functions are skipped', () => {
    const wasm = jsArity({ default(input?: unknown) { return input; }, read(a: number, b: number) { return a + b; }, parse: Math.max });
    const native = jsArity({ default() {}, read(a: number) { return a; }, parse: Math.min });
    expect(wasm.has('parse')).toBe(false);
    expect(compareArity(wasm, native, { default: { reason: 'init input ignored natively', pending: 'item 3' } }))
      .toEqual(['read: takes 2 parameters on wasm, 1 on native']);
  });

  test('a declared arity difference that no longer exists fails', () => {
    const same = jsArity({ default(input?: unknown) { return input; } });
    expect(compareArity(same, same, { default: { reason: 'init input ignored natively', pending: 'item 3' } }))
      .toEqual(['default: declared arity difference no longer exists; remove the stale entry']);
  });
});

describe('merging and scoping', () => {
  test('a name declared differently by two merged modules is named', () => {
    const { problems } = mergeDeclarations([
      ['edit', parseSurface('export function close_display_list(handle: number): void;')],
      ['query', parseSurface('export function close_display_list(handle: number, all: boolean): void;')],
    ]);
    expect(problems).toEqual(['close_display_list: declared differently in edit and query']);
  });

  test('allowances are scoped to the pair that exports them, and an entry no pair reaches is stale', () => {
    const allow = none();
    allow.nativeOnly['Doc.prototype.renderPngTarget'] = { reason: 'native raster entry point', backs: 'renderPng' };
    allow.nativeOnly['Gone.prototype.x'] = { reason: 'stale', backs: 'x' };
    const reached = new Set<string>();
    const pngPair = scopeAllowances(allow, discover({ renderPng() {} }), discover({ renderPng() {} }), reached);
    expect(Object.keys(pngPair.nativeOnly)).toEqual([]);
    const native = discover(nativeModule((proto) => { proto.renderPngTarget = () => new Uint8Array(); }));
    const editPair = scopeAllowances(allow, discover(wasmModule()), native, reached);
    expect(compareRuntime(discover(wasmModule()), native, editPair).problems).toEqual([]);
    expect(unreached(allow, reached)).toEqual(['Gone.prototype.x: allowed but not exported by any engine module; remove the stale entry']);
  });
});
