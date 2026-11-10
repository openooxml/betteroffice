import type { Allowances, Extra, Internal } from '../scripts/raw-engine-surface';
import type { DocxNativeExtras, NativeExtras, PptxNativeExtras, XlsxNativeExtras } from './native-extras-types';

/** Exactly one reason per declared extra: a missing or excess key is a type error. */
export type ExtraReasons<E extends NativeExtras> = {
  functions: { [K in keyof E['functions']]-?: Extra };
  members: { [C in keyof E['members']]-?: { [M in keyof E['members'][C]]-?: Extra } };
};

const fontPacks = {
  configureFontPacks: { reason: 'points the native engine at the installed font add-on packages', backs: 'the font pack loading of the engine module' },
  fontPackDiagnosticsJson: { reason: 'reports the font packs configureFontPacks left out', backs: 'fontPackDiagnostics()' },
} as const satisfies Record<string, Extra>;

export const xlsxExtras = {
  functions: {
    ...fontPacks,
    engineVersion: { reason: 'native-only engine version; no wasm equivalent mapped yet', pending: 'item 8' },
    engineHeapLiveBytes: { reason: 'native-only heap reporting; no wasm equivalent mapped yet', pending: 'item 8' },
    setHeapLimit: { reason: 'native-only heap budget; no wasm equivalent mapped yet', pending: 'item 8' },
  },
  members: {
    XlsxDocument: { renderPngTarget: { reason: 'renders in the native rasterizer instead of the wasm canvas path', backs: 'renderPng' } },
  },
} satisfies ExtraReasons<XlsxNativeExtras>;

export const pptxExtras = {
  functions: {
    ...fontPacks,
    engineVersion: { reason: 'native-only engine version; no wasm equivalent mapped yet', pending: 'item 8' },
  },
  members: {
    PptxRenderer: {
      registerDefaultFonts: { reason: 'registers the bundled faces the native rasterizer needs for a slide', backs: 'renderPng' },
      renderPngTarget: { reason: 'renders in the native rasterizer instead of the wasm canvas path', backs: 'renderPng' },
    },
  },
} satisfies ExtraReasons<PptxNativeExtras>;

export const docxExtras = {
  functions: {},
  members: {
    EditSession: { renderPngTarget: { reason: 'renders in the native rasterizer instead of the wasm canvas path', backs: 'renderPng' } },
  },
} satisfies ExtraReasons<DocxNativeExtras>;

const WASM_CLASSES = {
  xlsx: { wrapped: ['XlsxDocument'], all: ['XlsxDocument', 'XlsxSnapshotBuilder'] },
  pptx: { wrapped: ['PptxDocument'], all: ['PptxDocument', 'PptxRenderer', 'MetafileConverter'] },
  docx: { wrapped: ['RetainedLayoutMeta', 'RunInheritance'], all: ['EditSession', 'RetainedLayoutMeta', 'RunInheritance'] },
} as const;

type Format = keyof typeof WASM_CLASSES;

function wasmInternals(format: Format, level: 'raw' | 'engine'): Record<string, Internal> {
  const internals: Record<string, Internal> = {};
  for (const name of WASM_CLASSES[format].all) {
    if (name === 'MetafileConverter' && level === 'engine') continue;
    internals[`${name}.prototype.__destroy_into_raw`] = { reason: 'wasm-bindgen glue: releases the Rust pointer for free()' };
  }
  for (const name of WASM_CLASSES[format].wrapped) {
    internals[`${name}.__wrap`] = { reason: 'wasm-bindgen glue: wraps a Rust pointer returned by the engine' };
  }
  if (level === 'raw') {
    internals.initSync = { reason: 'wasm-bindgen instantiation; the native binding needs none' };
    internals.default = { reason: 'wasm-bindgen instantiation; the native binding needs none' };
  }
  return internals;
}

/** The binding classes get [Symbol.dispose] from their engine module (shared/nativeDispose.ts), so only the raw binding lacks it. */
const disposeFromEngineModule = (format: Format): Record<string, Extra> => Object.fromEntries(
  WASM_CLASSES[format].all
    .filter((name) => name !== 'MetafileConverter')
    .map((name) => [`${name}.prototype.[Symbol.dispose]`, { reason: 'installed on the binding class by its engine module', backs: 'Symbol.dispose' }]),
);

function flatten(extras: { functions: object; members: object }): Record<string, Extra> {
  const flat: Record<string, Extra> = {};
  for (const [name, extra] of Object.entries(extras.functions)) flat[name] = extra as Extra;
  for (const [className, members] of Object.entries(extras.members)) {
    for (const [name, extra] of Object.entries(members as object)) flat[`${className}.prototype.${name}`] = extra as Extra;
  }
  return flat;
}

const NAPI_INTERNALS: Record<string, Internal> = { __napiBindingTarget: { reason: 'napi-rs loader: which prebuilt binary was loaded' } };

const PANIC_HOOK: Internal = { reason: 'wasm panic-hook installer; not part of any engine module surface' };

/** Formats whose wasm glue exports install_panic_hook; docx-native carries an empty one (gap item 2). */
const PANIC_HOOK_FORMATS: Record<Format, { wasm: boolean; native: boolean }> = {
  xlsx: { wasm: true, native: false },
  pptx: { wasm: false, native: false },
  docx: { wasm: true, native: true },
};

/** Engine module functions whose parameter count differs, per wasm engine module: native init takes no input. */
export function arityAllowances(format: Format): Record<string, Record<string, Extra>> {
  const ignored: Extra = { reason: 'native init ignores the wasm module input', pending: 'item 3' };
  if (format !== 'docx') return { [`packages/${format}/src/wasm/engine.ts`]: { default: ignored } };
  return Object.fromEntries(['Edit', 'Parse', 'Query', 'Opc', 'Preflight']
    .map((name) => [`packages/docx/src/wasm/engine${name}.ts`, { default: ignored, initSync: ignored }]));
}

/** Allowances for the raw compare: wasm-bindgen glue against the napi binding. */
export function rawAllowances(format: Format): Allowances {
  const extras = { xlsx: xlsxExtras, pptx: pptxExtras, docx: docxExtras }[format];
  const wasmOnly: Record<string, Extra> = { ...disposeFromEngineModule(format) };
  if (format === 'pptx') wasmOnly.MetafileConverter = { reason: 'metafile conversion runs inside the native engine; no standalone converter', pending: 'item 11' };
  return {
    nativeOnly: flatten(extras),
    wasmOnly,
    internals: {
      wasm: { ...wasmInternals(format, 'raw'), ...(PANIC_HOOK_FORMATS[format].wasm ? { install_panic_hook: PANIC_HOOK } : {}) },
      native: { ...NAPI_INTERNALS, ...(PANIC_HOOK_FORMATS[format].native ? { install_panic_hook: PANIC_HOOK } : {}) },
    },
  };
}

/** Allowances for the engine modules the packages link: the shims re-export only the shared API. */
export function engineAllowances(format: Format): Allowances {
  const extras = { xlsx: xlsxExtras, pptx: pptxExtras, docx: docxExtras }[format];
  return {
    nativeOnly: flatten({ functions: {}, members: extras.members }),
    wasmOnly: {},
    internals: { wasm: wasmInternals(format, 'engine'), native: {} },
  };
}

/** Engine module exports whose declared type differs on purpose, per wasm engine module; checked by engine-surface.typecheck.ts. */
export const typeAllowances: Record<string, Record<string, Extra>> = {
  'packages/pptx/src/wasm/engine.ts': {
    metafileSvgBindings: { reason: 'rejects on native: the native engine converts metafiles itself, so there are no converter bindings', pending: 'item 11' },
  },
};
