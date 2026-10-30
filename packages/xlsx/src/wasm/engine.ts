import { version as packageVersion } from '../../package.json';
import wasmInit, {
  XlsxDocument as WasmXlsxDocument,
  XlsxSnapshotBuilder as WasmXlsxSnapshotBuilder,
  allocationFailure as wasmAllocationFailure,
  displayListCellLimit as wasmDisplayListCellLimit,
  exportXlsxMarkdownJson as wasmExportXlsxMarkdownJson,
  exportXlsxStructuredJson as wasmExportXlsxStructuredJson,
  renderXlsxMarkdownJson as wasmRenderXlsxMarkdownJson,
} from './generated/xlsx_wasm.js';

export type { InitInput } from './generated/xlsx_wasm.js';
export { packageVersion };

export type EngineBackend = 'wasm' | 'native';

export interface BackendRuntime {
  hasProcess: boolean;
  hasGetBuiltinModule: boolean;
  hasNodeVersion: boolean;
  hasDom: boolean;
}

export interface SelectBackendOptions {
  env: string | undefined;
  runtime: BackendRuntime;
  loadNative: () => unknown;
  packageVersion: string;
}

interface NativeEngineModule {
  XlsxDocument?: unknown;
  XlsxSnapshotBuilder?: unknown;
  allocationFailure?: unknown;
  displayListCellLimit?: unknown;
  exportXlsxStructuredJson?: unknown;
  exportXlsxMarkdownJson?: unknown;
  renderXlsxMarkdownJson?: unknown;
  engineVersion?: unknown;
}

export type ProcessLike = {
  getBuiltinModule?: (id: string) => unknown;
  versions?: { node?: unknown };
  env?: Record<string, string | undefined>;
};

function processLike(): ProcessLike | undefined {
  const proc = (globalThis as { process?: unknown }).process;
  return typeof proc === 'object' && proc !== null ? (proc as ProcessLike) : undefined;
}

function detectRuntime(): BackendRuntime {
  const proc = processLike();
  const globals = globalThis as { window?: unknown; document?: unknown };
  const nodeVersion = proc?.versions?.node;
  return {
    hasProcess: proc !== undefined,
    hasGetBuiltinModule: typeof proc?.getBuiltinModule === 'function',
    hasNodeVersion: typeof nodeVersion === 'string' && nodeVersion.length > 0,
    hasDom: typeof globals.window !== 'undefined' || typeof globals.document !== 'undefined',
  };
}

export function readEnvOverride(proc: ProcessLike | undefined = processLike()): string | undefined {
  try {
    const value = proc?.env?.BETTEROFFICE_ENGINE;
    return typeof value === 'string' ? value : undefined;
  } catch {
    return undefined;
  }
}

function loadNativeModule(): unknown {
  const proc = processLike();
  if (typeof proc?.getBuiltinModule !== 'function') {
    throw new Error('process.getBuiltinModule is unavailable');
  }
  const nodeModule = proc.getBuiltinModule('node:module') as { createRequire?: unknown };
  if (typeof nodeModule?.createRequire !== 'function') {
    throw new Error('node:module does not export createRequire');
  }
  const requireFromHere = (nodeModule.createRequire as (url: string) => (id: string) => unknown)(
    import.meta.url
  );
  return requireFromHere('@betteroffice/xlsx-native');
}

export function selectBackend(options: SelectBackendOptions): EngineBackend {
  const { env, runtime, loadNative, packageVersion } = options;
  if (env === 'wasm') return 'wasm';
  const forced = env === 'native';
  const serverSide =
    runtime.hasProcess && runtime.hasGetBuiltinModule && runtime.hasNodeVersion && !runtime.hasDom;
  if (!serverSide) {
    if (forced) {
      throw new Error(
        'BETTEROFFICE_ENGINE=native requires a server-side JS runtime (Node, Bun or Deno) ' +
          'with process.getBuiltinModule and no DOM'
      );
    }
    return 'wasm';
  }
  let loaded: unknown;
  try {
    loaded = loadNative();
  } catch (error) {
    if (forced) {
      throw new Error(
        `BETTEROFFICE_ENGINE=native but @betteroffice/xlsx-native failed to load: ${
          error instanceof Error ? error.message : String(error)
        }`
      );
    }
    return 'wasm';
  }
  const native = loaded as NativeEngineModule | null | undefined;
  if (typeof native?.XlsxDocument !== 'function') {
    if (forced) {
      throw new Error(
        'BETTEROFFICE_ENGINE=native but @betteroffice/xlsx-native does not export XlsxDocument; ' +
          'the native engine is not installed or is incomplete'
      );
    }
    return 'wasm';
  }
  let nativeVersion: string | undefined;
  try {
    nativeVersion =
      typeof native.engineVersion === 'function'
        ? (native.engineVersion as () => string)()
        : undefined;
  } catch {
    nativeVersion = undefined;
  }
  if (nativeVersion !== packageVersion) {
    if (forced) {
      throw new Error(
        `BETTEROFFICE_ENGINE=native but @betteroffice/xlsx-native version ${
          nativeVersion ?? 'unknown'
        } does not match @betteroffice/xlsx version ${packageVersion}`
      );
    }
    return 'wasm';
  }
  return 'native';
}

let loadedNative: unknown;
let nativeLoadAttempted = false;

function loadNativeOnce(): unknown {
  if (!nativeLoadAttempted) {
    nativeLoadAttempted = true;
    loadedNative = loadNativeModule();
  }
  return loadedNative;
}

export const backend: EngineBackend = selectBackend({
  env: readEnvOverride(),
  runtime: detectRuntime(),
  loadNative: loadNativeOnce,
  packageVersion,
});

/** Which engine runs this workbook code ('native' when @betteroffice/xlsx-native is installed and matches this version on Node, Bun or Deno, else 'wasm'). */
export function engineBackend(): EngineBackend {
  return backend;
}

const nativeModule = backend === 'native' ? (loadedNative as NativeEngineModule) : undefined;

export const XlsxDocument = (nativeModule?.XlsxDocument ?? WasmXlsxDocument) as typeof WasmXlsxDocument;
export type XlsxDocument = WasmXlsxDocument;
export const XlsxSnapshotBuilder = (nativeModule?.XlsxSnapshotBuilder ??
  WasmXlsxSnapshotBuilder) as typeof WasmXlsxSnapshotBuilder;
export type XlsxSnapshotBuilder = WasmXlsxSnapshotBuilder;
export const allocationFailure = (nativeModule?.allocationFailure ??
  wasmAllocationFailure) as typeof wasmAllocationFailure;
export const displayListCellLimit = (nativeModule?.displayListCellLimit ??
  wasmDisplayListCellLimit) as typeof wasmDisplayListCellLimit;
export const exportXlsxStructuredJson = (nativeModule?.exportXlsxStructuredJson ??
  wasmExportXlsxStructuredJson) as typeof wasmExportXlsxStructuredJson;
export const exportXlsxMarkdownJson = (nativeModule?.exportXlsxMarkdownJson ??
  wasmExportXlsxMarkdownJson) as typeof wasmExportXlsxMarkdownJson;
export const renderXlsxMarkdownJson = (nativeModule?.renderXlsxMarkdownJson ??
  wasmRenderXlsxMarkdownJson) as typeof wasmRenderXlsxMarkdownJson;

async function initNativeEngine(): Promise<{ memory: undefined }> {
  return { memory: undefined };
}

export default (backend === 'native' ? initNativeEngine : wasmInit) as unknown as typeof wasmInit;
