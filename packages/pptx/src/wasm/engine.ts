import { version } from '../../package.json';
import wasmInit, {
  decodeTiffPng as wasmDecodeTiffPng,
  exportPptxMarkdownJson as wasmExportPptxMarkdownJson,
  exportPptxStructuredJson as wasmExportPptxStructuredJson,
  parsePptxJson as wasmParsePptxJson,
  parsePptxJsonWithoutMedia as wasmParsePptxJsonWithoutMedia,
  PptxDocument as WasmPptxDocument,
  PptxRenderer as WasmPptxRenderer,
  renderPptxMarkdownJson as wasmRenderPptxMarkdownJson,
  rendererVersion as wasmRendererVersion,
} from './generated/pptx_wasm.js';

export type { InitInput } from './generated/pptx_wasm.js';
export const packageVersion: string = version;

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
  decodeTiffPng?: unknown;
  exportPptxMarkdownJson?: unknown;
  exportPptxStructuredJson?: unknown;
  parsePptxJson?: unknown;
  parsePptxJsonWithoutMedia?: unknown;
  PptxDocument?: unknown;
  PptxRenderer?: unknown;
  renderPptxMarkdownJson?: unknown;
  rendererVersion?: unknown;
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
  return requireFromHere('@betteroffice/pptx-native');
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
        `BETTEROFFICE_ENGINE=native but @betteroffice/pptx-native failed to load: ${
          error instanceof Error ? error.message : String(error)
        }`
      );
    }
    return 'wasm';
  }
  const native = loaded as NativeEngineModule | null | undefined;
  if (typeof native?.PptxDocument !== 'function') {
    if (forced) {
      throw new Error(
        'BETTEROFFICE_ENGINE=native but @betteroffice/pptx-native does not export PptxDocument; ' +
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
        `BETTEROFFICE_ENGINE=native but @betteroffice/pptx-native version ${
          nativeVersion ?? 'unknown'
        } does not match @betteroffice/pptx version ${packageVersion}`
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

/** Which engine runs this presentation code ('native' when @betteroffice/pptx-native is installed and matches this version on Node, Bun or Deno, else 'wasm'). */
export function engineBackend(): EngineBackend {
  return backend;
}

const nativeModule = backend === 'native' ? (loadedNative as NativeEngineModule) : undefined;

export const decodeTiffPng = (nativeModule?.decodeTiffPng ??
  wasmDecodeTiffPng) as typeof wasmDecodeTiffPng;
export const exportPptxMarkdownJson = (nativeModule?.exportPptxMarkdownJson ??
  wasmExportPptxMarkdownJson) as typeof wasmExportPptxMarkdownJson;
export const exportPptxStructuredJson = (nativeModule?.exportPptxStructuredJson ??
  wasmExportPptxStructuredJson) as typeof wasmExportPptxStructuredJson;
export const parsePptxJson = (nativeModule?.parsePptxJson ?? wasmParsePptxJson) as typeof wasmParsePptxJson;
export const parsePptxJsonWithoutMedia = (nativeModule?.parsePptxJsonWithoutMedia ??
  wasmParsePptxJsonWithoutMedia) as typeof wasmParsePptxJsonWithoutMedia;
export const PptxDocument = (nativeModule?.PptxDocument ?? WasmPptxDocument) as typeof WasmPptxDocument;
export type PptxDocument = WasmPptxDocument;
export const PptxRenderer = (nativeModule?.PptxRenderer ?? WasmPptxRenderer) as typeof WasmPptxRenderer;
export type PptxRenderer = WasmPptxRenderer;
export const renderPptxMarkdownJson = (nativeModule?.renderPptxMarkdownJson ??
  wasmRenderPptxMarkdownJson) as typeof wasmRenderPptxMarkdownJson;
export const rendererVersion = (nativeModule?.rendererVersion ??
  wasmRendererVersion) as typeof wasmRendererVersion;

async function initNativeEngine(): Promise<{ memory: undefined }> {
  return { memory: undefined };
}

export default (backend === 'native' ? initNativeEngine : wasmInit) as unknown as typeof wasmInit;
