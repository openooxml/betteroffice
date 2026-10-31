import { fileURLToPath } from 'node:url';
import { resolve } from 'node:path';

const root = fileURLToPath(new URL('..', import.meta.url));

type NativeEngine = { engine: string; native: string };
type NativeAsset = { asset: string; native: string };

/** The engine modules each wasm package links, and their native replacements. */
export const NATIVE_ENGINES = {
  xlsx: {
    engine: 'packages/xlsx/src/wasm/engine.ts',
    native: 'bindings/xlsx-native/js/engine.ts',
  },
  pptx: {
    engine: 'packages/pptx/src/wasm/engine.ts',
    native: 'bindings/pptx-native/js/engine.ts',
  },
  docx: [
    { engine: 'packages/docx/src/wasm/engineEdit.ts', native: 'bindings/docx-native/js/engineEdit.ts' },
    { engine: 'packages/docx/src/wasm/engineParse.ts', native: 'bindings/docx-native/js/engineParse.ts' },
    { engine: 'packages/docx/src/wasm/engineQuery.ts', native: 'bindings/docx-native/js/engineQuery.ts' },
    { engine: 'packages/docx/src/wasm/engineOpc.ts', native: 'bindings/docx-native/js/engineOpc.ts' },
    { engine: 'packages/docx/src/wasm/enginePreflight.ts', native: 'bindings/docx-native/js/enginePreflight.ts' },
  ],
} as const;

/** The wasm asset modules each -native package build replaces. */
export const NATIVE_ASSETS = {
  xlsx: {
    asset: 'packages/xlsx/src/wasm/asset.ts',
    native: 'bindings/xlsx-native/js/asset.ts',
  },
  pptx: {
    asset: 'packages/pptx/src/wasm/asset.ts',
    native: 'bindings/pptx-native/js/asset.ts',
  },
  docx: [
    { asset: 'packages/docx/src/wasm/assetEdit.ts', native: 'bindings/docx-native/js/assetEdit.ts' },
    { asset: 'packages/docx/src/wasm/assetParse.ts', native: 'bindings/docx-native/js/assetParse.ts' },
    { asset: 'packages/docx/src/wasm/assetQuery.ts', native: 'bindings/docx-native/js/assetQuery.ts' },
    { asset: 'packages/docx/src/wasm/assetOpc.ts', native: 'bindings/docx-native/js/assetOpc.ts' },
    { asset: 'packages/docx/src/wasm/assetPreflight.ts', native: 'bindings/docx-native/js/assetPreflight.ts' },
  ],
} as const;

export const NATIVE_ENTRIES: Record<string, Record<string, string>> = {
  docx: {
    '.': 'src/core.ts',
    './docx': 'src/docx/index.ts',
    './docx/wrapTypes': 'src/docx/wrapTypes.ts',
    './styles': 'src/styles/index.ts',
    './layout/pagination': 'src/layout/pagination/index.ts',
    './layout/render': 'src/layout/render/index.ts',
    './layout': 'src/layout/index.ts',
    './plugin-api': 'src/plugin-api/index.ts',
    './plugin-api/RenderedDomContext': 'src/plugin-api/RenderedDomContext.ts',
    './plugin-api/resolveItemPositions': 'src/plugin-api/resolveItemPositions.ts',
    './types/document': 'src/types/document.ts',
    './types/content': 'src/types/content.ts',
    './utils': 'src/utils/index.ts',
    './utils/cardStyles': 'src/utils/cardStyles.ts',
    './utils/comments': 'src/utils/comments.ts',
    './utils/stylePreview': 'src/utils/stylePreview.ts',
    './utils/findReplace': 'src/utils/findReplace.ts',
    './utils/findVerticalScrollParent': 'src/utils/findVerticalScrollParent.ts',
    './utils/fontOptions': 'src/utils/fontOptions.ts',
    './utils/listState': 'src/utils/listState.ts',
    './utils/reportIssue': 'src/utils/reportIssue.ts',
    './utils/sidebarConstants': 'src/utils/sidebarConstants.ts',
    './utils/units': 'src/utils/units.ts',
    './utils/autoScroll': 'src/utils/autoScroll.ts',
    './yrs': 'src/yrs/index.ts',
    './collaboration': 'src/collaboration/index.ts',
    './editor': 'src/editor/index.ts',
    './worker': 'src/worker.ts',
  },
};

export const NATIVE_ENGINE_FILTER = /[\\/]src[\\/]wasm[\\/](engine|asset)[A-Za-z]*\.ts$/;

function modules<T>(value: T | readonly T[]): readonly T[] {
  return Array.isArray(value) ? value : [value as T];
}

const aliases = new Map(
  Object.values(NATIVE_ENGINES).flatMap((value) => modules<NativeEngine>(value))
    .map(({ engine, native }) => [resolve(root, engine), resolve(root, native)])
);
const preloadAssets = new Map(
  Object.values(NATIVE_ASSETS).flatMap((value) => Array.isArray(value) ? modules<NativeAsset>(value) : [])
    .map(({ asset, native }) => [resolve(root, asset), resolve(root, native)])
);

/** The native engine module that replaces a wasm package's engine module, else undefined. */
export function nativeEngineFor(path: string): string | undefined {
  return aliases.get(resolve(path));
}

function reexport(path: string): string {
  const specifier = JSON.stringify(path);
  return `export * from ${specifier};\nexport { default } from ${specifier};\n`;
}

/** Module source for the native preload replacement. */
export function nativeEngineSource(path: string): string | undefined {
  const native = nativeEngineFor(path);
  if (native !== undefined) return reexport(native);
  const asset = preloadAssets.get(resolve(path));
  return asset === undefined ? undefined : `export * from ${JSON.stringify(asset)};\n`;
}

interface LoadArgs {
  path: string;
}

interface LoadResult {
  contents: string;
  loader: 'ts';
}

interface PluginBuild {
  onLoad(options: { filter: RegExp }, callback: (args: LoadArgs) => LoadResult | undefined): void;
}

/** Source a -native package build links in place of an engine or asset module. */
export function nativeBuildSource(format: keyof typeof NATIVE_ENGINES, path: string): string | undefined {
  const absolute = resolve(path);
  for (const { engine, native } of modules<NativeEngine>(NATIVE_ENGINES[format])) {
    if (absolute === resolve(root, engine)) return reexport(resolve(root, native));
  }
  for (const { asset, native } of modules<NativeAsset>(NATIVE_ASSETS[format])) {
    if (absolute === resolve(root, asset)) return `export * from ${JSON.stringify(resolve(root, native))};\n`;
  }
  return undefined;
}

/** esbuild plugin linking the native engine and asset modules of a format. */
export function nativeEnginePlugin(format: keyof typeof NATIVE_ENGINES) {
  return {
    name: 'betteroffice-native-engine',
    setup(build: PluginBuild) {
      build.onLoad({ filter: NATIVE_ENGINE_FILTER }, ({ path }) => {
        const contents = nativeBuildSource(format, path);
        return contents === undefined ? undefined : { contents, loader: 'ts' };
      });
    },
  };
}
