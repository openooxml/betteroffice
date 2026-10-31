import { fileURLToPath } from 'node:url';
import { resolve } from 'node:path';

const root = fileURLToPath(new URL('..', import.meta.url));

/** The engine module each wasm package links, and the module the matching -native package links in its place. */
export const NATIVE_ENGINES = {
  xlsx: {
    engine: 'packages/xlsx/src/wasm/engine.ts',
    native: 'bindings/xlsx-native/js/engine.ts',
  },
  pptx: {
    engine: 'packages/pptx/src/wasm/engine.ts',
    native: 'bindings/pptx-native/js/engine.ts',
  },
} as const;

/** The wasm asset module each -native package build replaces, so no .wasm file is referenced. */
export const NATIVE_ASSETS = {
  xlsx: {
    asset: 'packages/xlsx/src/wasm/asset.ts',
    native: 'bindings/xlsx-native/js/asset.ts',
  },
  pptx: {
    asset: 'packages/pptx/src/wasm/asset.ts',
    native: 'bindings/pptx-native/js/asset.ts',
  },
} as const;

export const NATIVE_ENGINE_FILTER = /[\\/]src[\\/]wasm[\\/]engine\.ts$/;

const aliases = new Map(
  Object.values(NATIVE_ENGINES).map(({ engine, native }) => [resolve(root, engine), resolve(root, native)])
);

/** The native engine module that replaces a wasm package's engine module, else undefined. */
export function nativeEngineFor(path: string): string | undefined {
  return aliases.get(resolve(path));
}

function reexport(path: string): string {
  const specifier = JSON.stringify(path);
  return `export * from ${specifier};\nexport { default } from ${specifier};\n`;
}

/** Module source that re-exports the native engine in place of the wasm engine module at `path`. */
export function nativeEngineSource(path: string): string | undefined {
  const native = nativeEngineFor(path);
  return native === undefined ? undefined : reexport(native);
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

/** Source a -native package build of `format` links in place of the wasm engine or asset module at `path`, else undefined. */
export function nativeBuildSource(format: keyof typeof NATIVE_ENGINES, path: string): string | undefined {
  const absolute = resolve(path);
  const { engine, native } = NATIVE_ENGINES[format];
  if (absolute === resolve(root, engine)) return reexport(resolve(root, native));
  const asset = (NATIVE_ASSETS as Record<string, { asset: string; native: string }>)[format];
  if (asset && absolute === resolve(root, asset.asset)) return `export * from ${JSON.stringify(resolve(root, asset.native))};\n`;
  return undefined;
}

/** esbuild plugin for a -native package build: links the native engine and asset modules of `format`. */
export function nativeEnginePlugin(format: keyof typeof NATIVE_ENGINES) {
  return {
    name: 'betteroffice-native-engine',
    setup(build: PluginBuild) {
      build.onLoad({ filter: /[\\/]src[\\/]wasm[\\/](engine|asset)\.ts$/ }, ({ path }) => {
        const contents = nativeBuildSource(format, path);
        return contents === undefined ? undefined : { contents, loader: 'ts' };
      });
    },
  };
}
