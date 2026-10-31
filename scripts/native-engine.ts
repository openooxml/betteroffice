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

export const NATIVE_ENGINE_FILTER = /[\\/]src[\\/]wasm[\\/]engine\.ts$/;

const aliases = new Map(
  Object.values(NATIVE_ENGINES).map(({ engine, native }) => [resolve(root, engine), resolve(root, native)])
);

/** The native engine module that replaces a wasm package's engine module, else undefined. */
export function nativeEngineFor(path: string): string | undefined {
  return aliases.get(resolve(path));
}

/** Module source that re-exports the native engine in place of the wasm engine module at `path`. */
export function nativeEngineSource(path: string): string | undefined {
  const native = nativeEngineFor(path);
  if (native === undefined) return undefined;
  const specifier = JSON.stringify(native);
  return `export * from ${specifier};\nexport { default } from ${specifier};\n`;
}
