import { fileURLToPath } from 'node:url';
import { dirname, isAbsolute, resolve } from 'node:path';

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

export const NATIVE_ENGINE_FILTER = /(^|[\\/])engine(\.[jt]s)?$/;

const withoutExtension = (path: string) => path.replace(/\.[jt]s$/, '');

const aliases = new Map(
  Object.values(NATIVE_ENGINES).map(({ engine, native }) => [
    withoutExtension(resolve(root, engine)),
    resolve(root, native),
  ])
);

/** The native engine module for an import of a wasm package's engine module, else undefined. */
export function nativeEngineFor(specifier: string, importer: string): string | undefined {
  const target = isAbsolute(specifier) ? specifier : resolve(dirname(importer), specifier);
  return aliases.get(withoutExtension(target));
}
