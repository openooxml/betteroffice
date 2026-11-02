import { compileSessionWasm } from '#session-worker';
import { metafileWasmAssetUrl } from '../wasm/asset';
import { metafileSvgBindings } from '../wasm/engine';

const MAX_SOURCE_BYTES = 16 * 1024 * 1024;

export interface MetafileSvgModule {
  MetafileConverter: new () => {
    toSvg(bytes: Uint8Array): Uint8Array | undefined;
  };
}

export type MetafileSvgModuleLoader = () => Promise<MetafileSvgModule>;

export const commandMetafileSvgModule = {
  async load(): Promise<MetafileSvgModule> {
    const module = await metafileSvgBindings();
    const url = metafileWasmAssetUrl();
    await module.default({ module_or_path: url.protocol === 'file:' ? await compileSessionWasm(url) : url });
    return module;
  },
};

function contentKey(assetId: string, bytes: Uint8Array): string {
  let hash = 0x811c9dc5;
  for (const byte of bytes) hash = Math.imul(hash ^ byte, 0x01000193) >>> 0;
  return JSON.stringify([assetId, bytes.byteLength, hash]);
}

export function createMetafileSvg(load: MetafileSvgModuleLoader = () => commandMetafileSvgModule.load()) {
  let converter: InstanceType<MetafileSvgModule['MetafileConverter']> | undefined;
  let loading: Promise<void> | undefined;
  const cache = new Map<string, Uint8Array | undefined>();
  return {
    start(): Promise<void> {
      if (loading) return loading;
      loading = Promise.resolve().then(load).then((loaded) => { converter = new loaded.MetafileConverter(); }).catch((error) => {
        loading = undefined;
        throw error;
      });
      return loading;
    },
    ready(): boolean { return converter !== undefined; },
    convert(assetId: string, bytes: Uint8Array): Uint8Array | undefined {
      if (!converter || bytes.byteLength > MAX_SOURCE_BYTES) return undefined;
      const key = contentKey(assetId, bytes);
      if (cache.has(key)) return cache.get(key);
      let svg: Uint8Array | undefined;
      try {
        const output = converter.toSvg(bytes);
        if (output) svg = new Uint8Array(output);
      } catch {}
      cache.set(key, svg);
      return svg;
    },
  };
}

export type MetafileSvg = ReturnType<typeof createMetafileSvg>;
