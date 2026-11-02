import { compileSessionWasm } from '#session-worker';
import { metafileWasmAssetUrl } from '../wasm/asset';
import { metafileSvgBindings } from '../wasm/engine';

const MAX_SOURCE_BYTES = 16 * 1024 * 1024;
const SESSION_BYTES = 64 * 1024 * 1024;

export interface MetafileSvgModule {
  metafileToSvg(bytes: Uint8Array): Uint8Array | undefined;
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
  let module: MetafileSvgModule | undefined;
  let loading: Promise<void> | undefined;
  let sourceBytes = SESSION_BYTES;
  let svgBytes = SESSION_BYTES;
  const cache = new Map<string, Uint8Array | undefined>();
  return {
    start(): Promise<void> {
      if (loading) return loading;
      loading = Promise.resolve().then(load).then((loaded) => { module = loaded; }).catch((error) => {
        loading = undefined;
        throw error;
      });
      return loading;
    },
    ready(): boolean { return module !== undefined; },
    convert(assetId: string, bytes: Uint8Array): Uint8Array | undefined {
      if (!module || bytes.byteLength > MAX_SOURCE_BYTES) return undefined;
      const key = contentKey(assetId, bytes);
      if (cache.has(key)) return cache.get(key);
      if (bytes.byteLength > sourceBytes || svgBytes === 0) return undefined;
      sourceBytes -= bytes.byteLength;
      let svg: Uint8Array | undefined;
      try { svg = module.metafileToSvg(bytes); } catch {}
      if (svg && svg.byteLength <= svgBytes) {
        svg = new Uint8Array(svg);
        svgBytes -= svg.byteLength;
      } else svg = undefined;
      cache.set(key, svg);
      return svg;
    },
  };
}

export type MetafileSvg = ReturnType<typeof createMetafileSvg>;
