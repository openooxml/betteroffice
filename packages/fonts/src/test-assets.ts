import latinManifest from '../../fonts-latin/assets/manifest.json';
import cjkManifest from '../../fonts-cjk/assets/manifest.json';

const latinFiles = new Set(latinManifest.faces.map((face) => face.file));
const cjkFiles = new Set(cjkManifest.faces.map((face) => face.file));

export async function fontAssetResponse(input: RequestInfo | URL): Promise<Response> {
  const path = new URL(String(input)).pathname;
  if (path.endsWith('/decoder/fontpack.wasm')) {
    return new Response(await Bun.file(new URL('../dist/decoder/fontpack.wasm', import.meta.url)).arrayBuffer());
  }
  const file = path.slice(path.lastIndexOf('/') + 1);
  const pack = latinFiles.has(file) ? 'fonts-latin' : cjkFiles.has(file) ? 'fonts-cjk' : 'fonts-intl';
  return new Response(await Bun.file(new URL(`../../${pack}/assets/${file}`, import.meta.url)).arrayBuffer());
}
