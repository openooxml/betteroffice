import { parentPort } from 'node:worker_threads';
import { stubBrowser } from '../../../docx/src/yrs/renderPngTestSupport';
import { setFontCdnConfiguration } from '../../../../shared/office-session/fontCdn';
import type { SessionScope } from '../../../../shared/office-session/transport';

const port = parentPort!;
Bun.plugin({
  name: 'uninstalled-font-packs',
  setup(build) {
    build.onResolve({ filter: /^@betteroffice\/fonts-(?:latin|intl|cjk)$/ }, ({ path }) => ({ path, namespace: 'uninstalled-font-pack' }));
    build.onLoad({ filter: /.*/, namespace: 'uninstalled-font-pack' }, ({ path }) => ({
      loader: 'object',
      exports: { then: (_resolve: unknown, reject: (error: Error) => void) => reject(new Error(`Cannot find package '${path}'`)) },
    }));
  },
});
setFontCdnConfiguration({ baseUrl: 'https://stale-worker-cdn.example/', timeoutMs: 1000 });
const fetchAsset = globalThis.fetch;
globalThis.fetch = ((input, options) => {
  const url = input instanceof Request ? input.url : String(input);
  if (url.startsWith('https://cdn.jsdelivr.net/') || url.startsWith('https://stale-worker-cdn.example/')) {
    port.postMessage({ kind: 'unexpected-cdn-fetch', url });
    return Promise.resolve(new Response(null, { status: 404 }));
  }
  return fetchAsset(input, options);
}) as typeof fetch;

stubBrowser();
class Face {
  constructor(readonly family: string, bytes: ArrayBuffer | ArrayBufferView) {
    port.postMessage({ kind: 'font-face', byteLength: bytes.byteLength });
  }
  load() { return Promise.resolve(this); }
}
const faces = new Set<Face>();
const document = { fonts: {
  add: (face: Face) => faces.add(face),
  delete: (face: Face) => faces.delete(face),
  load: async (font: string) => [...faces].filter(({ family }) => font.includes(`"${family}"`)),
} };
Object.assign(globalThis, { FontFace: Face, document, window: { document } });
const scope = port as unknown as SessionScope;
if (process.env.BETTEROFFICE_TEST_NODE_WORKER_FORMAT === 'pptx') {
  (await import('../../../pptx/src/session/workerStart')).startSessionWorker(scope);
} else {
  (await import('./workerStart')).startSessionWorker(scope);
}
