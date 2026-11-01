import type { HostMessage } from '../../../../../shared/office-session/protocol';

/** Test-only: delays or fails the worker's raster wasm download; every other fetch passes through. */
export function installRasterFetch(mode: { delayMs: number } | { failFirst: true }): void {
  const fetchAsset = globalThis.fetch;
  let rasterFetches = 0;
  globalThis.fetch = ((input, init) => {
    const url = input instanceof Request ? input.url : String(input);
    self.postMessage({ protocol: 1, kind: 'event', name: 'png-test-fetch', payload: url } satisfies HostMessage);
    if (!url.endsWith('/xlsx_raster_bg.wasm')) return fetchAsset(input, init);
    rasterFetches += 1;
    if ('failFirst' in mode) {
      return rasterFetches === 1 ? Promise.reject(new TypeError('Raster download failed')) : fetchAsset(input, init);
    }
    return new Promise<Response>((resolve, reject) => {
      setTimeout(() => { fetchAsset(input, init).then(resolve, reject); }, mode.delayMs);
    });
  }) as typeof fetch;
}
