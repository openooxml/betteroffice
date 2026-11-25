import { parentPort } from 'node:worker_threads';
import { fontAssetResponse, withBrowserFontGlobals } from './fontAssets.test-support';
import type { FontCdnConfiguration } from './fontCdn';
import { openSlicedFace } from './progressive-entry';
import { createFontProvider, loadBundledFontBytes, resolveMetricCompatFace } from './worker';
import { setWorkerFontAssets } from './workerAssets';

const port = parentPort!;
port.once('message', async (cdn: FontCdnConfiguration) => {
  withBrowserFontGlobals(false);
  setWorkerFontAssets(async () => ({ cdn }));
  const requests: string[] = [];
  const cjkRequests: string[] = [];
  let warnings = 0;
  let cjkOffline = false;
  console.warn = () => { warnings += 1; };
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    const url = String(input);
    if (url.includes('/fonts-latin@')) {
      requests.push(url);
      throw new TypeError('Offline');
    }
    if (url.includes('/fonts-cjk@') && cjkOffline) {
      cjkRequests.push(url);
      throw new TypeError('Offline');
    }
    return fontAssetResponse(input);
  }) as typeof fetch;
  const statuses: string[] = [];
  for (const [family, bold] of [['Calibri', false], ['Calibri', true], ['Times New Roman', false]] as const) {
    const [result] = await Promise.allSettled([
      loadBundledFontBytes(resolveMetricCompatFace(family, bold, false)!, undefined, { strict: true }),
    ]);
    statuses.push(result!.status);
  }

  const provider = createFontProvider(undefined, { strict: true });
  const japanese = provider.resolveScriptFallback('cjk-jp', false, false)!;
  const korean = provider.resolveScriptFallback('cjk-kr', false, false)!;
  const open = await openSlicedFace(japanese.slices!);
  let text = '';
  for (let cp = 0x4e00; cp <= 0x9fff && !text; cp += 1) {
    const character = String.fromCodePoint(cp);
    const missing = new Set(open.chunksOf(character).filter((chunk) => !open.loadedChunks().has(chunk)));
    if (missing.size === 1) text = character;
  }
  if (!text) throw new Error('No character uses exactly one unloaded CJK chunk');
  cjkOffline = true;
  await Promise.allSettled([open.ensureText(text)]);
  await Promise.allSettled([open.ensureText(text), korean(), openSlicedFace(korean.slices!)]);
  port.postMessage({ requests, statuses, cjkRequests, warnings });
});
