import { parentPort } from 'node:worker_threads';
import { withBrowserFontGlobals } from '../../../../shared/office-session/testing/fontAssets';
import type { FontCdnConfiguration } from '../fontCdn';
import { loadBundledFontBytes, resolveMetricCompatFace } from '../worker';
import { setWorkerFontAssets } from '../workerAssets';

const port = parentPort!;
port.once('message', async (cdn: FontCdnConfiguration) => {
  withBrowserFontGlobals(false);
  setWorkerFontAssets(async () => ({ cdn }));
  const requests: string[] = [];
  let warnings = 0;
  console.warn = () => { warnings += 1; };
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    requests.push(String(input));
    throw new TypeError('Offline');
  }) as typeof fetch;
  const statuses: string[] = [];
  for (const [family, bold] of [['Calibri', false], ['Calibri', true], ['Times New Roman', false]] as const) {
    const [result] = await Promise.allSettled([
      loadBundledFontBytes(resolveMetricCompatFace(family, bold, false)!, undefined, { strict: true }),
    ]);
    statuses.push(result!.status);
  }
  port.postMessage({ requests, statuses, warnings });
});
