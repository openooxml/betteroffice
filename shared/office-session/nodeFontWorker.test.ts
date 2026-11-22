import { expect, test } from 'bun:test';
import { openPresentationSession } from '../../packages/pptx/src/session/client';
import { openWorkbookSession } from '../../packages/xlsx/src/session/client';
import { withBrowserFontGlobals } from '../../packages/fonts/src/test-assets';
import { getFontCdnConfiguration, setFontCdnConfiguration } from './fontCdn';
import { createSessionWorker } from './nodeSessionWorker';
import { isClientMessage } from './protocol';

const engine = (globalThis as Record<symbol, unknown>)[Symbol.for('betteroffice.testEngine')];

test.skipIf(engine === 'native')('Node worker_threads PPTX and XLSX export without the page CDN', async () => {
  const previous = getFontCdnConfiguration();
  const previousFormat = process.env.BETTEROFFICE_TEST_NODE_WORKER_FORMAT;
  const opened: unknown[] = [];
  const requests: string[] = [];
  setFontCdnConfiguration(undefined);
  try {
    for (const format of ['pptx', 'xlsx'] as const) {
      process.env.BETTEROFFICE_TEST_NODE_WORKER_FORMAT = format;
      const worker = () => {
        const thread = createSessionWorker(new URL('./testing/nodeFontWorker.ts', import.meta.url), { type: 'module' }, 'the in-thread test API');
        const post = thread.postMessage.bind(thread);
        thread.postMessage = ((message: unknown, transfer: Transferable[]) => {
          if (isClientMessage(message) && message.kind === 'call' && message.method === 'open') opened.push(message.args[1]);
          post(message, transfer);
        }) as Worker['postMessage'];
        thread.addEventListener('message', (event) => {
          if (event.data?.kind === 'unexpected-cdn-fetch') requests.push(event.data.url);
        });
        return thread;
      };
      if (format === 'pptx') {
        const bytes = await Bun.file(new URL('../../apps/demo/public/betteroffice-demo.pptx', import.meta.url)).bytes();
        const session = await openPresentationSession(bytes, { worker });
        try { expect(await session.call.renderPng({ slide: 0 })).toEqual(new TextEncoder().encode('png')); }
        finally { await session.dispose(); }
      } else {
        const bytes = await Bun.file(new URL('../../packages/xlsx/test-fixtures/sample.xlsx', import.meta.url)).bytes();
        const session = await openWorkbookSession(bytes, { worker });
        try { expect(await session.call.renderPng({ sheet: 0, range: 'A1:B3' })).toEqual(new TextEncoder().encode('png')); }
        finally { await session.dispose(); }
      }
    }
    expect(opened).toHaveLength(2);
    for (const input of opened) expect(input).not.toHaveProperty('fontCdn');
    expect(requests).toEqual([]);
    const restore = withBrowserFontGlobals();
    try { expect(getFontCdnConfiguration()).toMatchObject({ implicit: true }); }
    finally { restore(); }
  } finally {
    if (previousFormat === undefined) delete process.env.BETTEROFFICE_TEST_NODE_WORKER_FORMAT;
    else process.env.BETTEROFFICE_TEST_NODE_WORKER_FORMAT = previousFormat;
    setFontCdnConfiguration(previous);
  }
}, 30_000);
