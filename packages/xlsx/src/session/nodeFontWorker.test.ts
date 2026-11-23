import { expect, test } from 'bun:test';
import { createPresentationSession } from '../../../pptx/src/session/client';
import { createWorkbookSession } from './client';
import { BASE_LATIN_FONTS } from '@betteroffice/fonts';
import { getFontCdnConfiguration, setFontCdnConfiguration } from '../../../../shared/office-session/fontCdn';
import { createSessionWorker } from '../../../../shared/office-session/nodeSessionWorker';
import { isClientMessage } from '../../../../shared/office-session/protocol';
import { withinTestBudget } from '../../../../shared/office-session/testing/loadBudget';
import { createWorkerTransport, type SessionTransport } from '../../../../shared/office-session/transport';

const engine = (globalThis as Record<symbol, unknown>)[Symbol.for('betteroffice.testEngine')];

test.skipIf(engine === 'native')('T6: Node worker_threads export with Base Latin and clear a stale CDN when opening without a config', async () => {
  const previous = getFontCdnConfiguration();
  const previousFormat = process.env.BETTEROFFICE_TEST_NODE_WORKER_FORMAT;
  const opened: unknown[] = [];
  const requests: string[] = [];
  const rendered: unknown[] = [];
  const faceLengths = new Map<string, number[]>();
  const budgetMs = 10_000;
  setFontCdnConfiguration(undefined);
  try {
    for (const format of ['pptx', 'xlsx'] as const) {
      process.env.BETTEROFFICE_TEST_NODE_WORKER_FORMAT = format;
      const source = format === 'pptx'
        ? new URL('../../../../apps/demo/public/betteroffice-demo.pptx', import.meta.url)
        : new URL('../../test-fixtures/sample.xlsx', import.meta.url);
      const bytes = await Bun.file(source).bytes();
      const thread = createSessionWorker(new URL('./nodeFontWorker.test-support.ts', import.meta.url), { type: 'module' }, 'the in-thread test API');
      const workerTransport = createWorkerTransport(thread);
      const controller = new AbortController();
      const transport: SessionTransport = {
        ...workerTransport,
        post(message, transfer) {
          if (isClientMessage(message) && message.kind === 'call' && message.method === 'open') opened.push(message.args[1]);
          workerTransport.post(message, transfer);
        },
        listen(listener) {
          return workerTransport.listen((message) => {
            const probe = message as { kind?: string; url?: string; byteLength?: number } | null;
            if (probe?.kind === 'unexpected-cdn-fetch' && typeof probe.url === 'string') requests.push(probe.url);
            else if (probe?.kind === 'font-face' && typeof probe.byteLength === 'number') {
              const lengths = faceLengths.get(format) ?? [];
              lengths.push(probe.byteLength);
              faceLengths.set(format, lengths);
            } else listener(message);
          });
        },
        onError(listener) {
          const abort = () => listener(controller.signal.reason);
          controller.signal.addEventListener('abort', abort, { once: true });
          const unsubscribe = workerTransport.onError(listener);
          return () => {
            controller.signal.removeEventListener('abort', abort);
            unsubscribe();
          };
        },
      };
      let session: { dispose(): Promise<void> } | undefined;
      try {
        if (format === 'pptx') {
          const pptx = await withinTestBudget(createPresentationSession(bytes, {}, transport), 'PPTX worker open', budgetMs);
          session = pptx;
          rendered.push(await withinTestBudget(pptx.call.renderPng({ slide: 0 }), 'PPTX worker export', budgetMs));
        } else {
          const xlsx = await withinTestBudget(createWorkbookSession(bytes, {}, transport), 'XLSX worker open', budgetMs);
          session = xlsx;
          rendered.push(await withinTestBudget(xlsx.call.renderPng({ sheet: 0, range: 'A1:B3' }), 'XLSX worker export', budgetMs));
        }
      } finally {
        try {
          if (session) await withinTestBudget(session.dispose(), `${format} worker disposal`);
        } finally {
          controller.abort();
          workerTransport.close();
        }
      }
    }
    expect(opened).toHaveLength(2);
    for (const input of opened) expect(input).not.toHaveProperty('fontCdn');
    expect(requests).toEqual([]);
    const png = new TextEncoder().encode('png');
    expect(rendered).toEqual([png, png]);
    for (const format of ['pptx', 'xlsx']) {
      expect(faceLengths.get(format)?.some((length) => BASE_LATIN_FONTS.some((face) => face.byteLength === length))).toBe(true);
    }
  } finally {
    if (previousFormat === undefined) delete process.env.BETTEROFFICE_TEST_NODE_WORKER_FORMAT;
    else process.env.BETTEROFFICE_TEST_NODE_WORKER_FORMAT = previousFormat;
    setFontCdnConfiguration(previous);
  }
}, 50_000);
