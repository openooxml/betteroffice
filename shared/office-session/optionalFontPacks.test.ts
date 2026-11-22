import { expect, mock, spyOn, test } from 'bun:test';
import { createSessionWorker as createDocxWorker } from '../../packages/docx/src/yrs/sessionWorker.browser';
import { createSessionWorker as createPptxWorker } from '../../packages/pptx/src/session/sessionWorker.browser';
import { createSessionWorker as createXlsxWorker } from '../../packages/xlsx/src/session/sessionWorker.browser';
import {
  createFontPackResponder, createWorkerFontPacks, installWorkerFontPacks,
  type FontAssetKind, type FontPackReply, type FontPackRequest, type FontPackUrls,
} from './optionalFontPacks';
import { attachFontPackSource } from './fontPackSource';
import { createWorkerTransport } from './transport';

test('browser worker constructors attach a font pack listener that the transport removes', () => {
  const original = globalThis.Worker;
  const listen = mock((_type: string, _listener: EventListener) => {});
  const remove = mock((_type: string, _listener: EventListener) => {});
  const terminate = mock(() => {});
  globalThis.Worker = class {
    addEventListener = listen;
    removeEventListener = remove;
    terminate = terminate;
  } as unknown as typeof Worker;
  try {
    for (const create of [createDocxWorker, createPptxWorker, createXlsxWorker]) {
      listen.mockClear();
      remove.mockClear();
      terminate.mockClear();
      const worker = create();
      expect(listen).toHaveBeenCalledTimes(1);
      expect(listen).toHaveBeenCalledWith('message', expect.any(Function));
      createWorkerTransport(worker).close();
      expect(listen).toHaveBeenCalledTimes(1);
      expect(remove.mock.calls).toEqual(listen.mock.calls);
      expect(terminate).toHaveBeenCalledTimes(1);
    }
  } finally {
    globalThis.Worker = original;
  }
});

test('the worker keeps its lazy pack responder through session listener handoff', async () => {
  const page = Object.assign(new EventTarget(), {
    postMessage(data: unknown): void { scope.dispatchEvent(new MessageEvent('message', { data })); },
  });
  const scope = Object.assign(new EventTarget(), {
    postMessage(data: unknown): void { page.dispatchEvent(new MessageEvent('message', { data })); },
  });
  const imports = mock(async (pack: FontAssetKind) => ({ [`${pack}.zst`]: `https://fonts.example/${pack}.zst` }));
  const listen = spyOn(page, 'addEventListener');
  attachFontPackSource(page, imports);
  attachFontPackSource(page, imports);
  expect(listen).toHaveBeenCalledTimes(1);
  listen.mockRestore();
  let source!: ReturnType<typeof createWorkerFontPacks>['source'];
  const dispose = installWorkerFontPacks(scope, (installed) => { source = installed; });
  const prepared = mock((_event: Event) => {});
  const session = mock((_event: Event) => {});
  const host = mock((_event: Event) => {});
  page.addEventListener('message', prepared);
  scope.addEventListener('message', host);
  try {
    await Promise.resolve();
    expect(imports).not.toHaveBeenCalled();
    const first = source('intl');
    expect(source('intl')).toBe(first);
    page.removeEventListener('message', prepared);
    page.addEventListener('message', session);
    expect((await first)?.urls).toEqual({ 'intl.zst': 'https://fonts.example/intl.zst' });
    await source('intl');
    await source('cjk');
    await source('cjk');
    await source('latin');
    await source('decoder');
    expect(imports.mock.calls).toEqual([['intl'], ['intl'], ['cjk'], ['cjk'], ['latin'], ['decoder']]);
    expect(prepared).not.toHaveBeenCalled();
    expect(session).not.toHaveBeenCalled();
    expect(host).not.toHaveBeenCalled();
    scope.postMessage({ protocol: 1, kind: 'progress' });
    expect(session).toHaveBeenCalledTimes(1);
  } finally {
    dispose();
  }
});

test('the installer consumes replies before later boot and session listeners', async () => {
  const requests: Array<{ id: string }> = [];
  const scope = Object.assign(new EventTarget(), { postMessage: (message: { id: string }) => requests.push(message) });
  let source!: ReturnType<typeof createWorkerFontPacks>['source'];
  const dispose = installWorkerFontPacks(scope, (installed) => { source = installed; });
  const held = mock(() => {});
  const session = mock(() => {});
  scope.addEventListener('message', held);
  scope.addEventListener('message', session);
  const pending = source('intl');
  const urls = { 'intl.zst': 'https://fonts.example/intl.zst' };
  scope.dispatchEvent(new MessageEvent('message', { data: {
    protocol: 1, kind: 'font-pack-reply', id: requests[0]!.id, urls,
  } }));
  expect(await pending).toEqual({ urls });
  scope.dispatchEvent(new MessageEvent('message', { data: { protocol: 1, kind: 'font-pack-reply', id: 'unknown' } }));
  expect(held).not.toHaveBeenCalled();
  expect(session).not.toHaveBeenCalled();
  scope.dispatchEvent(new MessageEvent('message', { data: { protocol: 1, kind: 'open' } }));
  expect(held).toHaveBeenCalledTimes(1);
  expect(session).toHaveBeenCalledTimes(1);
  dispose();
});

test('the page shares concurrent lookups and retries failed or missing packs', async () => {
  const urls = { 'cjk.zst': 'https://fonts.example/cjk.zst' };
  const imports = mock(async (): Promise<FontPackUrls | undefined> => urls)
    .mockRejectedValueOnce(new Error('Import failed')).mockResolvedValueOnce(undefined);
  const replies: FontPackReply[] = [];
  const responder = createFontPackResponder((reply) => replies.push(reply), imports);
  const send = async (...ids: string[]) => {
    for (const id of ids) responder.receive({ protocol: 1, kind: 'font-pack-request', id, pack: 'cjk' });
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
  };
  await send('first', 'concurrent');
  await send('missing');
  await send('installed');
  await send('cached');
  expect(imports).toHaveBeenCalledTimes(4);
  expect(replies).toEqual([
    { protocol: 1, kind: 'font-pack-reply', id: 'first', error: 'Import failed' },
    { protocol: 1, kind: 'font-pack-reply', id: 'concurrent', error: 'Import failed' },
    { protocol: 1, kind: 'font-pack-reply', id: 'missing', urls: undefined },
    { protocol: 1, kind: 'font-pack-reply', id: 'installed', urls },
    { protocol: 1, kind: 'font-pack-reply', id: 'cached', urls },
  ]);
});

test('the worker can continue after the page replies without font assets', async () => {
  const requests: FontPackRequest[] = [];
  const worker = createWorkerFontPacks((message) => { requests.push(message); page.receive(message); });
  const page = createFontPackResponder(worker.receive, async () => undefined);
  try {
    const first = worker.source('latin');
    expect(worker.source('latin')).toBe(first);
    expect(await first).toEqual({ urls: undefined });
    expect(await worker.source('latin')).toEqual({ urls: undefined });
    expect(await worker.source('decoder')).toEqual({ urls: undefined });
    expect(requests.map(({ pack }) => pack)).toEqual(['latin', 'latin', 'decoder']);
  } finally {
    worker.dispose();
  }
});

test('the page passes the configured CDN independently of the installed URL map', async () => {
  const cdn = { baseUrl: 'https://cdn.example/root/', timeoutMs: 1234 };
  const worker = createWorkerFontPacks((message) => { page.receive(message); });
  const page = createFontPackResponder(worker.receive, async () => undefined, () => cdn);
  try {
    expect(await worker.source('cjk')).toEqual({ urls: undefined, cdn });
  } finally {
    worker.dispose();
  }
});

test('one 15-second timeout settles all requests until a late reply from the page', async () => {
  const requests: FontPackRequest[] = [];
  const packs = createWorkerFontPacks((message) => requests.push(message));
  const other = createWorkerFontPacks(() => {});
  const timer = spyOn(globalThis, 'setTimeout');
  try {
    const first = packs.source('decoder');
    expect(packs.source('decoder')).toBe(first);
    const [expire, delay] = timer.mock.calls.at(-1)!;
    const concurrent = packs.source('cjk');
    if (typeof expire !== 'function') throw new Error('Expected a timer callback');
    expire();
    expect(delay).toBe(15_000);
    expect(await Promise.all([first, concurrent])).toEqual([undefined, undefined]);
    packs.receive({ protocol: 1, kind: 'font-pack-reply', id: 'unknown' });
    const timers = timer.mock.calls.length;
    for (const pack of ['base', 'latin', 'intl', 'cjk', 'decoder'] as const) {
      expect(await packs.source(pack)).toBeUndefined();
    }
    expect(timer.mock.calls).toHaveLength(timers);
    expect(requests).toHaveLength(2);

    const independent = other.source('base');
    const closed = expect(independent).rejects.toMatchObject({ code: 'disposed' });
    other.dispose();
    await closed;
    expect(timer.mock.calls).toHaveLength(timers + 1);

    packs.receive({ protocol: 1, kind: 'font-pack-reply', id: requests[0]!.id });
    const retry = packs.source('decoder');
    expect(requests).toHaveLength(3);
    const urls = { 'fontpack.wasm': 'https://fonts.example/fontpack.wasm' };
    packs.receive({ protocol: 1, kind: 'font-pack-reply', id: requests[2]!.id, urls });
    expect(await retry).toEqual({ urls });
  } finally {
    other.dispose();
    packs.dispose();
    timer.mockRestore();
  }
});

test('the worker requests a pack again after an error or missing reply', async () => {
  const requests: FontPackRequest[] = [];
  const packs = createWorkerFontPacks((message) => requests.push(message));
  const reply = (result: Pick<FontPackReply, 'urls' | 'error'>) => packs.receive({
    protocol: 1, kind: 'font-pack-reply', id: requests.at(-1)!.id, ...result,
  });
  const failed = packs.source('cjk');
  reply({ error: 'Import failed' });
  await expect(failed).rejects.toThrow('Import failed');
  const emptyError = packs.source('cjk');
  reply({ error: '' });
  await expect(emptyError).rejects.toMatchObject({ code: 'message', message: '' });
  const missing = packs.source('cjk');
  reply({});
  expect(await missing).toEqual({ urls: undefined });
  const installed = packs.source('cjk');
  const urls = { 'cjk.zst': 'https://fonts.example/cjk.zst' };
  reply({ urls });
  expect(requests).toHaveLength(4);
  expect(await installed).toEqual({ urls });
  const next = packs.source('cjk');
  expect(requests).toHaveLength(5);
  reply({ urls });
  expect(await next).toEqual({ urls });
  packs.dispose();
});

test('font pack replies correlate by unique id and preserve source failures', async () => {
  const requests: Array<{ id: string }> = [];
  const first = createWorkerFontPacks((message) => requests.push(message));
  const second = createWorkerFontPacks((message) => requests.push(message));
  const one = first.source('intl');
  const two = second.source('intl');
  expect(requests[0]!.id).not.toBe(requests[1]!.id);
  const reply = (id: string, error: string) => ({ protocol: 1, kind: 'font-pack-reply', id, error });
  first.receive(reply(requests[1]!.id, 'Reply for a different channel'));
  first.receive(reply(requests[0]!.id, 'Unreadable font pack'));
  second.dispose();
  first.dispose();
  await Promise.all([
    expect(one).rejects.toMatchObject({ code: 'message', message: 'Unreadable font pack' }),
    expect(two).rejects.toMatchObject({ code: 'disposed' }),
  ]);
});

test('a page that never answers costs one bounded wait per worker, not one per request', async () => {
  const timers: Array<{ run: () => void; delay: number }> = [];
  const timeout = spyOn(globalThis, 'setTimeout').mockImplementation(((run: () => void, delay?: number) => {
    timers.push({ run, delay: delay ?? 0 });
    return timers.length as unknown as ReturnType<typeof setTimeout>;
  }) as typeof setTimeout);
  const clear = spyOn(globalThis, 'clearTimeout').mockImplementation((() => {}) as typeof clearTimeout);
  try {
    const posted: FontPackRequest[] = [];
    const channel = createWorkerFontPacks((request) => { posted.push(request); }, 15_000);
    const first = [channel.source('base'), channel.source('latin'), channel.source('cjk')];
    expect(timers.every(({ delay }) => delay === 15_000)).toBe(true);
    timers[0]!.run();
    expect(await Promise.all(first)).toEqual([undefined, undefined, undefined]);
    const created = timers.length;
    const sent = posted.length;
    expect(await Promise.all([channel.source('base'), channel.source('intl'), channel.source('decoder')]))
      .toEqual([undefined, undefined, undefined]);
    expect(timers.length).toBe(created);
    expect(posted.length).toBe(sent);
  } finally {
    timeout.mockRestore();
    clear.mockRestore();
  }
});
