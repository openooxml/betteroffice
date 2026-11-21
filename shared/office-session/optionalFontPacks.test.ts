import { expect, mock, spyOn, test } from 'bun:test';
import { createSessionWorker as createDocxWorker } from '../../packages/docx/src/yrs/sessionWorker.browser';
import { createSessionWorker as createPptxWorker } from '../../packages/pptx/src/session/sessionWorker.browser';
import { createSessionWorker as createXlsxWorker } from '../../packages/xlsx/src/session/sessionWorker.browser';
import {
  createFontPackResponder, createWorkerFontPacks, installWorkerFontPacks,
  type FontPackReply, type FontPackRequest, type FontPackUrls, type OptionalFontPack,
} from './optionalFontPacks';
import { attachFontPackSource } from './fontPackSource';

test('all three browser worker constructors attach a message listener before returning', () => {
  const original = globalThis.Worker;
  const listen = mock((_type: string, _listener: EventListener) => {});
  globalThis.Worker = class { addEventListener = listen; } as unknown as typeof Worker;
  try {
    for (const create of [createDocxWorker, createPptxWorker, createXlsxWorker]) {
      listen.mockClear();
      create();
      expect(listen).toHaveBeenCalledTimes(1);
      expect(listen).toHaveBeenCalledWith('message', expect.any(Function));
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
  const imports = mock(async (pack: OptionalFontPack) => ({ [`${pack}.zst`]: `https://fonts.example/${pack}.zst` }));
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
    expect(await first).toEqual({ 'intl.zst': 'https://fonts.example/intl.zst' });
    await source('intl');
    await source('cjk');
    await source('cjk');
    expect(imports.mock.calls).toEqual([['intl'], ['cjk']]);
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
  expect(await pending).toEqual(urls);
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
  expect(imports).toHaveBeenCalledTimes(3);
  expect(replies).toEqual([
    { protocol: 1, kind: 'font-pack-reply', id: 'first', error: 'Import failed' },
    { protocol: 1, kind: 'font-pack-reply', id: 'concurrent', error: 'Import failed' },
    { protocol: 1, kind: 'font-pack-reply', id: 'missing', urls: undefined },
    { protocol: 1, kind: 'font-pack-reply', id: 'installed', urls },
    { protocol: 1, kind: 'font-pack-reply', id: 'cached', urls },
  ]);
});

test('the worker retries after a 15-second timeout and ignores late replies', async () => {
  const requests: FontPackRequest[] = [];
  const packs = createWorkerFontPacks((message) => requests.push(message));
  const timer = spyOn(globalThis, 'setTimeout');
  try {
    const first = packs.source('intl');
    const [expire, delay] = timer.mock.calls.at(-1)!;
    if (typeof expire !== 'function') throw new Error('Expected a timer callback');
    expire();
    expect(await first).toBeUndefined();
    expect(delay).toBe(15_000);
    const retry = packs.source('intl');
    expect(requests).toHaveLength(2);
    packs.receive({ protocol: 1, kind: 'font-pack-reply', id: requests[0]!.id });
    const urls = { 'intl.zst': 'https://fonts.example/intl.zst' };
    packs.receive({ protocol: 1, kind: 'font-pack-reply', id: requests[1]!.id, urls });
    expect(await retry).toEqual(urls);
  } finally {
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
  expect(await missing).toBeUndefined();
  const installed = packs.source('cjk');
  const urls = { 'cjk.zst': 'https://fonts.example/cjk.zst' };
  reply({ urls });
  expect(await installed).toEqual(urls);
  expect(packs.source('cjk')).toBe(installed);
  expect(requests).toHaveLength(4);
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
