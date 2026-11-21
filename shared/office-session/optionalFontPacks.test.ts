import { expect, mock, test } from 'bun:test';
import { createFontPackResponder, createWorkerFontPacks, installWorkerFontPacks, type OptionalFontPack } from './optionalFontPacks';
import { withFontPackSource } from './fontPackSource';
import type { SessionTransport } from './transport';

function channel(): [SessionTransport, SessionTransport] {
  const listeners = [new Set<(message: unknown) => void>(), new Set<(message: unknown) => void>()];
  const endpoint = (side: number): SessionTransport => ({
    post(message) { for (const listener of [...listeners[1 - side]!]) listener(message); },
    listen(listener) { listeners[side]!.add(listener); return () => { listeners[side]!.delete(listener); }; },
    onError() { return () => {}; },
    close() { listeners[side]!.clear(); },
  });
  return [endpoint(0), endpoint(1)];
}

test('sessions import optional packs only on request and share concurrent and repeated imports', async () => {
  const [page, worker] = channel();
  const imports = mock(async (pack: OptionalFontPack) => ({ [`${pack}.zst`]: `https://fonts.example/${pack}.zst` }));
  const client = withFontPackSource(page, imports);
  let source!: ReturnType<typeof createWorkerFontPacks>['source'];
  const scope = Object.assign(new EventTarget(), { postMessage: (message: unknown) => worker.post(message) });
  const unlisten = worker.listen((message) => scope.dispatchEvent(new MessageEvent('message', { data: message })));
  const dispose = installWorkerFontPacks(scope, (installed) => { source = installed; });
  const clientMessages: unknown[] = [];
  const hostMessages: unknown[] = [];
  client.listen((message) => clientMessages.push(message));
  scope.addEventListener('message', (event) => hostMessages.push((event as MessageEvent).data));
  await Promise.resolve();
  expect(imports).not.toHaveBeenCalled();
  expect(await Promise.all([source('intl'), source('intl')])).toEqual([
    { 'intl.zst': 'https://fonts.example/intl.zst' }, { 'intl.zst': 'https://fonts.example/intl.zst' },
  ]);
  expect(imports).toHaveBeenCalledTimes(1);
  await source('intl');
  expect(imports).toHaveBeenCalledTimes(1);
  await source('cjk');
  await source('cjk');
  expect(imports).toHaveBeenCalledTimes(2);
  expect(clientMessages).toEqual([]);
  expect(hostMessages).toEqual([]);
  client.close();
  dispose();
  unlisten();
  worker.close();
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

test('the page caches missing packs as well as installed packs', async () => {
  const imports = mock(async () => undefined);
  const replies: unknown[] = [];
  const responder = createFontPackResponder((reply) => replies.push(reply), imports);
  for (const id of ['first', 'second']) responder.receive({ protocol: 1, kind: 'font-pack-request', id, pack: 'cjk' });
  await new Promise<void>((resolve) => setTimeout(resolve, 0));
  expect(imports).toHaveBeenCalledTimes(1);
  expect(replies).toEqual(['first', 'second'].map((id) => ({ protocol: 1, kind: 'font-pack-reply', id, urls: undefined })));
  responder.dispose();
});

test('a page that never answers cannot keep font loading pending', async () => {
  const posted: unknown[] = [];
  const packs = createWorkerFontPacks((message) => posted.push(message), 1);
  expect(await packs.source('intl')).toBeUndefined();
  expect(posted).toHaveLength(1);
  packs.dispose();
});

test('font pack replies correlate by unique id and preserve source failures', async () => {
  const requests: Array<{ id: string }> = [];
  const first = createWorkerFontPacks((message) => requests.push(message));
  const second = createWorkerFontPacks((message) => requests.push(message));
  const one = first.source('intl');
  const two = second.source('intl');
  expect(requests[0]!.id).not.toBe(requests[1]!.id);
  const reply = (id: string) => ({ protocol: 1, kind: 'font-pack-reply', id, error: 'Unreadable font pack' });
  first.receive(reply(requests[1]!.id));
  const rejected = expect(one).rejects.toMatchObject({ code: 'message', message: 'Unreadable font pack' });
  first.receive(reply(requests[0]!.id));
  await rejected;
  const disposed = expect(two).rejects.toMatchObject({ code: 'disposed' });
  second.dispose();
  await disposed;
  first.dispose();
});
