import { expect, test } from 'bun:test';
import { cachedFontAsset } from './asset-cache';

test('concurrent consumers share a font download without sharing mutable ownership', async () => {
  const gate = Promise.withResolvers<ArrayBuffer>();
  let reads = 0;
  const read = () => { reads++; return gate.promise; };
  const first = cachedFontAsset('shared-font-test', read);
  const second = cachedFontAsset('shared-font-test', read);
  gate.resolve(new Uint8Array([1, 2, 3]).buffer);
  const [a, b] = await Promise.all([first, second]);
  expect(reads).toBe(1);
  structuredClone(a, { transfer: [a] });
  new Uint8Array(b)[0] = 99;
  const reused = await cachedFontAsset('shared-font-test', read);
  expect([...new Uint8Array(reused)]).toEqual([1, 2, 3]);
  expect(reads).toBe(2);
});

test('unrelated font reads complete while other resources remain pending', async () => {
  const gate = Promise.withResolvers<ArrayBuffer>();
  const pending = Array.from({ length: 12 }, (_, i) => cachedFontAsset(`pending-font-test-${i}`, () => gate.promise));
  let readyStarted = false;
  const ready = cachedFontAsset('ready-font-test', async () => {
    readyStarted = true;
    return new Uint8Array([42]).buffer;
  });
  try {
    await new Promise<void>(resolve => queueMicrotask(resolve));
    expect(readyStarted).toBe(true);
    expect([...new Uint8Array(await ready)]).toEqual([42]);
  } finally {
    gate.resolve(new ArrayBuffer(0));
    await Promise.all([...pending, ready]);
  }
});

test('a failed font read does not poison a later request for the same resource', async () => {
  await expect(cachedFontAsset('failed-font-test', async () => { throw new Error('network failure'); }))
    .rejects.toThrow('network failure');
  const retry = await cachedFontAsset('failed-font-test', async () => new Uint8Array([42]).buffer);
  expect([...new Uint8Array(retry)]).toEqual([42]);
});
