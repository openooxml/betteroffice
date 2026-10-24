import { expect, test } from 'bun:test';
import { residentMediaImages } from './residentMediaImages';

test('host-decoded pictures resolve without decoding and survive close', async () => {
  let closed = 0;
  const hosted = { close: () => { closed += 1; } } as unknown as ImageBitmap;
  let reads = 0;
  const images = residentMediaImages({ mediaDataUrl: () => { reads += 1; return null; } }, new Map([['media:0', hosted]]));
  expect(await images.resolve('media:0')).toBe(hosted);
  expect(reads).toBe(0);
  expect(await images.resolve('media:1')).toBeNull();
  expect(reads).toBe(1);
  await images.close();
  expect(closed).toBe(0);
});
