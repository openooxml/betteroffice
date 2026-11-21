import { expect, test } from 'bun:test';
import { createBrowserFrameImages } from './browserFrameImages';


test('trim closes released decodes and pinned ones at their last release', async () => {
  const closed: string[] = [];
  const images = createBrowserFrameImages(async (bytes) => ({ width: 1, height: 1, close() { closed.push(String(bytes[0])); } }) as unknown as ImageBitmap);
  const media = new Map([['a', Uint8Array.of(1)], ['b', Uint8Array.of(2)]]);
  const released = images.resolve({ media }), pinned = images.resolve({ media });
  await released('a'); await pinned('b');
  released.release();
  expect(closed).toEqual([]);
  images.trim();
  expect(closed).toEqual(['1']);
  pinned.release();
  expect(closed).toEqual(['1', '2']);
  images.dispose();
  expect(closed).toEqual(['1', '2']);
});
