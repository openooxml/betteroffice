import { expect, test } from 'bun:test';
import { residentMediaImages } from './residentMediaImages';
import { residentSvgSources } from './residentSource';
import type { Document } from '../types/document';

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

test('a parsed document names its SVG pictures by data URL, once each', () => {
  const svg = 'data:image/svg+xml;base64,PHN2Zy8+';
  const document = { package: {
    document: { content: [{ src: svg }, { nested: [{ src: svg }, { src: 'data:image/png;base64,AA==' }] }] },
    headers: new Map([['rId1', { content: [{ src: 'data:image/svg+xml;utf8,<svg/>' }] }]]),
  } } as unknown as Document;
  expect(residentSvgSources(document)).toEqual([svg, 'data:image/svg+xml;utf8,<svg/>']);
});
