import { expect, test } from 'bun:test';

test('text_selection_import_without_segmenter', async () => {
  const original = Object.getOwnPropertyDescriptor(Intl, 'Segmenter');
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    Object.defineProperty(Intl, 'Segmenter', { configurable: true, value: undefined });
    const path = './textSelection.ts?lazy-segmenters';
    const helpers = await Promise.race([import(path), new Promise<never>((_, reject) => {
      timer = setTimeout(() => reject(new Error('text selection import timed out')), 5000);
    })]);
    expect(helpers.paragraphRangeAt('one\ntwo', 1)).toEqual({ start: 0, end: 3 });
    if (original) Object.defineProperty(Intl, 'Segmenter', original);
    expect(helpers.wordRangeAt('Hello, café world', 8)).toEqual({ start: 7, end: 11 });
    expect(helpers.wordBoundary('Hello brave world', 8, -1)).toBe(6);
    expect([...helpers.graphemeBoundaries('e\u0301😀a')]).toEqual([0, 2, 4, 5]);
  } finally {
    clearTimeout(timer);
    if (original) Object.defineProperty(Intl, 'Segmenter', original);
    else Reflect.deleteProperty(Intl, 'Segmenter');
  }
});
