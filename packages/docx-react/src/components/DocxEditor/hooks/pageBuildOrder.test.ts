import { describe, expect, test } from 'bun:test';
import { nearestPages } from './pageBuildOrder';

function sortedNearestPages(
  unbuilt: readonly number[],
  start: number,
  end: number,
  count: number
): number[] {
  const distance = (index: number) => (index < start ? start - index : index - end + 1);
  return [...unbuilt]
    .sort((a, b) => distance(a) - distance(b))
    .slice(0, count)
    .sort((a, b) => a - b);
}

describe('nearestPages', () => {
  test('returns no pages for empty input or a nonpositive count', () => {
    expect(nearestPages([], 3, 6, 16)).toEqual([]);
    expect(nearestPages([0, 1, 8], 3, 6, 0)).toEqual([]);
    expect(nearestPages([0, 1, 8], 3, 6, -1)).toEqual([]);
  });

  test('prefers lower indexes on ties and returns ascending pages without mutating input', () => {
    const unbuilt = Object.freeze([1, 3, 5, 10, 12, 14]);
    expect(nearestPages(unbuilt, 7, 9, 1)).toEqual([5]);
    expect(nearestPages(unbuilt, 7, 9, 3)).toEqual([3, 5, 10]);
    expect(nearestPages(unbuilt, 7, 9, 20)).toEqual(unbuilt);
  });

  test('handles windows past either end of the input', () => {
    const unbuilt = [2, 5, 8, 11];
    expect(nearestPages(unbuilt, -4, -1, 2)).toEqual([2, 5]);
    expect(nearestPages(unbuilt, 14, 17, 2)).toEqual([8, 11]);
  });

  test('matches the previous sort order over seeded sparse inputs and windows', () => {
    let seed = 0x5eed;
    const random = () => {
      seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
      return seed / 0x100000000;
    };

    for (let trial = 0; trial < 200; trial += 1) {
      const pageCount = 20 + Math.floor(random() * 280);
      const width = 1 + Math.floor(random() * 10);
      const middle = Math.floor(random() * (pageCount - width));
      const pages = Array.from({ length: pageCount }, (_, index) => index).filter(
        () => random() < 0.7
      );
      const windows: [number, number][] = [
        [0, width],
        [middle, middle + width],
        [pageCount - width, pageCount],
        [-width, 0],
        [pageCount, pageCount + width],
      ];
      for (const [start, end] of windows) {
        const unbuilt = pages.filter((index) => index < start || index >= end);
        const counts = [0, 1, 16, 32, unbuilt.length + 5, Math.floor(random() * pageCount)];
        for (const count of counts) {
          expect(nearestPages(unbuilt, start, end, count)).toEqual(
            sortedNearestPages(unbuilt, start, end, count)
          );
        }
      }
    }
  });

  test('selects the nearest background batch from a large document', () => {
    const unbuilt = Array.from({ length: 200_000 }, (_, index) =>
      index < 100_000 ? index : index + 10
    );
    expect(nearestPages(unbuilt, 100_000, 100_010, 16)).toEqual([
      ...Array.from({ length: 8 }, (_, index) => 99_992 + index),
      ...Array.from({ length: 8 }, (_, index) => 100_010 + index),
    ]);
  });
});
