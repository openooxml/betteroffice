import { expect, test } from 'bun:test';
import { settledDisplayListJson } from './settledFrame';

test('a preview frame reads again until no conditional format is pending', () => {
  const frames = [
    { width: 1, height: 1, commands: [], conditionalPending: 4 },
    { width: 1, height: 1, commands: [], conditionalPending: 1 },
    { width: 1, height: 1, commands: [] },
  ].map((frame) => JSON.stringify(frame));
  let reads = 0;
  expect(settledDisplayListJson(() => frames[reads++]!)).toBe(frames[2]!);
  expect(reads).toBe(3);
});
