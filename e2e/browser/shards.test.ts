import { expect, test } from 'bun:test';
import { resolve } from 'node:path';

import { reportDurations } from '../../scripts/browser-shard-weights';
import { assignShards, browserFiles, parseShard, shardTestMatch, shardWeights } from './shards';

const files = browserFiles();
const weights = shardWeights();
const load = (shard: string[]) => shard.reduce((sum, file) => sum + (weights[file] ?? 0), 0);

test('every browser file runs in exactly one shard', () => {
  for (const total of [1, 2, 3, 4]) {
    const shards = assignShards(files, weights, total);
    expect(shards).toHaveLength(total);
    expect(shards.flat().sort()).toEqual(files);
  }
});

test('shards stay within one file of an even split of the measured time', () => {
  const heaviest = Math.max(...Object.values(weights));
  for (const total of [2, 3]) {
    const loads = assignShards(files, weights, total).map(load);
    const even = loads.reduce((sum, value) => sum + value, 0) / total;
    for (const value of loads) expect(Math.abs(value - even)).toBeLessThanOrEqual(heaviest);
  }
});

test('the assignment ignores the order files are listed in', () => {
  expect(assignShards([...files].reverse(), weights, 3)).toEqual(assignShards(files, weights, 3));
});

test('unmeasured files weigh the mean of the measured ones', () => {
  const shards = assignShards(['a', 'b', 'c', 'new'], { a: 10, b: 10, c: 40 }, 2);
  expect(shards).toEqual([['c'], ['a', 'b', 'new']]);
});

test('the shard spec selects the matching files and rejects malformed input', () => {
  expect(shardTestMatch(undefined)).toBeUndefined();
  const matched = [1, 2, 3].flatMap((current) => {
    const patterns = shardTestMatch(`${current}/3`)!;
    return files.filter((file) => patterns.some((pattern) => pattern.test(resolve(import.meta.dirname, '../..', file))));
  });
  expect(matched.sort()).toEqual(files);
  for (const spec of ['0/2', '3/2', '1', '1/0', 'a/b']) expect(() => parseShard(spec)).toThrow();
});

test('report durations are summed per repo file', () => {
  const rootDir = resolve(import.meta.dirname, 'plugin-host');
  const result = (duration: number) => ({ tests: [{ results: [{ duration }] }] });
  const durations = reportDurations({
    config: { rootDir },
    suites: [{ file: 'x.browser.ts', specs: [result(1500)], suites: [{ file: 'x.browser.ts', specs: [result(500)] }] }],
  });
  expect([...durations]).toEqual([['e2e/browser/plugin-host/x.browser.ts', 2]]);
});
