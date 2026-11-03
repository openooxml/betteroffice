import { expect, test } from 'bun:test';
import {
  MEASURE_BUDGET_SECONDS,
  combineShardReports,
  measureMatrix,
  measureSeconds,
  measureShard,
  measureShards,
} from './measure-shards.mjs';

type Sample = { id: string; format: string; metadata: { reference: { pages: number } } };

function plan(counts: Record<string, number[]>) {
  const samples: Sample[] = [];
  for (const [format, pages] of Object.entries(counts))
    pages.forEach((page, index) => samples.push({ id: `${format}-${String(index).padStart(5, '0')}`, format, metadata: { reference: { pages: page } } }));
  return { formats: Object.keys(counts), samples };
}

const repeat = (count: number, pages: (index: number) => number) => Array.from({ length: count }, (_, index) => pages(index));

test('the showcase collection keeps one measure job per format', () => {
  const showcase = plan({ docx: repeat(63, (i) => (i === 0 ? 178 : 8)), pptx: repeat(103, () => 10), xlsx: repeat(208, (i) => (i % 3 ? 1 : 2)) });
  expect(measureMatrix(showcase)).toEqual([{ format: 'docx', shard: 0 }, { format: 'pptx', shard: 0 }, { format: 'xlsx', shard: 0 }]);
});

test('a full corpus splits each format into balanced shards inside the budget', () => {
  const corpus = plan({ docx: repeat(1776, (i) => 1 + ((i * 37) % 23)), pptx: repeat(656, (i) => 5 + ((i * 11) % 30)), xlsx: repeat(3297, (i) => 1 + (i % 2)) });
  const shards = measureShards(corpus);
  for (const format of corpus.formats) {
    const own = shards.filter((shard) => shard.format === format);
    const samples = corpus.samples.filter((sample) => sample.format === format);
    const total = samples.reduce((sum, sample) => sum + measureSeconds(sample), 0);
    const largest = Math.max(...samples.map(measureSeconds));
    expect(own.length).toBe(Math.ceil(total / MEASURE_BUDGET_SECONDS));
    expect(own.map((shard) => shard.shard)).toEqual(own.map((_, index) => index));
    for (const shard of own) expect(shard.seconds).toBeLessThanOrEqual(total / own.length + largest);
    expect(own.flatMap((shard) => shard.samples).sort()).toEqual(samples.map((sample) => sample.id).sort());
  }
  expect(shards.length).toBeLessThan(256);
  expect(shards.filter((shard) => shard.format === 'docx').length).toBeGreaterThan(10);
});

test('a document longer than the budget gets a shard to itself, and shards never outnumber documents', () => {
  const long = plan({ docx: [1500, 2, 2] });
  const shards = measureShards(long);
  expect(shards.length).toBe(2);
  expect(shards[0].samples).toEqual(['docx-00000']);
  expect(measureShards(plan({ pptx: [2000, 2000] })).length).toBe(2);
});

test('the cost model rejects a sample without a usable page count', () => {
  for (const pages of [0, -1, 1.5, Number.NaN])
    expect(() => measureShards(plan({ docx: [pages] }))).toThrow('Invalid reference page count');
  expect(() => measureShard(plan({ docx: [1] }), 'docx', 1)).toThrow('No measure shard 1 for docx');
});

test('shard reports combine into one format report only when every planned sample is in its shard', () => {
  const corpus = plan({ docx: [1500, 2, 2] });
  const [first, second] = measureShards(corpus);
  const identity = { source_sha: 's', commit: 'c', versions: { docx: '1.0.0' }, react_version: '2.0.0', format: 'docx' };
  const report = (shard: number, ids: string[]) => ({ ...identity, shard, samples: ids.map((id) => ({ id })) });
  const combined = combineShardReports(corpus, 'docx', [report(0, first.samples), report(1, second.samples)]);
  expect(combined).toEqual({ ...identity, samples: [...first.samples, ...second.samples].map((id) => ({ id })) });
  expect(() => combineShardReports(corpus, 'docx', [report(0, first.samples)])).toThrow('Missing docx measure shards');
  expect(() => combineShardReports(corpus, 'docx', [report(1, first.samples), report(0, second.samples)])).toThrow('missing or mislabelled');
  expect(() => combineShardReports(corpus, 'docx', [report(0, second.samples), report(1, first.samples)])).toThrow('does not hold its planned samples');
  expect(() => combineShardReports(corpus, 'docx', [report(0, first.samples), { ...report(1, second.samples), commit: 'other' }])).toThrow('disagree on commit');
});

test('a plan needing more measure shards than one workflow matrix holds fails in planning, with the count', () => {
  const longest = plan({ docx: repeat(1235, () => 250) });
  expect(measureShards(longest).length).toBe(257);
  expect(() => measureMatrix(longest)).toThrow('Measurement needs 257 shards; one workflow matrix holds 256');
  expect(measureMatrix(plan({ docx: repeat(1230, () => 250) })).length).toBeLessThanOrEqual(256);
});
