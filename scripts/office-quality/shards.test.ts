import { readFile } from 'node:fs/promises';
import { expect, test } from 'bun:test';
import { JOB_TIMEOUT_MINUTES, SHARD_BUDGET_SECONDS, benchmarkShardPlan, budgetShards, shardSeconds, wallTimeLine } from './shards.mjs';
import { docxShards } from './docx-benchmark.mjs';
import { pptxShards } from './pptx-benchmark.mjs';
import { xlsxFidelityShards } from './xlsx-fidelity.mjs';
import { xlsxShards } from './xlsx-benchmark.mjs';

const shape = JSON.parse(await readFile(new URL('./fixtures/office-quality-shape.json', import.meta.url), 'utf8'));
const sample = (id: string, format: string, pages: number, cells: number | null = null) => ({
  id, format, metadata: { reference: { pages }, ...(cells ? { calculation: { status: 'ok', result_cells: cells } } : {}) },
});
const publicPlan = () => ({ formats: ['docx', 'pptx', 'xlsx'], samples: shape.samples.map(([id, format, pages, cells]: any) => sample(id, format, pages, cells)) });

function fixedShards(samples: any[], count: number, weight: (sample: any) => number) {
  const shards = Array.from({ length: Math.min(count, samples.length) }, () => [] as string[]);
  const weights = shards.map(() => 0);
  for (const entry of [...samples].sort((a, b) => weight(b) - weight(a) || a.id.localeCompare(b.id))) {
    const index = weights.indexOf(Math.min(...weights));
    shards[index].push(entry.id);
    weights[index] += weight(entry);
  }
  return shards;
}

test('the public collection gets exactly the shards it got before budgets', () => {
  const plan = publicPlan();
  const pages = (entry: any) => entry.metadata.reference.pages;
  const of = (format: string) => plan.samples.filter((entry: any) => entry.format === format);
  expect(docxShards(plan)).toEqual(fixedShards(of('docx'), 4, pages));
  expect(xlsxFidelityShards(plan)).toEqual(fixedShards(of('xlsx'), 4, pages));
  const calculated = of('xlsx').filter((entry: any) => entry.metadata.calculation);
  expect(xlsxShards(plan)).toEqual(fixedShards(calculated, 4, (entry) => entry.metadata.calculation.result_cells));
  expect(pptxShards(plan)).toEqual([of('pptx').map((entry: any) => entry.id)]);
  expect([docxShards(plan).length, pptxShards(plan).length, xlsxFidelityShards(plan).length, xlsxShards(plan).length]).toEqual([4, 1, 4, 4]);
});

test('a corpus-sized plan splits every benchmark job inside the budget and covers every document once', () => {
  const samples = [
    ...Array.from({ length: 1776 }, (_, index) => sample(`d${index}`, 'docx', 1 + ((index * 37) % 120))),
    ...Array.from({ length: 656 }, (_, index) => sample(`p${index}`, 'pptx', 1 + ((index * 13) % 70))),
    ...Array.from({ length: 3297 }, (_, index) => sample(`x${index}`, 'xlsx', 1 + (index % 3), index % 5 ? 50 + index * 7 : null)),
  ];
  for (const [job, format] of [['docx-benchmark', 'docx'], ['pptx-benchmark', 'pptx'], ['xlsx-fidelity', 'xlsx']] as const) {
    const selected = samples.filter((entry) => entry.format === format);
    const shards = budgetShards(job, selected);
    expect(shards.length).toBeGreaterThan(job === 'pptx-benchmark' ? 1 : 3);
    for (const shard of shards) expect(shard.seconds).toBeLessThanOrEqual(SHARD_BUDGET_SECONDS);
    expect(shards.flatMap((shard: any) => shard.samples).sort()).toEqual(selected.map((entry) => entry.id).sort());
  }
  const calculated = samples.filter((entry) => entry.metadata.calculation);
  for (const shard of budgetShards('xlsx-benchmark', calculated)) expect(shard.seconds).toBeLessThanOrEqual(SHARD_BUDGET_SECONDS);
  const lines = benchmarkShardPlan({ samples });
  expect(lines).toHaveLength(4);
  for (const line of lines) expect(line).toContain(`job timeout ${JOB_TIMEOUT_MINUTES})`);
  expect(lines.join('\n')).not.toMatch(/\b[dpx]\d+\b/);
  expect(SHARD_BUDGET_SECONDS).toBeLessThan(JOB_TIMEOUT_MINUTES * 60 * 0.75);
});

test('a document that alone exceeds the budget fails in planning, not at the job timeout', () => {
  expect(() => budgetShards('docx-benchmark', [sample('big', 'docx', 10)], 5)).toThrow('over the 5 s shard budget');
  expect(() => budgetShards('xlsx-benchmark', [sample('cells', 'xlsx', 1, null)])).toThrow('Invalid calculation reference size');
  expect(() => shardSeconds('docx-benchmark', sample('none', 'docx', 0))).toThrow('Invalid reference page count');
  expect(() => xlsxShards({ samples: [sample('huge', 'xlsx', 1, 250001)] })).toThrow('Invalid calculation reference size');
  expect(budgetShards('pptx-benchmark', [])).toEqual([]);
});

test('the shard count grows until the heaviest shard fits, never beyond one document per shard', () => {
  const samples = Array.from({ length: 6 }, (_, index) => sample(`s${index}`, 'pptx', 100));
  const seconds = shardSeconds('pptx-benchmark', samples[0]);
  expect(budgetShards('pptx-benchmark', samples, seconds * 2)).toHaveLength(3);
  expect(budgetShards('pptx-benchmark', samples, seconds)).toHaveLength(6);
  expect(budgetShards('pptx-benchmark', samples, seconds * 6)).toHaveLength(1);
});

test('twice the slides per deck gives about twice the PPTX shards, each kept in collection order', () => {
  const decks = (slides: number) => ({ formats: ['pptx'], samples: Array.from({ length: 656 }, (_, index) => sample(`p${String(index).padStart(3, '0')}`, 'pptx', slides)) });
  const single = pptxShards(decks(19));
  const double = pptxShards(decks(38));
  expect(double.length / single.length).toBeGreaterThan(1.6);
  expect(double.length / single.length).toBeLessThan(2.4);
  for (const shard of double) expect(shard).toEqual([...shard].sort());
});

test('the wall-time line names the shard, its planned minutes and the actual minutes, and no document', () => {
  const plan = { samples: [sample('deck-a', 'pptx', 14), sample('deck-b', 'pptx', 26)] };
  const line = wallTimeLine('pptx-benchmark', plan, 0, 1000, 1000 + 15 * 60);
  expect(line).toBe('pptx-benchmark shard 0: planned 1.2 min of work, job wall time 15.0 min (job timeout 90 min)');
  expect(() => wallTimeLine('pptx-benchmark', plan, 3, 0, 1)).toThrow('No pptx-benchmark shard 3');
});
