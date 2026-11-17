import { expect, test } from 'bun:test';
import { docxShards } from './docx-benchmark.mjs';
import { xlsxFidelityShards } from './xlsx-fidelity.mjs';

const BUDGET_SECONDS = 60 * 60;
const OBSERVED = {
  docx: (shard: any[]) => shard.reduce((sum, entry) => sum + 8.9 + 0.47 * entry.metadata.reference.pages, 0),
  pptx: (shard: any[]) => Math.max(18 * shard.length, 1.79 * shard.reduce((sum, entry) => sum + entry.metadata.reference.pages, 0)),
  xlsx: (shard: any[]) => Math.max(1.88 * shard.length, 1.14 * shard.reduce((sum, entry) => sum + entry.metadata.reference.pages, 0)),
};

const sample = (id: string, format: string, pages: number) => ({ id, format, metadata: { reference: { pages } } });
const corpus = () => ({
  formats: ['docx', 'pptx', 'xlsx'],
  samples: [
    ...Array.from({ length: 1776 }, (_, index) => sample(`d${index}`, 'docx', 1 + ((index * 37) % 22))),
    ...Array.from({ length: 656 }, (_, index) => sample(`p${index}`, 'pptx', 1 + ((index * 13) % 37))),
    ...Array.from({ length: 3297 }, (_, index) => sample(`x${index}`, 'xlsx', 1 + (index % 2))),
  ],
});

async function pptxShards(plan: any) {
  const module: any = await import('./pptx-benchmark.mjs');
  return module.pptxShards ? module.pptxShards(plan) : [plan.samples.filter((entry: any) => entry.format === 'pptx').map((entry: any) => entry.id)];
}

test('on a full-size corpus every benchmark shard stays inside its time budget at the rates the showcase run measured', async () => {
  const plan = corpus();
  const byId = new Map(plan.samples.map((entry) => [entry.id, entry]));
  const longest = (shards: string[][], cost: (shard: any[]) => number) =>
    Math.max(...shards.map((ids) => cost(ids.map((id) => byId.get(id)))));
  const planned = {
    docx: longest(docxShards(plan), OBSERVED.docx),
    pptx: longest(await pptxShards(plan), OBSERVED.pptx),
    xlsx: longest(xlsxFidelityShards(plan), OBSERVED.xlsx),
  };
  expect(Object.fromEntries(Object.entries(planned).map(([format, seconds]) => [format, seconds <= BUDGET_SECONDS])))
    .toEqual({ docx: true, pptx: true, xlsx: true });
});
