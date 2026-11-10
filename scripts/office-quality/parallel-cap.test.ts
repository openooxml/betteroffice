import { expect, test } from 'bun:test';
import { measureShards } from './measure-shards.mjs';
import { parallelPlan } from './parallel.mjs';
import { roundtripMatrices } from './roundtrip.mjs';
import { jobShards } from './shards.mjs';
import { timingFormats } from './timing.mjs';

const sample = (id: string, format: string, pages: number) => ({ id, format, metadata: { source: { sha256: id }, reference: { pages } } });

test('a private run never gives a matrix more concurrent slots than it has jobs', () => {
  const plan = {
    formats: ['docx', 'pptx'],
    samples: [
      ...Array.from({ length: 8 }, (_, index) => sample(`d${index}`, 'docx', 1 + index * 3)),
      ...Array.from({ length: 4 }, (_, index) => sample(`p${index}`, 'pptx', 5 + index * 10)),
    ],
  };
  const [roundtripA, roundtripB] = roundtripMatrices(plan);
  const jobs: Record<string, number> = {
    measure: measureShards(plan).length,
    'docx-benchmark': jobShards('docx-benchmark', plan).length,
    'xlsx-benchmark': jobShards('xlsx-benchmark', plan).length,
    'pptx-benchmark': jobShards('pptx-benchmark', plan).length,
    'xlsx-fidelity': jobShards('xlsx-fidelity', plan).length,
    'roundtrip-a': roundtripA.length,
    'roundtrip-b': roundtripB.length,
    timing: timingFormats(plan).length,
  };
  const share = parallelPlan(plan, { privateRun: true, total: 12 });
  for (const name of Object.keys(jobs).filter((name) => jobs[name] > 0))
    expect({ name, share: share[name] <= jobs[name] }).toEqual({ name, share: true });
});
