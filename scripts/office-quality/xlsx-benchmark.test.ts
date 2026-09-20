import { expect, test } from 'bun:test';
import { calculationSummary, mergeXlsxBenchmarks, xlsxShards, METHOD, ABS_TOLERANCE, REL_TOLERANCE } from './xlsx-benchmark.mjs';
import { createPlan } from './plan.mjs';

const source = 'a'.repeat(40);
const published = 'b'.repeat(40);
const hash = 'c'.repeat(64);
const reference = 'd'.repeat(64);
const samples = ['first', 'second'].map(id => ({ id, format: 'xlsx', metadata: {
  id, format: 'xlsx', source: { sha256: hash },
  reference: { status: 'ok', dpi: 150, sha256: hash, pages: 1 }, reference_pages: [{}],
  calculation: { status: 'ok', result_cells: 25, reference: { sha256: reference } },
} }));
const plan = () => createPlan({ source_sha: source, commit: source, versions: { docx: '1.0.0', pptx: '1.0.0', xlsx: '1.0.0' },
  react_version: '1.0.0', xlsx_published_source_sha: published, samples });
const success = (correct = 25, ms = 10) => ({ status: 'ok', correct, total: 25, elapsed_ms: Array(5).fill(ms) });
const build = (sha: string) => ({ source_sha: sha, binary_sha256: hash, harness_sha256: hash, rustc: 'same', profile: 'same' });
const parts = () => xlsxShards(plan()).map((ids: string[], shard: number) => ({ schema_version: 1, plan_sha256: hash, shard,
  benchmark: { method: METHOD, trials: 5, absolute_tolerance: ABS_TOLERANCE, relative_tolerance: REL_TOLERANCE,
    libreoffice_version: '26.2.3.2', published_version: '1.0.0', builds: { published: build(published), commit: build(source) } },
  samples: ids.map(id => ({ id, source_sha256: hash, reference_sha256: reference,
    calculations: { published: success(), commit: success(), libreoffice: success(25,20) } })),
}));

test('calculation shards cover every referenced workbook once and omit files without an oracle', () => {
  expect(xlsxShards(plan()).flat().sort()).toEqual(['first','second']);
  expect(xlsxShards({ samples: [{ format:'xlsx',metadata:{} }] })).toEqual([]);
});

test('failed calculations stay in accuracy totals and timings use only jointly correct workbooks', () => {
  const inputs = parts();
  (inputs[1].samples[0].calculations.published as any) = { status:'failed', correct:0,total:25,error:'engine failed' };
  inputs[1].samples[0].calculations.commit = success(10,1);
  const result = mergeXlsxBenchmarks(plan(), {samples}, inputs, hash);
  const summary = calculationSummary(result.samples);
  expect(summary.common).toBe(1);
  expect(summary.channels.published).toEqual({correct:25,total:50,mean_ms:10});
  expect(summary.channels.commit).toEqual({correct:35,total:50,mean_ms:10});
  expect(summary.channels.libreoffice).toEqual({correct:50,total:50,mean_ms:20});
  expect(result.xlsx_benchmark.reference_coverage).toEqual({total:2,measured:2,no_formulas:[],missing_reference:[]});
});

test('rejects missing shards, stale references, mixed harnesses and malformed trials', () => {
  expect(() => mergeXlsxBenchmarks(plan(),{samples},parts().slice(1),hash)).toThrow();
  for (const mutate of [
    (p: any) => p[1].shard = 0,
    (p: any) => p[0].samples[0].reference_sha256 = 'e'.repeat(64),
    (p: any) => p[0].benchmark.builds.commit.harness_sha256 = 'e'.repeat(64),
    (p: any) => p[0].samples[0].calculations.commit.elapsed_ms = [1],
    (p: any) => p[0].samples[0].calculations.commit.correct = 26,
  ]) {
    const inputs = parts();
    mutate(inputs);
    expect(() => mergeXlsxBenchmarks(plan(),{samples},inputs,hash)).toThrow();
  }
});

import { renderSection } from './readme.mjs';

test('renders calculation accuracy and timing without inventing LibreOffice XLSX fidelity', () => {
  const input = { ...plan(), samples: samples.map(sample => ({ ...sample, comparisons: [] })) };
  const merged = mergeXlsxBenchmarks(plan(),input,parts(),hash);
  const text = renderSection(merged).split('### XLSX')[1];
  expect(text).toContain('LibreOffice (26.2.3.2)');
  expect(text).toContain('100.00% (50/50)');
  expect(text).toContain('<td>Recalc time (avg)</td><td align="right">10 ms</td><td align="right">10 ms</td><td align="right">20 ms</td>');
  expect(text).toContain('<td>Scored/total</td><td align="right">0/2</td><td align="right">0/2</td><td align="right">—</td>');
  expect(() => renderSection({...merged,source_sha:'e'.repeat(40)})).toThrow('revision');
});
