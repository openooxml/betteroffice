import { expect, test } from 'bun:test';
import { METHOD, ROUNDS, SAMPLE_SIZE, SEED, mergeTiming, pairedRatio, summarizeTiming, timingFormats, timingSamples } from './timing.mjs';

const source = 'a'.repeat(40);
const published = 'b'.repeat(40);
const hash = 'c'.repeat(64);
const sample = (format: string, index: number, calculation = 'ok') => ({
  id: `${format}-${index}`, format,
  metadata: { source: { sha256: `${format}-${index}-sha` }, reference: { pages: 1 }, reference_pages: [{}],
    ...(format === 'xlsx' ? { calculation: { status: calculation } } : {}) },
});
const plan = (counts: Record<string, number>) => ({
  source_sha: source, formats: Object.keys(counts),
  versions: { docx: '1.0.0', pptx: '1.0.0', xlsx: '1.0.0' },
  docx_published_source_sha: published, pptx_published_source_sha: published, xlsx_published_source_sha: published,
  samples: Object.entries(counts).flatMap(([format, count]) =>
    Array.from({ length: count }, (_, index) => sample(format, index, index % 3 ? 'ok' : 'not_applicable'))),
});
const ok = (...elapsed: number[]) => ({ status: 'ok', elapsed_ms: elapsed });
const part = (input: any, format: string, timings = (_: any, index: number) => ({
  published: ok(10 + index, 11 + index, 12 + index), commit: ok(10, 10, 10), libreoffice: ok(40, 40, 40),
})) => ({
  schema_version: 1, plan_sha256: hash, format,
  benchmark: { method: METHOD, rounds: ROUNDS, seed: SEED, sample_size: SAMPLE_SIZE, published_version: '1.0.0',
    libreoffice_version: '26.2.3.2', builds: { published: { source_sha: published }, commit: { source_sha: source } } },
  runner: { processors: 4 },
  samples: timingSamples(input, format).map((entry: any, index: number) =>
    ({ id: entry.id, source_sha256: entry.metadata.source.sha256, timings: timings(entry, index) })),
});

test('the sample is seeded, capped, independent of corpus order, and XLSX takes only workbooks with a calculation reference', () => {
  const input = plan({ docx: 450, pptx: 30, xlsx: 600 });
  const docx = timingSamples(input, 'docx').map((entry: any) => entry.id);
  expect(docx).toHaveLength(SAMPLE_SIZE);
  expect(new Set(docx).size).toBe(SAMPLE_SIZE);
  expect(timingSamples({ ...input, samples: [...input.samples].reverse() }, 'docx').map((entry: any) => entry.id)).toEqual(docx);
  expect(docx).not.toEqual(input.samples.filter((entry) => entry.format === 'docx').slice(0, SAMPLE_SIZE).map((entry) => entry.id));
  expect(timingSamples(input, 'pptx')).toHaveLength(30);
  const xlsx = timingSamples(input, 'xlsx');
  expect(xlsx).toHaveLength(SAMPLE_SIZE);
  expect(xlsx.every((entry: any) => entry.metadata.calculation.status === 'ok')).toBe(true);
  expect(timingFormats(input)).toEqual(['docx', 'pptx', 'xlsx']);
  expect(timingFormats(plan({ docx: 3, xlsx: 1 }))).toEqual(['docx']);
});

test('the summary reports the mean per document, medians, per-round spread and the paired ratio', () => {
  const input = plan({ pptx: 4 });
  const merged = mergeTiming(input, {}, [part(input, 'pptx')], hash);
  const summary = merged.timing_benchmark.pptx.summary;
  expect(summary.total).toBe(4);
  expect(summary.common).toBe(4);
  expect(summary.channels.commit.mean_ms).toBe(10);
  expect(summary.channels.libreoffice.median_ms).toBe(40);
  expect(summary.channels.published.mean_ms).toBe(12.5);
  expect(summary.channels.published.round_means_ms).toEqual([11.5, 12.5, 13.5]);
  expect(summary.channels.published.round_spread).toBeCloseTo(2 / 12.5, 10);
  expect(summary.ratio_to_libreoffice.commit).toEqual({ geometric_mean: 0.25, ci95: [0.25, 0.25] });
  const { geometric_mean, ci95 } = summary.ratio_to_libreoffice.published;
  expect(ci95[0]).toBeLessThanOrEqual(geometric_mean);
  expect(geometric_mean).toBeLessThanOrEqual(ci95[1]);
  expect(mergeTiming(input, {}, [part(input, 'pptx')], hash).timing_benchmark.pptx.summary).toEqual(summary);
});

test('a document that fails on any engine leaves the means but stays in the counts', () => {
  const input = plan({ docx: 3 });
  const failing = part(input, 'docx', (_: any, index: number) => ({
    published: ok(10, 10, 10), commit: index ? ok(20, 20, 20) : { status: 'failed', error: 'round 2: timed out' },
    libreoffice: ok(40, 40, 40),
  }));
  const summary = summarizeTiming('docx', failing.samples);
  expect(summary.common).toBe(2);
  expect(summary.channels.commit.successful).toBe(2);
  expect(summary.channels.published.successful).toBe(3);
  expect(summary.channels.commit.mean_ms).toBe(20);
  const none = summarizeTiming('docx', failing.samples.slice(0, 1));
  expect(none.channels.commit.mean_ms).toBeNull();
  expect(none.ratio_to_libreoffice.commit).toBeNull();
});

test('the merge refuses reports that do not match the plan, the method or the report revision', () => {
  const input = plan({ docx: 3, xlsx: 3 });
  const parts = () => [part(input, 'docx'), part(input, 'xlsx')];
  expect(Object.keys(mergeTiming(input, {}, parts(), hash).timing_benchmark)).toEqual(['docx', 'xlsx']);
  expect(() => mergeTiming(input, {}, parts().slice(1), hash)).toThrow();
  for (const mutate of [
    (p: any) => p[0].plan_sha256 = 'd'.repeat(64),
    (p: any) => p[0].benchmark.rounds = 2,
    (p: any) => p[0].benchmark.method = 'parallel',
    (p: any) => p[0].benchmark.seed = 'other',
    (p: any) => p[0].benchmark.sample_size = 10,
    (p: any) => p[0].benchmark.builds.commit.source_sha = published,
    (p: any) => p[0].benchmark.builds.published.source_sha = source,
    (p: any) => p[0].benchmark.libreoffice_version = 'latest',
    (p: any) => p[0].runner.processors = 0,
    (p: any) => p[1].format = 'docx',
    (p: any) => p[0].samples.pop(),
    (p: any) => p[0].samples.reverse(),
    (p: any) => p[0].samples[0].source_sha256 = hash,
    (p: any) => p[0].samples[0].timings.commit.elapsed_ms = [10, 10],
    (p: any) => p[0].samples[0].timings.commit.elapsed_ms = [10, 0, 10],
    (p: any) => p[0].samples[0].timings.commit = { status: 'failed', error: 'x', elapsed_ms: [1, 1, 1] },
    (p: any) => p[0].samples[0].timings.commit = { status: 'failed', error: ' ' },
    (p: any) => delete p[0].samples[0].timings.libreoffice,
  ]) {
    const changed = parts();
    mutate(changed);
    expect(() => mergeTiming(input, {}, changed, hash)).toThrow('Invalid timing benchmark');
  }
  const office = (change: any = {}) => ({ libreoffice_version: '26.2.3.2', libreoffice_build: 'LibreOffice 26.2.3.2',
    builds: { published: { source_sha: published }, commit: { source_sha: source } }, ...change });
  const built = () => parts().map((entry: any) => ({ ...entry, benchmark: { ...entry.benchmark, libreoffice_build: 'LibreOffice 26.2.3.2' } }));
  expect(Object.keys(mergeTiming(input, { docx_benchmark: office(), xlsx_benchmark: office() }, built(), hash).timing_benchmark)).toEqual(['docx', 'xlsx']);
  expect(() => mergeTiming(input, { docx_benchmark: office({ libreoffice_version: '25.8.1.1' }) }, built(), hash))
    .toThrow('LibreOffice version differs');
  expect(() => mergeTiming(input, { docx_benchmark: office({ libreoffice_build: 'LibreOffice 26.2.3.2 other' }) }, built(), hash))
    .toThrow('LibreOffice version differs');
  expect(() => mergeTiming(input, { xlsx_benchmark: office({ builds: { published: { source_sha: published },
    commit: { source_sha: source, binary_sha256: hash } } }) }, built(), hash)).toThrow('native builds differ');
});

test('a format with fewer documents than the cap still reconciles', () => {
  const input = plan({ pptx: 7 });
  const merged = mergeTiming(input, {}, [part(input, 'pptx')], hash);
  expect(merged.timing_benchmark.pptx.sample_size).toBe(SAMPLE_SIZE);
  expect(merged.timing_benchmark.pptx.summary.total).toBe(7);
});

test('the paired ratio interval is reproducible and brackets the estimate', () => {
  const ours = Array.from({ length: 50 }, (_, index) => 100 + (index % 7) * 10);
  const theirs = Array.from({ length: 50 }, (_, index) => 300 + (index % 5) * 40);
  const first = pairedRatio(ours, theirs, 'seed');
  expect(pairedRatio(ours, theirs, 'seed')).toEqual(first);
  expect(first!.ci95[0]).toBeLessThan(first!.geometric_mean);
  expect(first!.geometric_mean).toBeLessThan(first!.ci95[1]);
  expect(pairedRatio([], [], 'seed')).toBeNull();
});
