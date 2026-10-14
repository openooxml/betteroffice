import { expect, test } from 'bun:test';
import { CHANNELS, METHOD, TRIALS, docxShards, mergeDocxBenchmarks, timingSummary } from './docx-benchmark.mjs';
import { nativeFeatures } from './build-native.mjs';
import { renderSection } from './readme.mjs';

function fixture() {
  const current = 'a'.repeat(40);
  const published = 'b'.repeat(40);
  const hash = 'c'.repeat(64);
  const samples = [10, 5, 3, 2, 1, 1].map((pages, index) => ({
    id: `doc-${index}`, format: 'docx', metadata: {
      source: { sha256: hash }, reference: { pages },
    },
  }));
  const plan = { source_sha: current, commit: current, docx_published_source_sha: published,
    versions: { docx: '0.2.1', pptx: '0.2.0', xlsx: '0.2.0' }, samples };
  const build = (source_sha: string) => ({ source_sha, binary_sha256: hash, harness_sha256: hash,
    rustc: 'rustc 1.98.0', profile: 'release opt-level=3 lto=thin' });
  const benchmark = { method: METHOD, trials: TRIALS, dpi: 96, libreoffice_version: '26.2.3.2',
    published_version: '0.2.1', fonts_sha256: hash, builds: { published: build(published), commit: build(current) } };
  const parts = docxShards(plan).map((ids, shard) => ({
    schema_version: 1, shard, plan_sha256: hash, benchmark: structuredClone(benchmark), runner: {},
    samples: ids.map((id) => ({ id, source_sha256: hash, libreoffice: {
      channel: 'libreoffice', version: benchmark.libreoffice_version, status: 'ok',
      source_verified: true, reference: { status: 'ok', sha256: hash }, actual: { status: 'ok', sha256: hash },
      reference_pages: samples.find((sample) => sample.id === id)!.metadata.reference.pages,
      actual_pages: 2, penalized_ssim: 0.7, resized: false,
    }, timings: Object.fromEntries(CHANNELS.map((channel, index) => [channel,
      { status: 'ok', elapsed_ms: Array.from({ length: TRIALS }, () => (index + 1) * 100) }])) })),
  }));
  const report = { ...plan, samples: samples.map((sample) => ({ ...sample, comparisons: [] })) };
  return { plan, parts, report, hash };
}

test('balances whole DOCX documents and assigns every sample exactly once', () => {
  const { plan } = fixture();
  plan.samples.push({ ...plan.samples[0], id: 'slides', format: 'pptx' });
  const shards = docxShards(plan);
  expect(shards).toHaveLength(4);
  expect(shards.flat().sort()).toEqual(['doc-0', 'doc-1', 'doc-2', 'doc-3', 'doc-4', 'doc-5']);
  expect(shards[0]).toEqual(['doc-0']);
  expect(docxShards({ samples: [] })).toEqual([]);
  expect(docxShards({ samples: plan.samples.slice(0, 1) })).toEqual([['doc-0']]);
});

test('reconciles all shards and renders three versioned columns with native means last', () => {
  const { plan, report, parts, hash } = fixture();
  const merged = mergeDocxBenchmarks(plan, report, parts.reverse(), hash);
  expect(merged.samples.map((sample) => sample.id)).toEqual(plan.samples.map((sample) => sample.id));
  expect(timingSummary(merged.samples)).toMatchObject({ total: 6, common: 6,
    channels: { published: { mean_ms: 100 }, commit: { mean_ms: 200 }, libreoffice: { mean_ms: 300 } } });
  const text = renderSection(merged);
  expect(text).toContain('LibreOffice (26.2.3.2)');
  expect(text).toContain('BetterOffice (<a');
  expect(text).not.toContain('Latest published');
  expect(text).not.toContain('### Fidelity');
  expect(text).toContain('<tr><td>Render time (avg)</td><td align="right">100 ms</td><td align="right">200 ms</td><td align="right">300 ms</td></tr>\n</table>');
  expect(text).not.toContain('Timed/total');
  expect(text).not.toContain('Native CLI timing:');
  expect(text).toContain('[benchmark methodology](scripts/office-quality/README.md)');
  expect(() => renderSection({ ...merged, versions: { ...plan.versions, docx: '9.0.0' } })).toThrow('revision');
});

test('all averages use the common successful set, while coverage counts every engine', () => {
  const { plan, report, parts, hash } = fixture();
  const failed = parts[0].samples[0];
  failed.timings.published = { status: 'failed', error: 'unsupported shape' } as any;
  failed.timings.commit.elapsed_ms = Array(TRIALS).fill(9000);
  const merged = mergeDocxBenchmarks(plan, report, parts, hash);
  expect(timingSummary(merged.samples)).toMatchObject({ total: 6, common: 5,
    channels: { published: { successful: 5, mean_ms: 100 }, commit: { successful: 6, mean_ms: 200 } } });
  for (const sample of merged.samples)
    sample.native_timings.published = { status: 'failed', error: 'unsupported shape' };
  expect(timingSummary(merged.samples)).toMatchObject({ common: 0, channels: { commit: { mean_ms: null } } });
  expect(renderSection(merged)).toContain('<td>Render time (avg)</td><td align="right">—</td>');
});

test('rejects incomplete coverage, mixed builds, bad sources, and malformed measurements', () => {
  const mutations = [
    (f: any) => f.parts.pop(),
    (f: any) => f.parts[1] = f.parts[0],
    (f: any) => f.parts[0].samples.pop(),
    (f: any) => f.parts[0].samples[0].id = 'unknown',
    (f: any) => f.parts[0].samples[0].source_sha256 = 'd'.repeat(64),
    (f: any) => f.parts[0].plan_sha256 = 'd'.repeat(64),
    (f: any) => f.parts[0].benchmark.builds.commit.source_sha = 'd'.repeat(40),
    (f: any) => f.parts[0].benchmark.builds.published.harness_sha256 = 'd'.repeat(64),
    (f: any) => f.parts[0].benchmark.builds.published.rustc = 'other',
    (f: any) => f.parts[0].benchmark.libreoffice_version = '99.0.0',
    (f: any) => f.parts[0].benchmark.fonts_sha256 = 'd'.repeat(64),
    (f: any) => f.parts[0].benchmark.published_version = '9.0.0',
    (f: any) => f.parts[0].samples[0].timings.commit.elapsed_ms.pop(),
    (f: any) => f.parts[0].samples[0].timings.commit.elapsed_ms[0] = NaN,
    (f: any) => f.parts[0].samples[0].timings.commit.elapsed_ms[0] = -1,
    (f: any) => f.parts[0].samples[0].timings.commit.status = 'failed',
    (f: any) => delete f.parts[0].samples[0].timings.commit,
    (f: any) => f.parts[0].samples[0].libreoffice.reference_pages++,
    (f: any) => f.parts[0].samples[0].libreoffice.actual.sha256 = 'd'.repeat(64),
    (f: any) => f.parts[0].samples[0].libreoffice.status = 'pending',
  ];
  for (const mutate of mutations) {
    const f = fixture();
    mutate(f);
    expect(() => mergeDocxBenchmarks(f.plan, f.report, f.parts, f.hash)).toThrow();
  }
});

test('retains LibreOffice failures without inventing a score', () => {
  const f = fixture();
  f.parts[0].samples[0].libreoffice = { channel: 'libreoffice', version: '26.2.3.2',
    status: 'failed', stage: 'capture', error: 'export failed' } as any;
  const merged = mergeDocxBenchmarks(f.plan, f.report, f.parts, f.hash);
  expect(merged.samples[0].comparisons[0].status).toBe('failed');
  f.parts[0].samples[0].libreoffice.penalized_ssim = 0;
  expect(() => mergeDocxBenchmarks(f.plan, f.report, f.parts, f.hash)).toThrow();
});

test('builds older and current Rust APIs with the same native host', () => {
  expect(nativeFeatures('')).toEqual([]);
  expect(nativeFeatures('pub fn register_substitute_measure_font(id: u32, name: &str) {}')).toEqual(['substitute-metrics']);
  expect(nativeFeatures('pub fn register_substitute_measure_font(id: u32, name: &str, bold: bool, italic: bool) {}'))
    .toEqual(['substitute-styles']);
});
