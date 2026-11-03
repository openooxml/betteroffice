import { expect, test } from 'bun:test';
import { failureClusters, failureCounts, failureRecords, normalizeReason } from './failures.mjs';

const hash = 'ab'.repeat(32);

function sample(id: string, format = 'docx') {
  return {
    id,
    format,
    source_sha256: hash,
    metadata_url: `https://corpus.example/${id}/metadata.json`,
    comparisons: [
      { channel: 'published', status: 'failed', stage: 'capture', error: `Error: page.goto: net::ERR_FAILED at https://corpus.example/${id}/source.${format}` },
      { channel: 'commit', status: 'ok' },
    ],
    native_timings: {
      published: { status: 'failed', error: `/home/runner/work/_temp/native/published/${format}-native-bench exited 1: {"error":"unsupported shape field: effects","status":"failed"}\n` },
      commit: { status: 'ok' },
    },
    roundtrip: {
      published: { parse: { status: 'ok' }, roundtrip: { status: 'failed', stage: 'preserve', error: 'Package differs outside the exact intended edit' } },
      libreoffice: { parse: { status: 'failed', error: `cannot open ${id}.${format}` }, roundtrip: { status: 'failed', stage: 'edit', error: 'skipped' } },
    },
  };
}

test('every failed render, timing, calculation, parse, roundtrip and probe becomes one record', () => {
  const records = failureRecords({
    ...sample('alpha'),
    calculations: { libreoffice: { status: 'failed', error: 'Recalculation results changed between trials' } },
    roundtrip_probe: { status: 'unavailable', error: 'probe timed out' },
  });
  expect(records.map(({ area, channel, stage }) => [area, channel, stage])).toEqual([
    ['render', 'published', 'capture'],
    ['timing', 'published', 'timing'],
    ['recalc', 'libreoffice', 'calculation'],
    ['roundtrip', 'published', 'preserve'],
    ['parse', 'libreoffice', 'parse'],
    ['probe', null, 'probe'],
  ]);
});

test('reasons drop ids, hashes, URLs, paths and numbers', () => {
  expect(normalizeReason('/home/runner/work/_temp/native/commit/docx-native-bench exited 101: {"error":"renderer skipped 3 images","status":"failed"}'))
    .toBe('renderer skipped N images');
  expect(normalizeReason('/opt/x/pptx-native-bench exited 1: slide is 7680x4320px, past the 4096px limit\n')).toBe('<path> exited N: slide is NxNpx, past the Npx limit');
  expect(normalizeReason('cannot open secret-report-7.docx', ['secret-report-7'])).toBe('cannot open <sample>.docx');
  expect(normalizeReason(`fetch https://corpus.example/x/a.png failed for ${hash}`)).toBe('fetch <url> failed for <hash>');
  expect(normalizeReason('Range A12:H40 produced 2 pages; expected one at the recorded scale')).toBe('Range AN:HN produced N pages; expected one at the recorded scale');
  expect(normalizeReason('   ')).toBe('no message');
  expect(normalizeReason('x'.repeat(400)).length).toBe(160);
});

test('clusters group one reason across documents, largest first, and keep the ids', () => {
  const report = { samples: [sample('alpha'), sample('beta'), sample('gamma', 'pptx')] };
  const clusters = failureClusters(report);
  const capture = clusters.find((cluster) => cluster.format === 'docx' && cluster.stage === 'capture')!;
  expect(capture).toEqual({
    format: 'docx', area: 'render', channel: 'published', stage: 'capture',
    reason: 'Error: page.goto: net::ERR_FAILED at <url>', count: 2, samples: ['alpha', 'beta'],
  });
  const parse = clusters.find((cluster) => cluster.format === 'docx' && cluster.area === 'parse')!;
  expect(parse.reason).toBe('cannot open <sample>.docx');
  expect(clusters.map((cluster) => cluster.count)).toEqual([...clusters.map((cluster) => cluster.count)].sort((a, b) => b - a));
  expect(clusters.filter((cluster) => cluster.format === 'docx').length).toBe(4);
  const text = JSON.stringify(clusters.map(({ samples, ...cluster }) => cluster));
  for (const leak of ['alpha', 'beta', 'gamma', 'corpus.example', hash, '/home/runner']) expect(text.includes(leak) && leak).toBe(false);
});

test('counts per format carry area, channel and stage only', () => {
  const report = { samples: [sample('alpha'), sample('beta'), sample('gamma', 'pptx')] };
  const counts = failureCounts(report, 'docx');
  expect(counts).toEqual([
    { area: 'parse', channel: 'libreoffice', stage: 'parse', count: 2 },
    { area: 'render', channel: 'published', stage: 'capture', count: 2 },
    { area: 'roundtrip', channel: 'published', stage: 'preserve', count: 2 },
    { area: 'timing', channel: 'published', stage: 'timing', count: 2 },
  ]);
  expect(failureCounts(report, 'xlsx')).toEqual([]);
});

test('a document without a usable probe counts once, and a usable probe counts nothing', () => {
  const unavailable = failureRecords({ id: 'x', format: 'xlsx', roundtrip_probe: { status: 'unavailable', error: 'No unambiguous plain-text or numeric-literal edit candidate' } });
  expect(unavailable).toEqual([{ area: 'probe', channel: null, stage: 'probe', error: 'No unambiguous plain-text or numeric-literal edit candidate' }]);
  expect(failureRecords({ id: 'y', format: 'xlsx', roundtrip_probe: { status: 'ok', part: 'xl/worksheets/sheet1.xml' } })).toEqual([]);
});

test('long messages normalize in linear time', () => {
  for (const message of ['a'.repeat(200000), 'ab://'.repeat(40000), '/'.repeat(200000), `${'f'.repeat(100000)} ${'1'.repeat(100000)}`]) {
    const started = performance.now();
    expect(normalizeReason(message).length).toBeLessThanOrEqual(160);
    expect(performance.now() - started).toBeLessThan(200);
  }
});

test('a failure is scrubbed of its own document even when the plan holds thousands', () => {
  const samples = Array.from({ length: 6000 }, (_, index) => ({
    ...sample(`doc-${index}`), comparisons: [{ channel: 'published', status: 'failed', stage: 'capture', error: `timeout opening doc-${index}.docx` }],
  }));
  const started = performance.now();
  const clusters = failureClusters({ samples });
  expect(performance.now() - started).toBeLessThan(5000);
  expect(clusters.find((cluster) => cluster.area === 'render')).toMatchObject({ reason: 'timeout opening <sample>.docx', count: 6000 });
});
