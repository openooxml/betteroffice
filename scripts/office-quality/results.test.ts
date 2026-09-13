import { expect, test } from 'bun:test';
import { measureSamples, summarizeError } from './results.mjs';
import { renderSection } from './readme.mjs';

const commit = 'a'.repeat(40);
const comparison = {
  source_verified: true,
  reference: { status: 'ok', sha256: 'source' },
  actual: { status: 'ok', sha256: 'source' },
  penalized_ssim: 0.8,
};

test('capture and comparison failures preserve later samples and the other channel', async () => {
  const samples = ['capture-fails', 'compare-fails', 'works'].map((id) => ({
    id,
    format: 'xlsx',
    comparisons: [] as Record<string, unknown>[],
  }));
  const calls: string[] = [];
  const logs: string[] = [];
  for (const channel of ['published', 'commit']) {
    await measureSamples(
      samples,
      { channel, version: '0.1.0', renderer_source_commit: commit },
      {
        capture: async (sample: { id: string }) => {
          calls.push(`${channel}:capture:${sample.id}`);
          if (channel === 'published' && sample.id === 'capture-fails')
            throw Object.assign(new Error('Command failed with local arguments'), {
              stdout: JSON.stringify({
                error:
                  'Error: viewport must have finite positive dimensions\n at /private/source.js:12',
              }),
            });
        },
        compare: async (sample: { id: string }) => {
          calls.push(`${channel}:compare:${sample.id}`);
          if (channel === 'published' && sample.id === 'compare-fails')
            throw Object.assign(new Error('Command failed'), {
              stderr:
                'usage: compare.py ...\ncompare.py: error: page 1: (100, 200) vs (110, 200)\n',
            });
          return comparison;
        },
        log: (message: string) => logs.push(message),
      }
    );
  }
  expect(calls).toEqual([
    'published:capture:capture-fails',
    'published:capture:compare-fails',
    'published:compare:compare-fails',
    'published:capture:works',
    'published:compare:works',
    'commit:capture:capture-fails',
    'commit:compare:capture-fails',
    'commit:capture:compare-fails',
    'commit:compare:compare-fails',
    'commit:capture:works',
    'commit:compare:works',
  ]);
  expect(samples[0].comparisons[0]).toMatchObject({
    status: 'failed',
    stage: 'capture',
  });
  expect(samples[1].comparisons[0]).toMatchObject({
    status: 'failed',
    stage: 'compare',
  });
  expect(samples[0].comparisons[0]).not.toHaveProperty('penalized_ssim');
  expect(samples[1].comparisons[0]).not.toHaveProperty('penalized_ssim');
  expect(samples.every((sample) => sample.comparisons[1].status === 'ok')).toBe(true);
  expect(logs[0]).toBe(
    'capture-fails published: FAILED (capture): Error: viewport must have finite positive dimensions'
  );
  const report = {
    commit,
    versions: { docx: '0.1.0', pptx: '0.1.0', xlsx: '0.1.0' },
    samples,
  };
  const section = renderSection(JSON.parse(JSON.stringify(report)));
  expect(section).toMatch(
    /\| XLSX .*\| 0\.8000 \| 1\/3 .*\| 0\.8000 \| 3\/3 \| 2 \/ 0 \|/
  );
  expect(section).toContain('| capture-fails | published | capture | Error: viewport');
  expect(section).toContain(
    '| compare-fails | published | compare | compare.py: error: page 1:'
  );
  expect(section).not.toContain('/private/');
});

test('mismatched sources remain unscored while valid later comparisons complete', async () => {
  const samples = ['mismatch', 'valid'].map((id) => ({
    id,
    comparisons: [] as Record<string, unknown>[],
  }));
  await measureSamples(
    samples,
    { channel: 'commit', renderer_source_commit: commit },
    {
      capture: async () => {},
      compare: async (sample: { id: string }) =>
        sample.id === 'mismatch'
          ? { ...comparison, actual: { status: 'ok', sha256: 'other' } }
          : comparison,
      log: () => {},
    }
  );
  expect(samples[0].comparisons[0]).toMatchObject({
    status: 'failed',
    stage: 'compare',
  });
  expect(samples[0].comparisons[0]).not.toHaveProperty('penalized_ssim');
  expect(samples[1].comparisons[0]).toMatchObject({
    status: 'ok',
    penalized_ssim: 0.8,
  });
});

test('failure summaries remove local paths, URLs, terminal escapes and stack traces', () => {
  expect(
    summarizeError(
      new Error(
        '\u001b[31mCannot read /Users/private/source.xlsx at https://example.com/file?token=secret\u001b[0m\n at run.js:1'
      )
    )
  ).toBe('Cannot read [path] at [url]');
  expect(summarizeError(new Error('Cannot read C:\\Users\\private\\source.xlsx'))).toBe(
    'Cannot read [path]'
  );
  expect(summarizeError('Unexpected external requests: [{"cookie":"private"}]')).toBe(
    'Unexpected external requests during capture'
  );
  expect(summarizeError('x'.repeat(500)).length).toBe(240);
});
