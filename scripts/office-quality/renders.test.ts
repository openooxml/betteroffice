import { expect, test } from 'bun:test';
import {
  commitComparison,
  latestManifest,
  planRenders,
  renderKey,
  reportKey,
} from './renders.mjs';

const commit = 'a'.repeat(40);
const other = 'b'.repeat(40);

function comparison(actualPages: number) {
  return {
    channel: 'commit',
    renderer_source_commit: commit,
    status: 'ok',
    actual_pages: actualPages,
    reference_pages: actualPages,
  };
}

test('render keys are content addressed and reject unusable identifiers', () => {
  expect(renderKey(commit, 'oxi-en-legal-01', 7)).toBe(
    `renders/${commit}/oxi-en-legal-01/page_0007.png`
  );
  expect(reportKey(commit)).toBe(`renders/${commit}/report.json`);
  expect(() => renderKey('abc', 'sample', 1)).toThrow('full commit SHA');
  expect(() => renderKey(commit, '../escape', 1)).toThrow('Invalid sample folder');
  expect(() => renderKey(commit, 'sample', 0)).toThrow('Invalid page');
  expect(() => reportKey('HEAD')).toThrow('full commit SHA');
});

test('only the measured commit channel is published', () => {
  const sample = {
    id: 'sample',
    comparisons: [
      { ...comparison(1), channel: 'published', version: '0.2.1' },
      { channel: 'commit', renderer_source_commit: other, status: 'ok' },
    ],
  };
  expect(commitComparison(sample, commit)).toBeNull();
  expect(commitComparison({ id: 'sample', comparisons: [] }, commit)).toBeNull();
  expect(
    commitComparison(
      {
        id: 'sample',
        comparisons: [
          {
            channel: 'commit',
            renderer_source_commit: commit,
            status: 'failed',
            stage: 'capture',
            error: 'boom',
          },
        ],
      },
      commit
    )
  ).toBeNull();
  expect(() =>
    commitComparison({ id: 'sample', comparisons: [comparison(1), comparison(2)] }, commit)
  ).toThrow('Duplicate commit comparison');
});

test('the plan covers every rendered page of every scored sample', () => {
  const report = {
    commit,
    samples: [
      { id: 'alpha', comparisons: [comparison(3)] },
      {
        id: 'beta',
        comparisons: [
          {
            channel: 'commit',
            renderer_source_commit: commit,
            status: 'failed',
            stage: 'compare',
            error: 'boom',
          },
        ],
      },
      { id: 'gamma', comparisons: [comparison(0)] },
    ],
  };
  const plan = planRenders(report);
  expect(plan.sha).toBe(commit);
  expect(plan.samples).toEqual({ alpha: 3 });
  expect(plan.uploads.map((upload: { key: string }) => upload.key)).toEqual([
    `renders/${commit}/alpha/page_0001.png`,
    `renders/${commit}/alpha/page_0002.png`,
    `renders/${commit}/alpha/page_0003.png`,
  ]);
  expect(plan.uploads[0]).toEqual({
    sample: 'alpha',
    page: 1,
    name: 'page_0001.png',
    key: `renders/${commit}/alpha/page_0001.png`,
  });
  expect(() => planRenders({ commit: 'main', samples: [] })).toThrow('full commit SHA');
});

test('the latest pointer records the report and the page count per sample', () => {
  const plan = planRenders({ commit, samples: [{ id: 'alpha', comparisons: [comparison(2)] }] });
  expect(latestManifest(plan, '2026-09-18T00:00:00Z')).toEqual({
    schema_version: 1,
    sha: commit,
    published_at_utc: '2026-09-18T00:00:00Z',
    report: `renders/${commit}/report.json`,
    samples: { alpha: 2 },
  });
});
