import { expect, test } from 'bun:test';
import {
  hasRender,
  pageName,
  parseLatest,
  parseReport,
  rank,
  referenceUrl,
  renderUrl,
  reportUrl,
  score,
} from './report';

const commit = 'a'.repeat(40);

const report = {
  commit,
  versions: { docx: '0.2.1' },
  samples: [
    {
      id: 'alpha',
      format: 'docx',
      comparisons: [
        {
          channel: 'published',
          version: '0.2.1',
          status: 'ok',
          penalized_ssim: 0.9,
          common_page_ssim: 0.91,
          reference_pages: 2,
          actual_pages: 2,
          pages: [],
        },
        {
          channel: 'commit',
          renderer_source_commit: commit,
          status: 'ok',
          penalized_ssim: 0.8,
          common_page_ssim: 0.85,
          reference_pages: 2,
          actual_pages: 3,
          pages: [
            { page: 2, ssim: 0.7, reference_size: [1275, 1650], actual_size: [1275, 1650] },
            { page: 1, ssim: 0.9, reference_size: [1275, 1650], actual_size: [1275, 1650] },
            { page: 3, ssim: null, reference_size: null, actual_size: [1275, 1650] },
          ],
        },
      ],
    },
    {
      id: 'beta',
      format: 'xlsx',
      comparisons: [
        {
          channel: 'commit',
          renderer_source_commit: commit,
          status: 'failed',
          stage: 'capture',
          error: 'viewport must have finite positive dimensions',
        },
      ],
    },
  ],
};

test('parsing normalizes both channels and orders pages', () => {
  const parsed = parseReport(report);
  expect(parsed.commit).toBe(commit);
  expect(parsed.versions.docx).toBe('0.2.1');
  const alpha = parsed.documents[0]!;
  expect(alpha.pages.map((page) => page.page)).toEqual([1, 2, 3]);
  expect(alpha.pages[2]).toEqual({ page: 3, ssim: null, width: 1275, height: 1650 });
  expect(alpha.commit?.status).toBe('ok');
  expect(score(alpha.commit)).toBe('0.8000');
  expect(score(alpha.published)).toBe('0.9000');
});

test('a failed comparison keeps its reason and never gains a score', () => {
  const beta = parseReport(report).documents[1]!;
  expect(beta.commit).toEqual({
    status: 'failed',
    revision: commit,
    stage: 'capture',
    error: 'viewport must have finite positive dimensions',
  });
  expect(score(beta.commit)).toBe('—');
  expect(score(beta.published)).toBe('—');
  expect(beta.pages).toEqual([]);
});

test('malformed reports are rejected rather than half-rendered', () => {
  expect(() => parseReport(null)).toThrow('Malformed fidelity report');
  expect(() => parseReport({ commit: 'main', samples: [] })).toThrow('Malformed fidelity report');
  expect(() => parseReport({ commit, samples: {} })).toThrow('Malformed fidelity report');
  expect(() =>
    parseReport({
      commit,
      samples: [{ id: 'a', format: 'docx', comparisons: [{ channel: 'commit', status: 'ok' }] }],
    })
  ).toThrow('Malformed fidelity report');
});

test('worst documents rank first and failures rank above every score', () => {
  const parsed = parseReport(report);
  expect(rank(parsed.documents).map((document) => document.id)).toEqual(['beta', 'alpha']);
});

test('page URLs are padded and scoped to their origin', () => {
  expect(pageName(7)).toBe('page_0007.png');
  expect(referenceUrl('oxi-en-legal-01', 12)).toBe(
    'https://corpus.betteroffice.dev/oxi-en-legal-01/reference/page_0012.png'
  );
  expect(renderUrl(commit, 'alpha', 1)).toBe(`/renders/${commit}/alpha/page_0001.png`);
  expect(reportUrl(commit)).toBe(`/renders/${commit}/report.json`);
  expect(() => renderUrl('HEAD', 'alpha', 1)).toThrow('full commit SHA');
  expect(() => reportUrl('HEAD')).toThrow('full commit SHA');
  expect(() => pageName(0)).toThrow('Invalid page');
});

test('the render manifest decides which overlays exist', () => {
  const latest = parseLatest({
    schema_version: 1,
    sha: commit,
    published_at_utc: '2026-09-18T00:00:00Z',
    samples: { alpha: 3, beta: 0, gamma: 'x' },
  });
  expect(latest.samples).toEqual({ alpha: 3 });
  expect(latest.publishedAt).toBe('2026-09-18T00:00:00Z');
  expect(hasRender(latest, 'alpha', 3)).toBe(true);
  expect(hasRender(latest, 'alpha', 4)).toBe(false);
  expect(hasRender(latest, 'beta', 1)).toBe(false);
  expect(hasRender(null, 'alpha', 1)).toBe(false);
  expect(() => parseLatest({ sha: 'nope' })).toThrow('Malformed render manifest');
});
