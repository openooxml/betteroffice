import { expect, test } from 'bun:test';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fetchCounts, parseMonthly, rewriteSnapshots, updateDownloadSnapshots } from './download-snapshot.mjs';

const README = [
  '<img src="https://betteroffice.dev/api/npm-downloads-badge" alt="npm downloads: 1/month (snapshot 2026-01-01)">',
  '<img src="https://betteroffice.dev/api/crates-downloads-badge" alt="crates.io downloads: 2/month (snapshot 2026-01-01)">',
  '<img src="https://betteroffice.dev/api/pypi-downloads-badge" alt="PyPI downloads: 3/month (snapshot 2026-01-01)">',
  '<img src="https://betteroffice.dev/api/badge" alt="license">'
].join('\n');

const answer = (counts: Record<string, string>) => async (url: string) => {
  const key = Object.keys(counts).find((name) => url.endsWith(`/${name}-downloads`));
  return key ? Response.json({ message: counts[key] }) : new Response('no', { status: 503 });
};

test('reads the badge message as a count', () => {
  expect(parseMonthly('484,591/month')).toBe(484591);
  expect(parseMonthly('unavailable')).toBeNull();
  expect(parseMonthly(undefined)).toBeNull();
});

test('rewrites the three alt strings and the date, and nothing else', () => {
  const next = rewriteSnapshots(README, { npm: 484591, crates: 203294, pypi: 2063 }, '2026-10-27');
  expect(next).toContain('alt="npm downloads: 484,591/month (snapshot 2026-10-27)"');
  expect(next).toContain('alt="crates.io downloads: 203,294/month (snapshot 2026-10-27)"');
  expect(next).toContain('alt="PyPI downloads: 2,063/month (snapshot 2026-10-27)"');
  expect(next).toContain('alt="license"');
});

test('a registry that could not be read keeps its old snapshot', async () => {
  const counts = await fetchCounts(answer({ npm: '10/month', pypi: '30/month' }) as never);
  expect(counts).toEqual({ npm: 10, pypi: 30 });
  const next = rewriteSnapshots(README, counts, '2026-10-27');
  expect(next).toContain('crates.io downloads: 2/month (snapshot 2026-01-01)');
  expect(next).toContain('npm downloads: 10/month (snapshot 2026-10-27)');
});

test('failing fetches leave the file untouched and do not throw', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'snap-'));
  const path = join(dir, 'README.md');
  writeFileSync(path, README);
  const broken = (async () => { throw new Error('offline'); }) as never;
  expect(await updateDownloadSnapshots(path, { fetchImpl: broken, today: '2026-10-27' })).toEqual({});
  expect(readFileSync(path, 'utf8')).toBe(README);
  const garbage = (async () => Response.json({ message: 'nope' })) as never;
  expect(await updateDownloadSnapshots(path, { fetchImpl: garbage, today: '2026-10-27' })).toEqual({});
  expect(readFileSync(path, 'utf8')).toBe(README);
});

test('a successful run writes the README', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'snap-'));
  const path = join(dir, 'README.md');
  writeFileSync(path, README);
  await updateDownloadSnapshots(path, { fetchImpl: answer({ npm: '5/month', crates: '6/month', pypi: '7/month' }) as never, today: '2026-10-27' });
  expect(readFileSync(path, 'utf8')).toContain('alt="PyPI downloads: 7/month (snapshot 2026-10-27)"');
});
