import { expect, test } from 'bun:test';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { buildSite, loadData } from './build';
import { renderPages } from './render';

test('build every page without client scripts, invalid values or an em dash', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'evals-build-'));
  try {
    const pages = await buildSite(directory);
    for (const path of ['index.html', 'docx/index.html', 'xlsx/index.html', 'pptx/index.html', 'failures/index.html', 'method/index.html', 'reproduce/index.html', 'suites/v2/index.html']) expect(pages.has(path)).toBe(true);
    for (const [path, page] of pages) {
      expect(await readFile(join(directory, path), 'utf8')).toBe(page);
      expect(page).not.toMatch(/[\u2014]|undefined|NaN|<script/);
      expect(page).toContain('name="viewport"');
    }
    expect(pages.get('index.html')).toContain('30 synthetic fixtures (10 DOCX, 10 XLSX, 10 PPTX), one run per task');
    const reproduce = pages.get('reproduce/index.html')!;
    const data = await loadData();
    for (const name of Object.keys(data.fixtures.fixtures)) expect(reproduce).toContain(data.fixtures.origin + '/' + name);
    expect(pages.get('method/index.html')).toContain('65/65');
    expect(pages.get('failures/index.html')!.match(/<li>/g)).toHaveLength(data.suites[0]!.runs.filter(r => !r.pass).length);
  } finally { await rm(directory, { recursive: true, force: true }); }
});

test('new suites preserve old pages and render repeated runs, a third arm, and a second model subset', async () => {
  const data = await loadData(), previous = data.suites[0]!;
  const third = { id: 'third', label: 'Third arm', summary: 'Additional tools.', tool_calls: 'Calls.' };
  const future = { ...previous, suite: 'v3', date: '2026-10-16', fixtures: { ...previous.fixtures, kind: 'real-world' }, arms: [...previous.arms, third], runs: previous.runs.flatMap(r => [1, 2, 3].map(run => ({ ...r, run }))) };
  future.runs.push({ ...future.runs[0]!, arm: 'third', model: 'second-model', pass: true, reason: null, cause: null });
  const pages = renderPages({ ...data, suites: [previous, future] });
  expect(pages.get('index.html')).toContain('v3');
  expect(pages.get('suites/v2/index.html')).toContain('v2');
  expect(pages.get('index.html')).toContain('All runs pass:');
  expect(pages.get('docx/index.html')).toContain('3/3 pass');
  expect(pages.get('docx/index.html')).toContain('Third arm');
  expect(pages.get('docx/index.html')).toContain('No run recorded.');
});
