import { expect, test } from 'bun:test';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { buildSite, loadData } from './build';
import { renderPages, taskAnchor } from './render';
import type { Suite } from './stats';

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
    expect(pages.get('suites/v2/index.html')).toContain('30 synthetic fixtures (10 DOCX, 10 XLSX, 10 PPTX), one run per task, complete agent workflows with the same model: generic file tools vs the BetterOffice MCP server.');
    const reproduce = pages.get('suites/v2/reproduce/index.html')!;
    const data = await loadData();
    const suite = data.suites.find(s => s.suite === 'v2')!;
    for (const task of suite.tasks) {
      expect(reproduce).toContain(suite.fixtures.url + task.fixture);
      expect(reproduce).toContain(task.sha256);
      expect(reproduce).toContain(task.bytes + ' bytes');
    }
    expect(pages.get('suites/v2/method/index.html')).toContain('All 65 correct outputs accepted, all 30 untouched inputs rejected and all 157 corrupted outputs rejected.');
    expect(pages.get('suites/v2/failures/index.html')!.match(/<li>/g)).toHaveLength(suite.runs.filter(r => !r.pass).length);
    for (const r of suite.runs.filter(r => !r.pass)) expect(pages.get('suites/v2/failures/index.html')).toContain(r.reason!.replaceAll('"', '&quot;'));
    expect(pages.get('suites/v2/docx/index.html')).toContain('Generic file tools: 2026-10-15; BetterOffice MCP: 2026-10-15');
    expect(pages.get('suites/v2/method/index.html')).toContain('Generic file tools: 2026-10-13, 2026-10-15; BetterOffice MCP: 2026-10-14, 2026-10-15');
  } finally { await rm(directory, { recursive: true, force: true }); }
});

test('new suites own their page copy, fixtures, methods and repetitions', async () => {
  const data = await loadData(), previous = data.suites.find(s => s.suite === 'v2')!;
  const third = { id: 'third', label: 'Third arm', summary: 'Additional tools.', details: ['Third arm sandbox and access policy.'], tool_calls: 'Third arm calls.' };
  const tasks = previous.tasks.map(t => ({ ...t, fixture: 'real-' + t.fixture, bytes: t.bytes + 1, sha256: 'f'.repeat(64) }));
  const future: Suite = {
    ...previous, suite: 'v3', date: '2026-10-01',
    fixtures: { kind: 'real-world', count: tasks.length, url: 'https://example.org/v3/', metadata: 'https://example.org/v3/metadata.json', description: ['Real documents pinned by content hash.', 'Independent fixture provenance.'] },
    limits: ['These real documents cover a limited set of workflows.'], reproduce: ['run-real-suite --three-runs'],
    harness: 'future harness', checker_validation: undefined, tasks, arms: [...previous.arms, third],
    runs: tasks.flatMap(t => [...previous.arms, third].flatMap(a => [1, 2, 3].map(run => ({
      ...previous.runs[0]!, model: 'model-a', task: t.id, arm: a.id, run, pass: a.id === 'betteroffice',
      reason: a.id === 'betteroffice' ? null : a.id + ' failed real document edit.', cause: a.id === 'betteroffice' ? null : 'model' as const,
    })))),
  };
  future.runs.push(...future.runs.filter(r => r.task === tasks[0]!.id).map(r => ({ ...r, model: 'model-b' })));
  const pages = renderPages({ ...data, suites: [future] });
  for (const path of ['index.html', 'method/index.html', 'reproduce/index.html', 'failures/index.html', 'docx/index.html', 'xlsx/index.html', 'pptx/index.html']) {
    const page = pages.get(path)!;
    expect(page).toContain('v3');
    expect(page).toContain('real-world');
    expect(page).not.toMatch(/v2|synthetic|one run per task|both arms|different days/);
    expect(page.replace(/<section[^>]*data-arm="generic"[\s\S]*?<\/section>/g, '')).not.toContain('python-docx');
  }
  expect(pages.get('index.html')).toContain('3 runs per task with model-a; 3 runs per task with model-b on 1 task');
  expect(pages.get('index.html')).toContain('All runs pass:');
  const method = pages.get('method/index.html')!;
  expect(method).toContain('future harness');
  expect(method).not.toContain('Checker validation:');
  for (const text of [...future.fixtures.description, ...future.limits, ...future.arms.flatMap(a => [a.summary, ...a.details, a.tool_calls])]) expect(method).toContain(text);
  const reproduce = pages.get('reproduce/index.html')!;
  expect(reproduce).toContain('Every v3 fixture');
  for (const text of [...future.reproduce, ...future.fixtures.description]) expect(reproduce).toContain(text);
  for (const task of tasks) {
    expect(reproduce).toContain(future.fixtures.url + task.fixture);
    expect(reproduce).toContain(task.sha256);
    expect(reproduce).toContain(task.bytes + ' bytes');
  }
  const failures = pages.get('failures/index.html')!;
  expect(failures).toContain('Where runs failed');
  expect(failures).toContain('Third arm');
  for (const run of future.runs.filter(r => !r.pass)) {
    const task = tasks.find(t => t.id === run.task)!;
    const anchor = taskAnchor(run.model, run.task);
    expect(failures).toContain('href="/' + task.format + '#' + anchor + '"');
    expect(pages.get(task.format + '/index.html')!.split('id="' + anchor + '"')).toHaveLength(2);
  }
  for (const format of ['docx', 'xlsx', 'pptx']) {
    const page = pages.get(format + '/index.html')!;
    expect(page).toContain('3/3 pass');
    expect(page).toContain('Third arm');
    expect(page).toContain('No run recorded.');
    const ids = [...page.matchAll(/\bid="([^"]+)"/g)].map(m => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
  }
  const combined = renderPages({ ...data, suites: [previous, future] });
  expect(combined.get('index.html')).toContain('v3');
  expect(combined.get('suites/v2/index.html')).toContain('30 synthetic fixtures');
  const rows = combined.get('index.html')!.match(/<tr><td data-label="Version">[\s\S]*?<\/tr>/g)!;
  for (const s of [previous, future]) {
    const row = rows.find(row => row.includes('href="/suites/' + s.suite + '"'))!;
    expect(row).toContain('href="/results/' + s.suite + '.json"');
  }
});

test('missing suite descriptions fail before rendering', async () => {
  const data = await loadData();
  for (const mutate of [
    (s: Suite) => { delete (s as Partial<Suite>).limits; },
    (s: Suite) => { delete (s as Partial<Suite>).reproduce; },
    (s: Suite) => { s.fixtures.description = []; },
    (s: Suite) => { s.arms[0]!.details = []; },
    (s: Suite) => { s.tasks[0]!.fixture = ''; },
    (s: Suite) => { s.tasks[0]!.bytes = 0; },
    (s: Suite) => { s.tasks[0]!.sha256 = ''; },
  ]) {
    const suite = structuredClone(data.suites[0]!);
    mutate(suite);
    expect(() => renderPages({ ...data, suites: [suite] })).toThrow();
  }
});

test('task anchors are unambiguous and safe for ids and URL fragments', () => {
  expect(taskAnchor('a-b', 'c')).not.toBe(taskAnchor('a', 'b-c'));
  expect(taskAnchor('model /é', 'task #"')).toMatch(/^task-[a-f0-9-]+$/);
});
