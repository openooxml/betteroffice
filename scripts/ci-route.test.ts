import { describe, expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { LABEL, SAMPLE_RUNS, allocate, busyHosted, hostedJobs, parseGroups, route } from './ci-route.mjs';

const script = fileURLToPath(new URL('./ci-route.mjs', import.meta.url));
const REPOSITORY = 'openooxml/betteroffice';

const job = (status: string, label: string) => ({ status, labels: [label] });

// A fake Actions API: runs maps a run id to its jobs; ids listed in queued come back as queued runs.
function actions(runs: Record<number, object[]>, { queued = [] as number[], failWith = 0 } = {}) {
  const calls: string[] = [];
  const fetchImpl = async (url: string) => {
    calls.push(url);
    if (failWith) return new Response('', { status: failWith });
    const path = new URL(url).pathname + new URL(url).search;
    const listed = path.match(/\/actions\/runs\?status=(\w+)/);
    if (listed) {
      const ids = Object.keys(runs).map(Number);
      const page = listed[1] === 'queued' ? queued : ids.filter((id) => !queued.includes(id));
      return Response.json({ workflow_runs: page.map((id) => ({ id })) });
    }
    const run = Number(path.match(/\/runs\/(\d+)\/jobs/)?.[1]);
    return Response.json({ jobs: runs[run] ?? [] });
  };
  return { calls, fetchImpl: fetchImpl as unknown as typeof fetch };
}

const env = (extra: Record<string, string> = {}) => ({
  ROUTE_GROUPS: 'packages python package_tests:2 rust',
  GITHUB_REPOSITORY: REPOSITORY,
  GH_TOKEN: 'token',
  ...extra,
});

describe('runner routing', () => {
  test('groups parse in order with their slot counts and reject malformed entries', () => {
    expect(parseGroups(' packages  package_tests:2\nrust ')).toEqual([
      { name: 'packages', slots: 1 },
      { name: 'package_tests', slots: 2 },
      { name: 'rust', slots: 1 },
    ]);
    for (const bad of ['rust:0', 'rust:x', 'Rust', 'package-tests', 'a:1:2']) {
      expect(() => parseGroups(bad), bad).toThrow('invalid route group');
    }
  });

  test('allocation fills free GitHub-hosted slots first fit, in order', () => {
    const groups = parseGroups('packages python package_tests:2 rust');
    expect(allocate(groups, 5)).toEqual({ packages: LABEL, python: LABEL, package_tests: LABEL, rust: LABEL });
    expect(allocate(groups, 3)).toEqual({ packages: LABEL, python: LABEL, package_tests: '', rust: LABEL });
    expect(allocate(groups, 0)).toEqual({ packages: '', python: '', package_tests: '', rust: '' });
  });

  test('only queued and running jobs on a single GitHub-hosted label count', () => {
    expect(
      hostedJobs([
        job('queued', 'ubuntu-latest'),
        job('in_progress', 'ubuntu-24.04'),
        job('in_progress', 'macos-15'),
        job('completed', 'ubuntu-latest'),
        job('in_progress', 'kyora-8x8'),
        { status: 'queued', labels: ['self-hosted', 'linux'] },
        { status: 'queued' },
      ])
    ).toBe(3);
  });

  test('the count takes every GitHub-hosted job, routed or not, and stops once the cap is reached', async () => {
    const runs = {
      1: [job('in_progress', 'ubuntu-latest'), job('queued', 'ubuntu-latest')],
      2: [job('in_progress', 'ubuntu-latest'), job('in_progress', 'kyora-8x16')],
      3: [job('in_progress', 'ubuntu-latest')],
    };
    const { fetchImpl } = actions(runs, { queued: [3] });
    const api = async (path: string) => (await fetchImpl(`https://api.github.com${path}`)).json();
    expect(await busyHosted(api, REPOSITORY, 18)).toBe(4);
    const stopped = actions(runs, { queued: [3] });
    const capped = async (path: string) => (await stopped.fetchImpl(`https://api.github.com${path}`)).json();
    expect(await busyHosted(capped, REPOSITORY, 2)).toBe(2);
    expect(stopped.calls.filter((url) => url.includes('/jobs'))).toHaveLength(1);
  });

  test('past the sample, the count scales the sampled runs up instead of listing every run', async () => {
    const runs: Record<number, object[]> = Object.fromEntries(Array.from({ length: SAMPLE_RUNS * 2 }, (_, i) => [i + 1, []]));
    runs[1] = [job('in_progress', 'ubuntu-latest')];
    runs[2] = [job('in_progress', 'ubuntu-latest'), job('in_progress', 'kyora-8x8')];
    runs[SAMPLE_RUNS + 1] = Array(5).fill(job('queued', 'ubuntu-latest'));
    const { calls, fetchImpl } = actions(runs);
    expect(await route(env(), fetchImpl)).toEqual({
      busy: 4,
      routes: { packages: LABEL, python: LABEL, package_tests: LABEL, rust: LABEL },
    });
    expect(calls.filter((url) => url.includes('/jobs'))).toHaveLength(SAMPLE_RUNS);
  });

  test('free slots are the limit less the reserve less the busy jobs', async () => {
    const runs = { 1: Array(15).fill(job('in_progress', 'ubuntu-latest')) };
    const { fetchImpl } = actions(runs);
    expect(await route(env(), fetchImpl)).toEqual({
      busy: 15,
      routes: { packages: LABEL, python: LABEL, package_tests: '', rust: LABEL },
    });
    expect((await route(env({ BO_CI_GH_LIMIT: '22', BO_CI_GH_RESERVE: '0' }), fetchImpl)).routes).toEqual({
      packages: LABEL,
      python: LABEL,
      package_tests: LABEL,
      rust: LABEL,
    });
    expect((await route(env({ BO_CI_GH_LIMIT: 'lots' }), fetchImpl)).busy).toBe(15);
  });

  test('the kill switch and a runner override skip the API and route nothing', async () => {
    for (const extra of [{ BO_CI_SPILL: 'off' }, { BO_CI_RUNNER: 'kyora-8x8' }]) {
      const { calls, fetchImpl } = actions({ 1: [job('in_progress', 'ubuntu-latest')] });
      expect((await route(env(extra), fetchImpl)).routes).toEqual(allocate(parseGroups(env().ROUTE_GROUPS), 0));
      expect(calls).toEqual([]);
    }
  });

  test('an API error rejects so the caller falls back to kyora', async () => {
    const { fetchImpl } = actions({}, { failWith: 403 });
    await expect(route(env(), fetchImpl)).rejects.toThrow('GitHub answered 403');
  });

  test('the deadline aborts requests that never answer and the command still exits cleanly with empty outputs', async () => {
    const server = Bun.serve({ port: 0, fetch: () => new Promise<Response>(() => {}) });
    const directory = mkdtempSync(join(tmpdir(), 'ci-route-'));
    try {
      const output = join(directory, 'output');
      const started = Date.now();
      const command = Bun.spawn(['node', script], {
        env: {
          ...process.env,
          ...env({ GITHUB_API_URL: `http://127.0.0.1:${server.port}`, BO_CI_ROUTE_DEADLINE_MS: '300' }),
          GITHUB_OUTPUT: output,
        },
        stdout: 'pipe',
      });
      expect(await command.exited).toBe(0);
      expect(Date.now() - started).toBeLessThan(5_000);
      expect(await new Response(command.stdout).text()).toContain('fell back to kyora');
      expect(readFileSync(output, 'utf8')).toBe('packages=\npython=\npackage_tests=\nrust=\n');
    } finally {
      server.stop(true);
      rmSync(directory, { recursive: true, force: true });
    }
  });

  test('the command writes every group, empty when routing fails, and exits cleanly', () => {
    const directory = mkdtempSync(join(tmpdir(), 'ci-route-'));
    try {
      const output = join(directory, 'output');
      const result = spawnSync('node', [script], {
        encoding: 'utf8',
        env: { ...process.env, ...env({ GITHUB_API_URL: 'http://127.0.0.1:9' }), GITHUB_OUTPUT: output },
      });
      expect(result.status).toBe(0);
      expect(result.stdout).toContain('fell back to kyora');
      expect(readFileSync(output, 'utf8')).toBe('packages=\npython=\npackage_tests=\nrust=\n');
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });
});
