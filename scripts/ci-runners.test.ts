import { describe, expect, test } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';

const REPOSITORY = 'openooxml/betteroffice';
const KYORA_CLASSES = ['kyora-4x8', 'kyora-8x8', 'kyora-8x16', 'kyora-16x32'];
const GITHUB_HOSTED = /^(ubuntu|macos|windows)-[\w.-]+$/;
const REUSABLE = '(reusable workflow)';

const pullRequest = (headRepository: string | null) => ({
  event_name: 'pull_request',
  repository: REPOSITORY,
  event: { pull_request: { head: { repo: headRepository && { full_name: headRepository } } } },
});

const EVENTS: Record<string, object> = {
  'push to main': { event_name: 'push', repository: REPOSITORY, event: {} },
  dispatch: { event_name: 'workflow_dispatch', repository: REPOSITORY, event: { inputs: {} } },
  'same-repository pull request': pullRequest(REPOSITORY),
  'fork pull request': pullRequest('contributor/betteroffice'),
  'pull request from a deleted fork': pullRequest(null),
};
const TRUSTED = ['push to main', 'dispatch', 'same-repository pull request'];
const UNTRUSTED = ['fork pull request', 'pull request from a deleted fork'];

const VARIABLES: Record<string, Record<string, string>> = {
  unset: {},
  'BO_CI_RUNNER=ubuntu-latest': { BO_CI_RUNNER: 'ubuntu-latest' },
  'BO_CI_RUNNER=kyora-8x8': { BO_CI_RUNNER: 'kyora-8x8' },
};

// GitHub's expression language as far as runs-on uses it: context paths, string and keyword
// literals, !, ==, !=, && and ||, and parentheses. Anything else throws.
function evaluate(source: string, context: Record<string, unknown>): unknown {
  const tokens = source.match(/'(?:[^']|'')*'|[A-Za-z_][\w.-]*|==|!=|&&|\|\||[!()]|\S/g) ?? [];
  let at = 0;
  const unsupported = () => new Error(`unsupported expression: ${source}`);
  const take = () => {
    if (at >= tokens.length) throw unsupported();
    return tokens[at++];
  };
  const primary = (): unknown => {
    const token = take();
    if (token === '(') {
      const value = or();
      if (take() !== ')') throw unsupported();
      return value;
    }
    if (token.startsWith("'")) return token.slice(1, -1).replaceAll("''", "'");
    if (token === 'true' || token === 'false') return token === 'true';
    if (token === 'null') return null;
    if (!/^[A-Za-z_]/.test(token) || tokens[at] === '(') throw unsupported();
    return token
      .split('.')
      .reduce<unknown>((value, key) => (value == null ? null : ((value as any)[key] ?? null)), context);
  };
  const unary = (): unknown => {
    if (tokens[at] !== '!') return primary();
    take();
    return !unary();
  };
  const comparison = (): unknown => {
    const left = unary();
    const operator = tokens[at];
    if (operator !== '==' && operator !== '!=') return left;
    take();
    const equal = equals(left, unary());
    return operator === '==' ? equal : !equal;
  };
  const and = (): unknown => {
    let value = comparison();
    while (tokens[at] === '&&') {
      take();
      const right = comparison();
      value = value ? right : value;
    }
    return value;
  };
  const or = (): unknown => {
    let value = and();
    while (tokens[at] === '||') {
      take();
      const right = and();
      value = value ? value : right;
    }
    return value;
  };
  const value = or();
  if (at !== tokens.length) throw unsupported();
  return value;
}

function equals(left: unknown, right: unknown): boolean {
  if (typeof left === 'string' && typeof right === 'string') {
    return left.toLowerCase() === right.toLowerCase();
  }
  if (typeof left === typeof right) return left === right;
  const number = (value: unknown) =>
    value === null || (typeof value === 'string' && value.trim() === '')
      ? 0
      : typeof value === 'object'
        ? NaN
        : Number(value);
  return number(left) === number(right);
}

function resolve(runsOn: unknown, context: Record<string, unknown>): string[] {
  if (Array.isArray(runsOn)) return runsOn.flatMap((label) => resolve(label, context));
  if (typeof runsOn !== 'string') throw new Error(`unsupported runs-on: ${JSON.stringify(runsOn)}`);
  const expression = runsOn.match(/^\$\{\{(.*)\}\}$/s);
  if (expression) return [String(evaluate(expression[1], context))];
  if (runsOn.includes('${{')) throw new Error(`unsupported runs-on: ${runsOn}`);
  return [runsOn];
}

// GitHub's matrix expansion: axes multiply, exclude removes, include extends the original rows it
// does not contradict or adds a row of its own.
function combinations(matrix: any): Record<string, unknown>[] {
  if (matrix === undefined) return [{}];
  if (typeof matrix !== 'object' || JSON.stringify(matrix).includes('${{')) {
    throw new Error(`runs-on reads a matrix that is an expression: ${JSON.stringify(matrix)}`);
  }
  const { include = [], exclude = [], ...axes } = matrix;
  const matches = (row: Record<string, unknown>, entry: Record<string, unknown>) =>
    Object.entries(entry).every(([key, value]) => JSON.stringify(row[key]) === JSON.stringify(value));
  let rows: Record<string, unknown>[] = Object.keys(axes).length === 0 ? [] : [{}];
  for (const [key, values] of Object.entries<unknown[]>(axes)) {
    rows = rows.flatMap((row) => values.map((value) => ({ ...row, [key]: value })));
  }
  rows = rows.filter((row) => !exclude.some((entry: any) => matches(row, entry)));
  const added: Record<string, unknown>[] = [];
  for (const entry of include) {
    const original = Object.fromEntries(Object.entries(entry).filter(([key]) => key in axes));
    const extended = rows.filter((row) => matches(row, original));
    if (extended.length === 0) added.push({ ...entry });
    for (const row of extended) Object.assign(row, entry);
  }
  return [...rows, ...added];
}

type Api = {
  runs: { queued: { id: number }[]; in_progress: { id: number }[] };
  jobs: Record<number, { status: string; labels: string[] }[]>;
  fail?: boolean;
  truncated?: 'runs' | 'jobs';
};

const IDLE: Api = { runs: { queued: [], in_progress: [] }, jobs: {} };
const directory = new URL('../.github/workflows/', import.meta.url);
const ciWorkflow = Bun.YAML.parse(readFileSync(new URL('ci.yml', directory), 'utf8')) as any;

async function route(github: object, vars: Record<string, string>, api: Api) {
  const step = ciWorkflow.jobs.route.steps.find((entry: any) => entry.id === 'pick');
  const env: Record<string, string> = {};
  for (const [key, raw] of Object.entries<unknown>(step.env ?? {})) {
    const expression = typeof raw === 'string' ? raw.match(/^\$\{\{(.*)\}\}$/s) : null;
    const value = expression ? evaluate(expression[1], { github, vars }) : raw;
    env[key] = value === null ? '' : typeof value === 'boolean' ? (value ? 'true' : 'false') : String(value);
  }
  const outputs: Record<string, string> = {};
  const calls: { endpoint: 'runs' | 'jobs'; params: Record<string, any> }[] = [];
  const fn = new (Object.getPrototypeOf(async function () {}).constructor)(
    'github',
    'context',
    'core',
    'process',
    step.with.script
  );
  await fn(
    {
      rest: {
        actions: {
          listWorkflowRunsForRepo: async (params: Record<string, any>) => {
            calls.push({ endpoint: 'runs', params });
            if (api.fail) throw new Error('API failure');
            const workflow_runs = api.runs[params.status as 'queued' | 'in_progress'];
            return {
              data: {
                total_count: workflow_runs.length + (api.truncated === 'runs' ? 1 : 0),
                workflow_runs,
              },
            };
          },
          listJobsForWorkflowRun: async (params: Record<string, any>) => {
            calls.push({ endpoint: 'jobs', params });
            if (api.fail) throw new Error('API failure');
            const jobs = api.jobs[params.run_id] ?? [];
            return {
              data: {
                total_count: jobs.length + (api.truncated === 'jobs' ? 1 : 0),
                jobs,
              },
            };
          },
        },
      },
    },
    { repo: { owner: 'openooxml', repo: 'betteroffice' } },
    {
      setOutput: (name: string, value: string) => {
        outputs[name] = value;
      },
      info: () => {},
    },
    { env }
  );
  return { runner: outputs.runner, calls };
}

const jobs = await Promise.all(
  readdirSync(directory)
    .filter((file) => /\.ya?ml$/.test(file))
    .flatMap((file) => {
      const workflow = Bun.YAML.parse(readFileSync(new URL(file, directory), 'utf8')) as any;
      const triggers: string[] =
        typeof workflow.on === 'string' ? [workflow.on] : Array.isArray(workflow.on) ? workflow.on : Object.keys(workflow.on);
      return Object.entries<any>(workflow.jobs).flatMap(([id, job]) => {
        const readsMatrix = JSON.stringify(job['runs-on'] ?? '').includes('matrix.');
        const needsRoute = [job.needs ?? []].flat().includes('route');
        return (readsMatrix ? combinations(job.strategy?.matrix) : [{}]).map(async (matrix) => {
          const runners: Record<string, Record<string, string[]>> = {};
          for (const [event, github] of Object.entries(EVENTS)) {
            runners[event] = {};
            for (const [variables, vars] of Object.entries(VARIABLES)) {
              const needs = needsRoute ? { route: { result: 'success', outputs: { runner: (await route(github, vars, IDLE)).runner } } } : {};
              runners[event][variables] = job.uses ? [REUSABLE] : resolve(job['runs-on'], { github, vars, matrix, needs });
            }
          }
          const suffix =
            typeof matrix.name === 'string' ? ` (${matrix.name})` : Object.keys(matrix).length > 0 ? ` ${JSON.stringify(matrix)}` : '';
          return { name: `${file} ${id}${suffix}`, workflow, triggers, job, runners };
        });
      });
    })
);

const githubHosted = (labels: string[]) => labels.length === 1 && GITHUB_HOSTED.test(labels[0]);
const resolved = (runners: Record<string, Record<string, string[]>>) =>
  Object.values(runners).flatMap((byVariables) => Object.values(byVariables));
const kyoraJobs = jobs.filter(({ runners }) =>
  resolved(runners).some((labels) => labels[0] !== REUSABLE && !githubHosted(labels))
);

describe('CI runner routing', () => {
  test('only the Rust gate and the corpus scenarios run on kyora, each on one class', () => {
    expect(Object.fromEntries(kyoraJobs.map(({ name, runners }) => [name, runners['push to main'].unset]))).toEqual({
      'ci.yml rust': ['kyora-8x16'],
      'e2e.yml scenarios (Corpus scenarios)': ['kyora-8x8'],
    });
    for (const { name, runners } of kyoraJobs) {
      for (const event of TRUSTED) {
        expect(runners[event].unset, `${name}: ${event}`).toEqual(runners['push to main'].unset);
        expect(runners[event]['BO_CI_RUNNER=kyora-8x8'], `${name}: ${event}`).toEqual(['kyora-8x8']);
      }
      for (const labels of resolved(runners)) {
        expect(githubHosted(labels) || (labels.length === 1 && KYORA_CLASSES.includes(labels[0])), name).toBe(true);
      }
    }
  });

  test('fork pull requests and BO_CI_RUNNER=ubuntu-latest run every job on GitHub-hosted runners', () => {
    for (const { name, runners } of jobs) {
      if (runners['push to main'].unset[0] === REUSABLE) continue;
      for (const event of UNTRUSTED) {
        for (const variables of Object.keys(VARIABLES)) {
          expect(githubHosted(runners[event][variables]), `${name}: ${event}, ${variables}`).toBe(true);
        }
      }
      for (const event of Object.keys(EVENTS)) {
        expect(githubHosted(runners[event]['BO_CI_RUNNER=ubuntu-latest']), `${name}: ${event}`).toBe(true);
      }
    }
  });

  test('kyora jobs read no secrets, write nothing, use no Docker and end before the class timeout', () => {
    for (const { name, workflow, triggers, job } of kyoraJobs) {
      expect(triggers, name).not.toContain('pull_request_target');
      expect(triggers, name).not.toContain('workflow_run');
      expect(JSON.stringify([workflow.env ?? null, workflow.defaults ?? null, job]), name).not.toMatch(/\bsecrets\b/);
      expect(workflow.permissions, name).toEqual({ contents: 'read' });
      expect(job.permissions, name).toBeUndefined();
      expect(job.services, name).toBeUndefined();
      expect(job.container, name).toBeUndefined();
      for (const step of job.steps) {
        expect(String(step.uses ?? ''), name).not.toMatch(/^docker:\/\/|^EmbarkStudios\/cargo-deny-action@/);
      }
      expect(job['timeout-minutes'], name).toBeLessThan(120);
    }
  });

  test('the route job load-balances the Rust gate', async () => {
    const threeBusy: Api = {
      runs: { queued: [{ id: 1 }], in_progress: [{ id: 2 }] },
      jobs: {
        1: [{ status: 'queued', labels: ['kyora-8x16'] }],
        2: [
          { status: 'in_progress', labels: ['kyora-8x16'] },
          { status: 'queued', labels: ['kyora-8x8'] },
        ],
      },
    };
    const fourBusy: Api = {
      runs: { queued: [{ id: 1 }], in_progress: [{ id: 2 }] },
      jobs: {
        ...threeBusy.jobs,
        1: [...threeBusy.jobs[1], { status: 'in_progress', labels: ['kyora-8x16'] }],
      },
    };
    const cases: {
      name: string;
      event?: string;
      vars?: Record<string, string>;
      api?: Api;
      expected: string;
      noApiCalls?: boolean;
    }[] = [
      { name: 'fork pull request, idle', event: 'fork pull request', expected: 'ubuntu-latest', noApiCalls: true },
      {
        name: 'fork pull request with BO_CI_RUNNER=kyora-8x8',
        event: 'fork pull request',
        vars: { BO_CI_RUNNER: 'kyora-8x8' },
        expected: 'ubuntu-latest',
        noApiCalls: true,
      },
      {
        name: 'pull request from a deleted fork',
        event: 'pull request from a deleted fork',
        expected: 'ubuntu-latest',
        noApiCalls: true,
      },
      { name: 'same-repository pull request, idle', event: 'same-repository pull request', expected: 'kyora-8x16' },
      { name: 'push to main, idle', event: 'push to main', expected: 'kyora-8x16' },
      { name: 'dispatch, idle', event: 'dispatch', expected: 'kyora-8x16' },
      {
        name: 'BO_CI_RUNNER=ubuntu-latest',
        vars: { BO_CI_RUNNER: 'ubuntu-latest' },
        expected: 'ubuntu-latest',
        noApiCalls: true,
      },
      {
        name: 'BO_CI_RUNNER=kyora-8x8',
        vars: { BO_CI_RUNNER: 'kyora-8x8' },
        expected: 'kyora-8x8',
        noApiCalls: true,
      },
      { name: 'BO_KYORA_CLASS=kyora-8x8, idle', vars: { BO_KYORA_CLASS: 'kyora-8x8' }, expected: 'kyora-8x8' },
      { name: '3 busy kyora jobs under BO_KYORA_CAP=4', vars: { BO_KYORA_CAP: '4' }, api: threeBusy, expected: 'kyora-8x16' },
      { name: '4 busy kyora jobs at BO_KYORA_CAP=4', vars: { BO_KYORA_CAP: '4' }, api: fourBusy, expected: 'ubuntu-latest' },
      { name: '4 busy kyora jobs under the default cap 20', api: fourBusy, expected: 'kyora-8x16' },
      {
        name: '4 busy kyora jobs with BO_KYORA_CAP=6',
        vars: { BO_KYORA_CAP: '6' },
        api: fourBusy,
        expected: 'kyora-8x16',
      },
      { name: 'BO_KYORA_CAP=0 when idle', vars: { BO_KYORA_CAP: '0' }, expected: 'ubuntu-latest' },
      {
        name: 'completed, GitHub-hosted, self-hosted and waiting jobs are not counted',
        vars: { BO_KYORA_CAP: '4' },
        api: {
          runs: { queued: [{ id: 1 }], in_progress: [{ id: 2 }] },
          jobs: {
            1: [
              ...threeBusy.jobs[1],
              { status: 'completed', labels: ['kyora-8x16'] },
              { status: 'queued', labels: ['ubuntu-latest'] },
            ],
            2: [
              ...threeBusy.jobs[2],
              { status: 'in_progress', labels: ['self-hosted'] },
              { status: 'waiting', labels: ['kyora-8x16'] },
            ],
          },
        },
        expected: 'kyora-8x16',
      },
      {
        name: 'same run id in queued and in_progress is counted once',
        vars: { BO_KYORA_CAP: '3' },
        api: {
          runs: { queued: [{ id: 7 }], in_progress: [{ id: 7 }] },
          jobs: {
            7: [
              { status: 'queued', labels: ['kyora-8x16'] },
              { status: 'in_progress', labels: ['kyora-8x8'] },
            ],
          },
        },
        expected: 'kyora-8x16',
      },
      { name: 'API failure falls back to ubuntu-latest', api: { ...IDLE, fail: true }, expected: 'ubuntu-latest' },
      {
        name: 'truncated runs listing falls back to ubuntu-latest',
        api: { ...IDLE, truncated: 'runs' },
        expected: 'ubuntu-latest',
      },
      {
        name: 'truncated jobs listing falls back to ubuntu-latest',
        api: { runs: { queued: [{ id: 1 }], in_progress: [] }, jobs: { 1: [] }, truncated: 'jobs' },
        expected: 'ubuntu-latest',
      },
    ];

    for (const { name, event = 'push to main', vars = {}, api = IDLE, expected, noApiCalls } of cases) {
      const { runner, calls } = await route(EVENTS[event], vars, api);
      expect(runner, name).toBe(expected);
      if (noApiCalls) {
        expect(calls, name).toHaveLength(0);
      } else {
        expect(calls.length, name).toBeGreaterThan(0);
        for (const call of calls) {
          expect(call.params.per_page, name).toBe(100);
          expect(call.params.request?.signal, name).toBeInstanceOf(AbortSignal);
          if (call.endpoint === 'jobs') {
            expect(call.params.filter, name).toBe('latest');
          }
        }
      }
    }
    const full = { status: 'queued', labels: ['kyora-8x16'] };
    const early = await route(EVENTS['push to main'], { BO_KYORA_CAP: '1' }, { runs: { queued: [{ id: 1 }, { id: 2 }], in_progress: [] }, jobs: { 1: [full], 2: [full] } });
    expect(early.runner).toBe('ubuntu-latest');
    expect(early.calls.filter((call) => call.endpoint === 'jobs')).toHaveLength(1);
  });

  test('the route job reads only Actions metadata and the Rust gate survives its failure', () => {
    const triggers = Object.keys(ciWorkflow.on);
    expect(triggers).not.toContain('pull_request_target');
    expect(triggers).not.toContain('workflow_run');

    const routeJob = ciWorkflow.jobs.route;
    expect(routeJob['runs-on']).toBe('ubuntu-latest');
    expect(routeJob.permissions).toEqual({ actions: 'read' });
    expect(routeJob['timeout-minutes']).toBe(2);
    expect(routeJob.outputs).toEqual({ runner: '${{ steps.pick.outputs.runner }}' });
    expect(routeJob.steps).toHaveLength(1);
    expect(routeJob.steps[0].uses).toBe('actions/github-script@v9');
    for (const step of routeJob.steps) {
      expect(String(step.uses ?? '')).not.toMatch(/^actions\/checkout/);
    }
    expect(JSON.stringify(routeJob)).not.toMatch(/\bsecrets\b/);
    expect(routeJob.steps[0].with.script).not.toContain('${{');

    const rustJob = ciWorkflow.jobs.rust;
    expect([rustJob.needs].flat()).toEqual(['route']);
    expect(rustJob.if).toBe('${{ !cancelled() }}');
    const routeEnds = [
      { result: 'success', outputs: { runner: '' } },
      { result: 'failure', outputs: { runner: 'kyora-8x16' } },
      { result: 'cancelled', outputs: { runner: 'kyora-8x16' } },
      { result: 'skipped', outputs: {} },
    ];
    for (const [event, github] of Object.entries(EVENTS)) {
      for (const vars of Object.values(VARIABLES)) {
        for (const route of routeEnds) {
          expect(resolve(rustJob['runs-on'], { github, vars, needs: { route } }), `${event}, ${route.result}`).toEqual(['ubuntu-latest']);
        }
      }
    }
  });

  test('the expression subset follows GitHub where routing depends on it', () => {
    const context = { github: pullRequest(null), vars: {}, matrix: { suite: 'corpus' } };
    expect(evaluate('github.event.pull_request.head.repo.full_name != github.repository', context)).toBe(true);
    expect(evaluate("vars.BO_CI_RUNNER || 'kyora-4x8'", context)).toBe('kyora-4x8');
    expect(evaluate("!(matrix.suite == 'CORPUS') && 'a' || 'b'", context)).toBe('b');
    expect(evaluate("null == ''", context)).toBe(true);
    expect(evaluate('matrix.guard == false', { matrix: { guard: [] } })).toBe(false);
    expect(() => evaluate("fromJSON('[]')", context)).toThrow('unsupported expression');
    expect(combinations({ os: ['a', 'b'], include: [{ os: 'a', x: 1 }, { os: 'c' }, { y: 2 }] })).toEqual([
      { os: 'a', x: 1, y: 2 },
      { os: 'b', y: 2 },
      { os: 'c' },
    ]);
  });
});
