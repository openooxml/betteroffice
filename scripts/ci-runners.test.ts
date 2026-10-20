import { describe, expect, test } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';
import { LABEL, parseGroups } from './ci-route.mjs';

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
  'BO_CI_SPILL=off': { BO_CI_SPILL: 'off' },
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

const directory = new URL('../.github/workflows/', import.meta.url);
const jobs = readdirSync(directory)
  .filter((file) => /\.ya?ml$/.test(file))
  .flatMap((file) => {
    const workflow = Bun.YAML.parse(readFileSync(new URL(file, directory), 'utf8')) as any;
    const triggers: string[] =
      typeof workflow.on === 'string' ? [workflow.on] : Array.isArray(workflow.on) ? workflow.on : Object.keys(workflow.on);
    return Object.entries<any>(workflow.jobs).flatMap(([id, job]) => {
      const readsMatrix = JSON.stringify(job['runs-on'] ?? '').includes('matrix.');
      const legs = () => (readsMatrix || !job.strategy?.matrix ? 1 : combinations(job.strategy.matrix).length);
      return (readsMatrix ? combinations(job.strategy?.matrix) : [{}]).map((matrix) => {
        // A router output names the GitHub-hosted label for its group; an empty one leaves the job on kyora.
        const runsOn = (github: object, vars: object, outputs: Record<string, string> = {}) =>
          job.uses ? [REUSABLE] : resolve(job['runs-on'], { github, vars, matrix, needs: { route: { outputs } } });
        const runners: Record<string, Record<string, string[]>> = {};
        for (const [event, github] of Object.entries(EVENTS)) {
          runners[event] = {};
          for (const [variables, vars] of Object.entries(VARIABLES)) runners[event][variables] = runsOn(github, vars);
        }
        const suffix =
          typeof matrix.name === 'string' ? ` (${matrix.name})` : Object.keys(matrix).length > 0 ? ` ${JSON.stringify(matrix)}` : '';
        return { name: `${file} ${id}${suffix}`, file, id, workflow, triggers, job, legs, runsOn, runners };
      });
    });
  });

function timeouts(job: any): unknown[] {
  const timeout = job['timeout-minutes'];
  const key = typeof timeout === 'string' ? timeout.match(/^\$\{\{\s*matrix\.(\w+)\s*\}\}$/)?.[1] : undefined;
  return key ? combinations(job.strategy?.matrix).map((row) => row[key]) : [timeout];
}

const githubHosted = (labels: string[]) => labels.length === 1 && GITHUB_HOSTED.test(labels[0]);
const resolved = (runners: Record<string, Record<string, string[]>>) =>
  Object.values(runners).flatMap((byVariables) => Object.values(byVariables));
const kyoraJobs = jobs.filter(({ runners }) =>
  resolved(runners).some((labels) => labels[0] !== REUSABLE && !githubHosted(labels))
);
const routers = jobs.filter(({ id }) => id === 'route');
const groupsOf = (router: (typeof jobs)[number]) =>
  parseGroups(router.job.steps.find((step: any) => step.id === 'route').env.ROUTE_GROUPS) as { name: string; slots: number }[];
const everyGroup = Object.fromEntries(routers.flatMap((router) => groupsOf(router).map(({ name }) => [name, LABEL])));

describe('CI runner routing', () => {
  test('with nothing routed, every CI and End-to-end Linux job but the routers, the two shortest and the timings upload runs on its kyora class', () => {
    expect(Object.fromEntries(kyoraJobs.map(({ name, runners }) => [name, runners['push to main'].unset]))).toEqual({
      'ci.yml packages': ['kyora-4x8'],
      'ci.yml package-tests': ['kyora-8x8'],
      'ci.yml rust': ['kyora-8x16'],
      'ci.yml python': ['kyora-4x8'],
      'ci.yml route': ['kyora-4x8'],
      'e2e.yml route': ['kyora-4x8'],
      'e2e.yml scenarios (Corpus scenarios)': ['kyora-8x8'],
      'e2e.yml scenarios (Browser editors (1/2))': ['kyora-8x8'],
      'e2e.yml scenarios (Browser editors (2/2))': ['kyora-8x8'],
    });
    const hosted = jobs.filter((entry) => /^(ci|e2e)\.yml /.test(entry.name) && !kyoraJobs.includes(entry));
    expect(hosted.map(({ name }) => name).sort()).toEqual(['ci.yml audit', 'ci.yml native', 'e2e.yml publish']);
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

  test('each workflow has one kyora router, placed without its own outputs, whose groups are exactly the jobs that read them', () => {
    expect(routers.map(({ name }) => name).sort()).toEqual(['ci.yml route', 'e2e.yml route']);
    for (const router of routers) {
      const { name, file, job } = router;
      expect(router.runsOn(EVENTS['push to main'], {}, everyGroup), name).toEqual(router.runners['push to main'].unset);
      expect(job.permissions, name).toEqual({ actions: 'read', contents: 'read' });
      for (const step of job.steps) expect(step['continue-on-error'], `${name}: ${step.name ?? step.uses}`).toBe(true);
      const groups = groupsOf(router);
      expect(job.outputs, name).toEqual(
        Object.fromEntries(groups.map((group) => [group.name, `\${{ steps.route.outputs.${group.name} }}`]))
      );
      const routed = kyoraJobs.filter((entry) => entry.file === file && entry.id !== 'route');
      const readers = new Map<string, string[]>();
      for (const group of groups) {
        const moved = routed.filter((entry) => githubHosted(entry.runsOn(EVENTS['push to main'], {}, { [group.name]: LABEL })));
        expect(moved.reduce((slots, entry) => slots + entry.legs(), 0), `${name}: ${group.name}`).toBe(group.slots);
        for (const entry of moved) readers.set(entry.name, [...(readers.get(entry.name) ?? []), group.name]);
      }
      for (const entry of routed) {
        expect(readers.get(entry.name), entry.name).toHaveLength(1);
        expect(entry.job.needs, entry.name).toBe('route');
        // A failed router must not skip a required job, which would count as passed.
        expect(entry.job.if, entry.name).toBe('${{ !cancelled() }}');
      }
    }
  });

  test('each router declares the total slots of the other routers as ROUTE_PEER_SLOTS', () => {
    for (const router of routers) {
      const env = router.job.steps.find((step: any) => step.id === 'route').env;
      const peerSlots = routers
        .filter((other) => other !== router)
        .flatMap((other) => groupsOf(other))
        .reduce((sum, { slots }) => sum + slots, 0);
      const raw = String(env.ROUTE_PEER_SLOTS ?? '');
      const declared = /^\d+$/.test(raw)
        ? Number(raw)
        : Number(raw.match(/^\$\{\{\s*github\.event_name == 'workflow_dispatch' && '0' \|\| '(\d+)'\s*\}\}$/)?.[1]);
      expect(declared, router.name).toBe(peerSlots);
    }
  });

  test('routed jobs run GitHub-hosted unless the pull request is a fork, BO_CI_RUNNER pins them or BO_CI_SPILL is off', () => {
    for (const entry of kyoraJobs.filter(({ id }) => id !== 'route')) {
      for (const event of TRUSTED) {
        expect(entry.runsOn(EVENTS[event], {}, everyGroup), `${entry.name}: ${event}`).toEqual([LABEL]);
        expect(entry.runsOn(EVENTS[event], VARIABLES['BO_CI_RUNNER=kyora-8x8'], everyGroup), entry.name).toEqual(['kyora-8x8']);
        expect(entry.runsOn(EVENTS[event], VARIABLES['BO_CI_SPILL=off'], everyGroup), entry.name).toEqual(
          entry.runners['push to main'].unset
        );
      }
      for (const event of UNTRUSTED) {
        expect(entry.runsOn(EVENTS[event], {}, everyGroup), `${entry.name}: ${event}`).toEqual(['ubuntu-latest']);
      }
    }
  });

  test('kyora jobs read no secrets, write nothing, use no Docker and end before the class timeout', () => {
    for (const { name, id, workflow, triggers, job } of kyoraJobs) {
      expect(triggers, name).not.toContain('pull_request_target');
      expect(triggers, name).not.toContain('workflow_run');
      expect(JSON.stringify([workflow.env ?? null, workflow.defaults ?? null, job]), name).not.toMatch(/\bsecrets\b/);
      expect(workflow.permissions, name).toEqual({ contents: 'read' });
      // The router reads the Actions API; every other kyora job keeps the workflow's read-only contents token.
      expect(job.permissions, name).toEqual(id === 'route' ? { actions: 'read', contents: 'read' } : undefined);
      expect(job.services, name).toBeUndefined();
      expect(job.container, name).toBeUndefined();
      for (const step of job.steps) {
        expect(String(step.uses ?? ''), name).not.toMatch(/^docker:\/\/|^docker\/|^EmbarkStudios\/cargo-deny-action@/);
      }
      for (const timeout of timeouts(job)) expect(timeout, name).toBeLessThan(120);
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
