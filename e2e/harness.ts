import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import type { Format } from './corpus';
import { pythonWithBindings } from './python-env';

export type StageProfile = Record<string, number>;
export type Detail = Record<string, number | string | boolean>;

export interface OpTiming {
  op: string;
  actor?: string;
  e2eMs: number;
  internal?: StageProfile;
  detail?: Detail;
  error?: string;
}

export interface OpStats {
  count: number;
  totalMs: number;
  meanMs: number;
  p50Ms: number;
  p95Ms: number;
  maxMs: number;
  stagesMs?: StageProfile;
}

export interface ScenarioMeta {
  scenario: string;
  sample: string;
  sampleSha256: string;
  description: string;
  participants: string[];
}

export interface ScenarioRun extends ScenarioMeta {
  format: Format;
  status: 'passed' | 'skipped' | 'failed';
  reason?: string;
  loadMs: number;
  ops: OpTiming[];
  summary: {
    opCount: number;
    totalMs: number;
    byOp: Record<string, OpStats>;
    byActor: Record<string, number>;
  };
}

export interface Environment {
  platform: string;
  arch: string;
  cpu: string;
  cpus: number;
  bun: string;
  rustc: string;
  python: string;
}

export interface RecordedRun {
  schemaVersion: 3;
  format: Format;
  commit: string;
  dirty: boolean;
  recordedAt: string;
  environment: Environment;
  scenarios: ScenarioRun[];
}

const REGRESSION = {
  percent: 25,
  minimumMs: 10,
  repeatedMinimumMs: 2,
  repeats: 5,
};

export function measure<T>(
  op: string,
  run: () => T,
  internal?: () => StageProfile,
  meta: { actor?: string; detail?: Detail } = {}
): { value: T; timing: OpTiming } {
  const started = performance.now();
  const value = run();
  const e2eMs = performance.now() - started;
  const timing: OpTiming = { op, e2eMs };
  if (meta.actor) timing.actor = meta.actor;
  if (internal) timing.internal = internal();
  if (meta.detail) timing.detail = meta.detail;
  return { value, timing };
}

export interface Timer {
  op<T>(name: string, run: () => T, internal?: () => StageProfile): T;
  opAsync<T>(name: string, run: () => Promise<T>): Promise<T>;
}

export class ScenarioRecorder {
  readonly ops: OpTiming[] = [];
  loadMs = 0;

  constructor(readonly format: Format, readonly meta: ScenarioMeta) {}

  op<T>(
    name: string,
    run: () => T,
    internal?: () => StageProfile,
    meta: { actor?: string; detail?: Detail } = {}
  ): T {
    const started = performance.now();
    try {
      const { value, timing } = measure(
        name,
        run,
        internal,
        this.operationMeta(meta)
      );
      this.ops.push(timing);
      return value;
    } catch (error) {
      this.ops.push({
        op: name,
        e2eMs: performance.now() - started,
        ...this.operationMeta(meta),
        error: errorText(error),
      });
      throw error;
    }
  }

  async opAsync<T>(
    name: string,
    run: () => Promise<T>,
    meta: { actor?: string; detail?: Detail } = {}
  ): Promise<T> {
    const started = performance.now();
    try {
      const value = await run();
      this.ops.push({
        op: name,
        e2eMs: performance.now() - started,
        ...this.operationMeta(meta),
      });
      return value;
    } catch (error) {
      this.ops.push({
        op: name,
        e2eMs: performance.now() - started,
        ...this.operationMeta(meta),
        error: errorText(error),
      });
      throw error;
    }
  }

  record(timing: OpTiming): void {
    this.ops.push(timing);
  }

  async loadAsync<T>(run: () => Promise<T>, actor?: string): Promise<T> {
    const value = await this.opAsync('open', run, { actor });
    this.loadMs = this.ops[this.ops.length - 1].e2eMs;
    return value;
  }

  load<T>(run: () => T, actor?: string): T {
    const value = this.op('open', run, undefined, { actor });
    this.loadMs = this.ops[this.ops.length - 1].e2eMs;
    return value;
  }

  as(actor: string): ActorRecorder {
    return new ActorRecorder(this, actor);
  }

  private operationMeta(meta: { actor?: string; detail?: Detail }) {
    return {
      ...meta,
      actor:
        meta.actor ??
        (this.meta.participants.length === 1
          ? this.meta.participants[0]
          : undefined),
    };
  }

  finish(
    status: ScenarioRun['status'] = 'passed',
    reason?: string
  ): ScenarioRun {
    return {
      format: this.format,
      ...this.meta,
      status,
      reason,
      loadMs: this.loadMs,
      ops: [...this.ops],
      summary: summarize(this.ops),
    };
  }

  failed(error: unknown): ScenarioRun {
    return this.finish('failed', errorText(error));
  }

  skipped(reason: string): ScenarioRun {
    return {
      format: this.format,
      ...this.meta,
      status: 'skipped',
      reason,
      loadMs: 0,
      ops: [],
      summary: summarize([]),
    };
  }
}

export class ActorRecorder {
  constructor(
    private readonly recorder: ScenarioRecorder,
    readonly actor: string
  ) {}

  op<T>(
    name: string,
    run: () => T,
    internal?: () => StageProfile,
    detail?: Detail
  ): T {
    return this.recorder.op(name, run, internal, { actor: this.actor, detail });
  }

  load<T>(run: () => T): T {
    return this.recorder.load(run, this.actor);
  }

  opAsync<T>(name: string, run: () => Promise<T>, detail?: Detail): Promise<T> {
    return this.recorder.opAsync(name, run, { actor: this.actor, detail });
  }

  loadAsync<T>(run: () => Promise<T>): Promise<T> {
    return this.recorder.loadAsync(run, this.actor);
  }
}

function percentile(sorted: number[], p: number): number {
  if (sorted.length === 0) return 0;
  const index = Math.min(
    sorted.length - 1,
    Math.ceil((p / 100) * sorted.length) - 1
  );
  return sorted[Math.max(0, index)];
}

export function summarize(ops: OpTiming[]): ScenarioRun['summary'] {
  const byOp: Record<string, OpStats> = {};
  const byActor: Record<string, number> = {};
  const groups = new Map<string, OpTiming[]>();
  for (const op of ops) {
    const group = groups.get(op.op) ?? [];
    group.push(op);
    groups.set(op.op, group);
    if (op.actor) byActor[op.actor] = (byActor[op.actor] ?? 0) + op.e2eMs;
  }
  for (const [name, group] of groups) {
    const sorted = group.map((op) => op.e2eMs).sort((a, b) => a - b);
    const totalMs = sorted.reduce((sum, value) => sum + value, 0);
    const stats: OpStats = {
      count: sorted.length,
      totalMs,
      meanMs: totalMs / sorted.length,
      p50Ms: percentile(sorted, 50),
      p95Ms: percentile(sorted, 95),
      maxMs: sorted[sorted.length - 1],
    };
    const staged = group.filter((op) => op.internal);
    if (staged.length > 0) {
      const stagesMs: StageProfile = {};
      const counts: Record<string, number> = {};
      for (const op of staged) {
        for (const [stage, ms] of Object.entries(op.internal!)) {
          stagesMs[stage] = (stagesMs[stage] ?? 0) + ms;
          counts[stage] = (counts[stage] ?? 0) + 1;
        }
      }
      for (const stage of Object.keys(stagesMs))
        stagesMs[stage] /= counts[stage];
      stats.stagesMs = stagesMs;
    }
    byOp[name] = stats;
  }
  return {
    opCount: ops.length,
    totalMs: ops.reduce((sum, op) => sum + op.e2eMs, 0),
    byOp,
    byActor,
  };
}

export function baselineDir(): string {
  return (
    process.env.BETTEROFFICE_E2E_BASELINE ??
    path.resolve(import.meta.dir, '../.source/e2e/baseline')
  );
}

export function errorText(error: unknown): string {
  return error instanceof Error ? error.stack ?? error.message : String(error);
}

export function scenarioKey(run: ScenarioMeta): string {
  return run.scenario + '/' + run.sample;
}

export function assertComplete(run: RecordedRun): void {
  if (run.scenarios.length === 0)
    throw new Error(run.format + ': no scenarios');
  const incomplete = run.scenarios.filter(
    (scenario) => scenario.status !== 'passed'
  );
  if (incomplete.length)
    throw new Error(
      run.format +
        ': incomplete run: ' +
        incomplete.map((s) => scenarioKey(s) + ' (' + s.status + ')').join(', ')
    );
}

export function assertComparable(
  previous: RecordedRun,
  current: RecordedRun
): void {
  assertComplete(previous);
  assertComplete(current);
  if (previous.format !== current.format) throw new Error('format mismatch');
  for (const field of [
    'platform',
    'arch',
    'cpu',
    'cpus',
    'bun',
    'rustc',
    'python',
  ] as const) {
    if (previous.environment[field] !== current.environment[field])
      throw new Error('incomparable environment: ' + field);
  }
  const inventory = (run: RecordedRun) =>
    run.scenarios
      .map((s) =>
        JSON.stringify([
          scenarioKey(s),
          s.sampleSha256,
          [...s.participants].sort(),
        ])
      )
      .sort();
  if (
    JSON.stringify(inventory(previous)) !== JSON.stringify(inventory(current))
  )
    throw new Error(
      'scenario, sample, or participant inventory changed; record a new baseline'
    );
}

export function actorStats(
  run: ScenarioRun
): Map<string, { actor: string; op: string; stats: OpStats; errors: number }> {
  const groups = new Map<string, OpTiming[]>();
  for (const timing of run.ops) {
    const key = JSON.stringify([timing.actor ?? 'unattributed', timing.op]);
    const group = groups.get(key) ?? [];
    group.push(timing);
    groups.set(key, group);
  }
  return new Map(
    [...groups].map(([key, ops]) => [
      key,
      {
        actor: ops[0].actor ?? 'unattributed',
        op: ops[0].op,
        stats: summarize(ops).byOp[ops[0].op],
        errors: ops.filter((op) => op.error !== undefined).length,
      },
    ])
  );
}

export function regressions(
  previous: RecordedRun,
  current: RecordedRun
): string[] {
  assertComparable(previous, current);
  const failures: string[] = [];
  const now = new Map(current.scenarios.map((run) => [scenarioKey(run), run]));
  for (const run of previous.scenarios) {
    const beforeGroups = actorStats(run);
    const afterGroups = actorStats(now.get(scenarioKey(run))!);
    if (
      JSON.stringify([...beforeGroups.keys()].sort()) !==
      JSON.stringify([...afterGroups.keys()].sort())
    ) {
      throw new Error(scenarioKey(run) + ': operation inventory changed');
    }
    for (const [key, beforeGroup] of beforeGroups) {
      const afterGroup = afterGroups.get(key)!;
      const before = beforeGroup.stats;
      const after = afterGroup.stats;
      const id =
        run.format +
        '/' +
        scenarioKey(run) +
        ':' +
        beforeGroup.actor +
        ':' +
        beforeGroup.op;
      if (
        before.count !== after.count ||
        beforeGroup.errors !== afterGroup.errors
      )
        throw new Error(id + ': operation count or outcome changed');
      const metrics =
        before.count >= 20
          ? (['p50Ms', 'p95Ms'] as const)
          : (['p50Ms'] as const);
      for (const metric of metrics) {
        const delta = after[metric] - before[metric];
        const percent =
          before[metric] === 0 ? Infinity : (delta / before[metric]) * 100;
        const floor =
          before.count >= REGRESSION.repeats
            ? REGRESSION.repeatedMinimumMs
            : REGRESSION.minimumMs;
        if (delta > floor && percent > REGRESSION.percent)
          failures.push(
            id +
              ': ' +
              metric +
              ' ' +
              before[metric].toFixed(2) +
              'ms -> ' +
              after[metric].toFixed(2) +
              'ms (+' +
              percent.toFixed(0) +
              '%)'
          );
      }
    }
  }
  return failures;
}

export function parseRecordedRun(value: unknown): RecordedRun {
  const fail = (message: string): never => {
    throw new Error('invalid e2e results: ' + message);
  };
  const object = (value: unknown): value is Record<string, any> =>
    typeof value === 'object' && value !== null && !Array.isArray(value);
  const finite = (value: unknown): value is number =>
    typeof value === 'number' && Number.isFinite(value) && value >= 0;
  if (!object(value) || value.schemaVersion !== 3)
    fail('expected schemaVersion 3');
  const run = value as Record<string, any>;
  if (!['docx', 'xlsx', 'pptx'].includes(run.format)) fail('format');
  if (
    typeof run.commit !== 'string' ||
    typeof run.recordedAt !== 'string' ||
    typeof run.dirty !== 'boolean'
  )
    fail('run metadata');
  if (
    !object(run.environment) ||
    !['platform', 'arch', 'cpu', 'bun', 'rustc', 'python'].every(
      (key) => typeof run.environment[key] === 'string'
    ) ||
    !Number.isInteger(run.environment.cpus) ||
    run.environment.cpus < 1
  )
    fail('environment');
  if (!Array.isArray(run.scenarios) || run.scenarios.length === 0)
    fail('empty scenario inventory');
  const seen = new Set<string>();
  for (const scenario of run.scenarios) {
    if (
      !object(scenario) ||
      scenario.format !== run.format ||
      !['scenario', 'sample', 'sampleSha256', 'description'].every(
        (key) => typeof scenario[key] === 'string'
      )
    )
      fail('scenario metadata');
    if (!/^[a-f0-9]{64}$/.test(scenario.sampleSha256)) fail('sample hash');
    if (
      !Array.isArray(scenario.participants) ||
      !scenario.participants.length ||
      !scenario.participants.every((p: unknown) => typeof p === 'string')
    )
      fail('participants');
    if (
      !['passed', 'skipped', 'failed'].includes(scenario.status) ||
      (scenario.status !== 'passed' && typeof scenario.reason !== 'string')
    )
      fail('scenario status');
    if (!finite(scenario.loadMs) || !Array.isArray(scenario.ops))
      fail('timings');
    const key = scenarioKey(scenario as ScenarioRun);
    if (seen.has(key)) fail('duplicate scenario ' + key);
    seen.add(key);
    for (const op of scenario.ops) {
      if (!object(op) || typeof op.op !== 'string' || !finite(op.e2eMs))
        fail('operation timing');
      if (
        op.actor !== undefined &&
        (typeof op.actor !== 'string' ||
          !scenario.participants.includes(op.actor))
      )
        fail('operation actor');
      if (op.error !== undefined && typeof op.error !== 'string')
        fail('operation error');
      if (
        op.internal !== undefined &&
        (!object(op.internal) || !Object.values(op.internal).every(finite))
      )
        fail('stage timing');
    }
    if (
      JSON.stringify(scenario.summary) !==
      JSON.stringify(summarize(scenario.ops))
    )
      fail('summary does not match raw operations');
  }
  return run as unknown as RecordedRun;
}

function stagesText(stages: StageProfile | undefined): string {
  if (!stages) return '';
  return (
    '  {' +
    Object.entries(stages)
      .map(([k, v]) => `${k} ${v.toFixed(2)}`)
      .join(', ') +
    '}'
  );
}

export function describeRun(run: ScenarioRun): string {
  const head = `${run.format} ${run.scenario} @ ${
    run.sample
  }  [${run.participants.join(', ')}]`;
  if (run.status !== 'passed') return `${head}  ${run.status}: ${run.reason}`;
  const lines = [
    `${head}  load ${run.loadMs.toFixed(
      1
    )}ms  total ${run.summary.totalMs.toFixed(1)}ms over ${
      run.summary.opCount
    } ops`,
  ];
  for (const [op, stats] of Object.entries(run.summary.byOp)) {
    const spread =
      stats.count > 1
        ? ` x${stats.count} p50 ${stats.p50Ms.toFixed(
            2
          )} p95 ${stats.p95Ms.toFixed(2)} max ${stats.maxMs.toFixed(2)}`
        : '';
    lines.push(
      `  ${op.padEnd(36)} ${stats.meanMs
        .toFixed(2)
        .padStart(8)}ms${spread}${stagesText(stats.stagesMs)}`
    );
  }
  return lines.join('\n');
}

export function e2eEnabled(): boolean {
  return process.env.BETTEROFFICE_E2E === '1';
}

export function environment(): Environment {
  const rustc = Bun.spawnSync(['rustc', '--version']);
  const ready = pythonWithBindings();
  const python =
    'python' in ready
      ? Bun.spawnSync([ready.python, '--version']).stdout.toString().trim()
      : 'unavailable';
  return {
    platform: process.platform,
    arch: process.arch,
    cpu: os.cpus()[0]?.model ?? 'unknown',
    cpus: os.cpus().length,
    bun: Bun.version,
    python,
    rustc: rustc.success ? rustc.stdout.toString().trim() : 'unavailable',
  };
}

export function recordedRun(
  format: Format,
  scenarios: ScenarioRun[]
): RecordedRun {
  const commit = Bun.spawnSync(['git', 'rev-parse', 'HEAD'], {
    cwd: import.meta.dir,
  });
  const status = Bun.spawnSync(['git', 'status', '--porcelain'], {
    cwd: import.meta.dir,
  });
  return {
    schemaVersion: 3,
    format,
    dirty: !status.success || status.stdout.length > 0,
    commit: commit.success ? commit.stdout.toString().trim() : 'unknown',
    recordedAt: new Date().toISOString(),
    environment: environment(),
    scenarios,
  };
}

export function finishFormat(format: Format, scenarios: ScenarioRun[]): void {
  for (const run of scenarios) console.log(describeRun(run));
  const run = recordedRun(format, scenarios);
  const output = process.env.BETTEROFFICE_E2E_OUTPUT;
  if (output) {
    fs.mkdirSync(output, { recursive: true });
    fs.writeFileSync(
      path.join(output, format + '.json'),
      JSON.stringify(run, null, 2) + '\n'
    );
  }
  if (process.env.GITHUB_STEP_SUMMARY)
    fs.appendFileSync(
      process.env.GITHUB_STEP_SUMMARY,
      summaryMarkdown(format, scenarios)
    );
  if (process.env.BETTEROFFICE_E2E_ALLOW_SKIP !== '1' || process.env.CI)
    assertComplete(run);
  else if (scenarios.some((scenario) => scenario.status === 'failed'))
    throw new Error(format + ': failed or unexecuted scenarios');
}

export function summaryMarkdown(
  format: Format,
  scenarios: ScenarioRun[]
): string {
  const lines: string[] = [];
  for (const run of scenarios) {
    lines.push(`### ${format} ${run.scenario} on \`${run.sample}\``, '');
    if (run.status !== 'passed') {
      lines.push(`${run.status}: ${run.reason}`, '');
      continue;
    }
    lines.push(
      `${run.description} (${run.participants.join(
        ', '
      )}); load ${run.loadMs.toFixed(1)} ms, ${
        run.summary.opCount
      } ops in ${run.summary.totalMs.toFixed(1)} ms`,
      ''
    );
    lines.push(
      '| op | n | p50 ms | p95 ms | max ms | stages (mean ms) |',
      '| --- | ---: | ---: | ---: | ---: | --- |'
    );
    for (const [op, stats] of Object.entries(run.summary.byOp)) {
      const stages = stats.stagesMs
        ? Object.entries(stats.stagesMs)
            .map(([k, v]) => `${k} ${v.toFixed(2)}`)
            .join(', ')
        : '';
      lines.push(
        `| ${op} | ${stats.count} | ${stats.p50Ms.toFixed(
          2
        )} | ${stats.p95Ms.toFixed(2)} | ${stats.maxMs.toFixed(
          2
        )} | ${stages} |`
      );
    }
    lines.push('');
  }
  return lines.join('\n');
}
