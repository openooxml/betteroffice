/**
 * Per-operation timing for the end-to-end scenarios. Every operation is
 * measured on its own: the wall-clock latency across the boundary it crosses
 * (wasm call, Python round trip, update exchange) plus whatever stage
 * breakdown the engine reports for it. A scenario run keeps the raw ops and
 * per-op distributions; a format's run is compared against its recorded one.
 */

import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import type { Format } from './corpus';

/** Engine-reported stage latencies, in milliseconds, keyed by stage. */
export type StageProfile = Record<string, number>;
export type Detail = Record<string, number | string | boolean>;

export interface OpTiming {
  op: string;
  /** Which participant issued it, for multi-editor and cross-SDK scenarios. */
  actor?: string;
  e2eMs: number;
  internal?: StageProfile;
  detail?: Detail;
}

export interface OpStats {
  count: number;
  totalMs: number;
  meanMs: number;
  p50Ms: number;
  p95Ms: number;
  maxMs: number;
  /** Mean of each engine stage over the ops that reported it. */
  stagesMs?: StageProfile;
}

export interface ScenarioMeta {
  scenario: string;
  sample: string;
  description: string;
  participants: string[];
}

export interface ScenarioRun extends ScenarioMeta {
  format: Format;
  status: 'passed' | 'skipped';
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
}

export interface RecordedRun {
  schemaVersion: 2;
  commit: string;
  recordedAt: string;
  environment: Environment;
  scenarios: ScenarioRun[];
}

/**
 * A regression needs a 25% slower median and a real absolute move. One-shot
 * operations carry single-sample noise, so they need 10ms; an operation the
 * scenario repeats has a stable median and only needs 2ms.
 */
const REGRESSION = { percent: 25, minimumMs: 10, repeatedMinimumMs: 2, repeats: 5 };

/** Each format keeps its own recorded run so the suites stay independent. */
export function resultsPath(format: Format): string {
  return path.resolve(import.meta.dir, 'results', `${format}.json`);
}

/** Time one operation across the boundary; `internal` is read after it returns. */
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

/** What scenario helpers need from a recorder; both recorder kinds qualify. */
export interface Timer {
  op<T>(name: string, run: () => T, internal?: () => StageProfile): T;
  opAsync<T>(name: string, run: () => Promise<T>): Promise<T>;
}

/** Collects one scenario's operations, in order, on one sample. */
export class ScenarioRecorder {
  readonly ops: OpTiming[] = [];
  loadMs = 0;

  constructor(readonly format: Format, readonly meta: ScenarioMeta) {}

  op<T>(name: string, run: () => T, internal?: () => StageProfile, meta: { actor?: string; detail?: Detail } = {}): T {
    const { value, timing } = measure(name, run, internal, meta);
    this.ops.push(timing);
    return value;
  }

  /** Times an awaited operation the way `op` times a synchronous one. */
  async opAsync<T>(name: string, run: () => Promise<T>, meta: { actor?: string; detail?: Detail } = {}): Promise<T> {
    const started = performance.now();
    const value = await run();
    this.ops.push({ op: name, e2eMs: performance.now() - started, ...(meta.actor ? { actor: meta.actor } : {}), ...(meta.detail ? { detail: meta.detail } : {}) });
    return value;
  }

  /** An operation whose latency was measured elsewhere (async boundaries). */
  record(timing: OpTiming): void {
    this.ops.push(timing);
  }

  /** The awaited document open that starts the scenario; also kept as `loadMs`. */
  async loadAsync<T>(run: () => Promise<T>, actor?: string): Promise<T> {
    const value = await this.opAsync('open', run, { actor });
    this.loadMs = this.ops[this.ops.length - 1].e2eMs;
    return value;
  }

  /** The document open that starts the scenario; also kept as `loadMs`. */
  load<T>(run: () => T, actor?: string): T {
    const value = this.op('open', run, undefined, { actor });
    this.loadMs = this.ops[this.ops.length - 1].e2eMs;
    return value;
  }

  /** A view that stamps every op with one participant. */
  as(actor: string): ActorRecorder {
    return new ActorRecorder(this, actor);
  }

  finish(): ScenarioRun {
    return { format: this.format, ...this.meta, status: 'passed', loadMs: this.loadMs, ops: this.ops, summary: summarize(this.ops) };
  }

  skipped(reason: string): ScenarioRun {
    return { format: this.format, ...this.meta, status: 'skipped', reason, loadMs: 0, ops: [], summary: summarize([]) };
  }
}

export class ActorRecorder {
  constructor(private readonly recorder: ScenarioRecorder, readonly actor: string) {}

  op<T>(name: string, run: () => T, internal?: () => StageProfile, detail?: Detail): T {
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
  const index = Math.min(sorted.length - 1, Math.ceil((p / 100) * sorted.length) - 1);
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
      for (const op of staged) {
        for (const [stage, ms] of Object.entries(op.internal!)) stagesMs[stage] = (stagesMs[stage] ?? 0) + ms / staged.length;
      }
      stats.stagesMs = stagesMs;
    }
    byOp[name] = stats;
  }
  return { opCount: ops.length, totalMs: ops.reduce((sum, op) => sum + op.e2eMs, 0), byOp, byActor };
}

/** The run recorded on the default branch, when there is one. */
export function readRecordedRun(format: Format): RecordedRun | undefined {
  const file = resultsPath(format);
  if (!fs.existsSync(file)) return undefined;
  const run = JSON.parse(fs.readFileSync(file, 'utf8')) as RecordedRun;
  return run.schemaVersion === 2 ? run : undefined;
}

/** Operations whose median got slower than the recorded run allows. */
export function regressions(previous: RecordedRun, current: ScenarioRun[]): string[] {
  const failures: string[] = [];
  const key = (run: ScenarioRun, op: string) => `${run.format}/${run.scenario}/${run.sample}:${op}`;
  const now = new Map<string, OpStats>();
  for (const run of current) for (const [op, stats] of Object.entries(run.summary.byOp)) now.set(key(run, op), stats);
  for (const run of previous.scenarios) {
    if (run.status !== 'passed') continue;
    for (const [op, before] of Object.entries(run.summary.byOp)) {
      const id = key(run, op);
      const after = now.get(id);
      if (!after) {
        failures.push(`${id}: missing from the current run`);
        continue;
      }
      const delta = after.p50Ms - before.p50Ms;
      const percent = before.p50Ms === 0 ? Infinity : (delta / before.p50Ms) * 100;
      const floor = before.count >= REGRESSION.repeats ? REGRESSION.repeatedMinimumMs : REGRESSION.minimumMs;
      if (delta > floor && percent > REGRESSION.percent) {
        failures.push(`${id}: p50 ${before.p50Ms.toFixed(2)}ms -> ${after.p50Ms.toFixed(2)}ms (+${percent.toFixed(0)}%)`);
      }
    }
  }
  return failures;
}

function stagesText(stages: StageProfile | undefined): string {
  if (!stages) return '';
  return '  {' + Object.entries(stages).map(([k, v]) => `${k} ${v.toFixed(2)}`).join(', ') + '}';
}

/** One line per operation group, the way a run reads in a terminal. */
export function describeRun(run: ScenarioRun): string {
  const head = `${run.format} ${run.scenario} @ ${run.sample}  [${run.participants.join(', ')}]`;
  if (run.status === 'skipped') return `${head}  skipped: ${run.reason}`;
  const lines = [`${head}  load ${run.loadMs.toFixed(1)}ms  total ${run.summary.totalMs.toFixed(1)}ms over ${run.summary.opCount} ops`];
  for (const [op, stats] of Object.entries(run.summary.byOp)) {
    const spread = stats.count > 1 ? ` x${stats.count} p50 ${stats.p50Ms.toFixed(2)} p95 ${stats.p95Ms.toFixed(2)} max ${stats.maxMs.toFixed(2)}` : '';
    lines.push(`  ${op.padEnd(36)} ${stats.meanMs.toFixed(2).padStart(8)}ms${spread}${stagesText(stats.stagesMs)}`);
  }
  return lines.join('\n');
}

/**
 * `BETTEROFFICE_E2E` selects the mode: `1` runs and prints, `record` also
 * writes the run as the new baseline, `compare` holds it against the baseline
 * and fails on regressions. Unset, the suites skip: they need the network and
 * built wasm bundles.
 */
export function e2eEnabled(): boolean {
  return ['1', 'record', 'compare'].includes(process.env.BETTEROFFICE_E2E ?? '');
}

export function shouldRecord(): boolean {
  return process.env.BETTEROFFICE_E2E === 'record';
}

export function shouldCompare(): boolean {
  return process.env.BETTEROFFICE_E2E === 'compare';
}

export function environment(): Environment {
  return { platform: process.platform, arch: process.arch, cpu: os.cpus()[0]?.model ?? 'unknown', cpus: os.cpus().length, bun: Bun.version };
}

export function writeRecordedRun(format: Format, scenarios: ScenarioRun[]): void {
  fs.mkdirSync(path.dirname(resultsPath(format)), { recursive: true });
  fs.writeFileSync(resultsPath(format), serialize(scenarios));
}

function serialize(scenarios: ScenarioRun[]): string {
  const run: RecordedRun = { schemaVersion: 2, commit: currentCommit(), recordedAt: new Date().toISOString(), environment: environment(), scenarios };
  return JSON.stringify(run, null, 2) + '\n';
}

function currentCommit(): string {
  const result = Bun.spawnSync(['git', 'rev-parse', 'HEAD'], { cwd: import.meta.dir });
  return result.success ? result.stdout.toString().trim() : 'unknown';
}

/**
 * Close one format's run: print every scenario, publish it where the
 * environment asks (`BETTEROFFICE_E2E_OUTPUT` dir, `GITHUB_STEP_SUMMARY`),
 * then record or compare per mode. Throws with the offending operations on a
 * regression.
 */
export function finishFormat(format: Format, scenarios: ScenarioRun[]): void {
  for (const run of scenarios) console.log(describeRun(run));
  const output = process.env.BETTEROFFICE_E2E_OUTPUT;
  if (output) {
    fs.mkdirSync(output, { recursive: true });
    fs.writeFileSync(path.join(output, `${format}.json`), serialize(scenarios));
  }
  if (process.env.GITHUB_STEP_SUMMARY) {
    fs.appendFileSync(process.env.GITHUB_STEP_SUMMARY, summaryMarkdown(format, scenarios));
  }
  if (shouldRecord()) {
    writeRecordedRun(format, scenarios);
    return;
  }
  if (!shouldCompare()) return;
  const previous = readRecordedRun(format);
  if (!previous) throw new Error(`no recorded ${format} run at ${resultsPath(format)}; run with BETTEROFFICE_E2E=record first`);
  const slower = regressions(previous, scenarios);
  if (slower.length > 0) {
    throw new Error(`${format} e2e regressions against ${previous.commit}:\n  ${slower.join('\n  ')}`);
  }
}

/** One table per scenario run: operation, count, median, p95, max, stages. */
export function summaryMarkdown(format: Format, scenarios: ScenarioRun[]): string {
  const lines: string[] = [];
  for (const run of scenarios) {
    lines.push(`### ${format} ${run.scenario} on \`${run.sample}\``, '');
    if (run.status === 'skipped') {
      lines.push(`skipped: ${run.reason}`, '');
      continue;
    }
    lines.push(`${run.description} (${run.participants.join(', ')}); load ${run.loadMs.toFixed(1)} ms, ${run.summary.opCount} ops in ${run.summary.totalMs.toFixed(1)} ms`, '');
    lines.push('| op | n | p50 ms | p95 ms | max ms | stages (mean ms) |', '| --- | ---: | ---: | ---: | ---: | --- |');
    for (const [op, stats] of Object.entries(run.summary.byOp)) {
      const stages = stats.stagesMs ? Object.entries(stats.stagesMs).map(([k, v]) => `${k} ${v.toFixed(2)}`).join(', ') : '';
      lines.push(`| ${op} | ${stats.count} | ${stats.p50Ms.toFixed(2)} | ${stats.p95Ms.toFixed(2)} | ${stats.maxMs.toFixed(2)} | ${stages} |`);
    }
    lines.push('');
  }
  return lines.join('\n');
}
