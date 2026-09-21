/**
 * Per-operation timing for the end-to-end runs. Every operation is measured
 * on its own: the wall-clock latency across the wasm boundary, and whatever
 * stage breakdown the engine reports for it. A run is a list of those, keyed
 * by format, sample and operation, and is compared against the recorded run.
 */

import * as fs from 'node:fs';
import * as path from 'node:path';
import type { Format } from './corpus';

/** Engine-reported stage latencies, in milliseconds, keyed by stage. */
export type StageProfile = Record<string, number>;

export interface OpTiming {
  op: string;
  e2eMs: number;
  internal?: StageProfile;
}

export interface SampleRun {
  format: Format;
  sample: string;
  loadMs: number;
  ops: OpTiming[];
}

export interface RecordedRun {
  schemaVersion: 1;
  commit: string;
  samples: SampleRun[];
}

const REGRESSION = { percent: 25, minimumMs: 2 };

/** Each format keeps its own recorded run so the suites stay independent. */
export function resultsPath(format: Format): string {
  return path.resolve(import.meta.dir, 'results', `${format}.json`);
}

/** Time one operation across the boundary; `internal` is read after it returns. */
export function measure<T>(op: string, run: () => T, internal?: () => StageProfile): { value: T; timing: OpTiming } {
  const started = performance.now();
  const value = run();
  const e2eMs = performance.now() - started;
  const timing: OpTiming = { op, e2eMs };
  if (internal) timing.internal = internal();
  return { value, timing };
}

/** Collects one sample's operations in order. */
export class SampleRecorder {
  readonly ops: OpTiming[] = [];
  constructor(readonly format: Format, readonly sample: string, readonly loadMs: number) {}

  op<T>(name: string, run: () => T, internal?: () => StageProfile): T {
    const { value, timing } = measure(name, run, internal);
    this.ops.push(timing);
    return value;
  }

  finish(): SampleRun {
    return { format: this.format, sample: this.sample, loadMs: this.loadMs, ops: this.ops };
  }
}

/** The run recorded on the default branch, when there is one. */
export function readRecordedRun(format: Format): RecordedRun | undefined {
  const file = resultsPath(format);
  if (!fs.existsSync(file)) return undefined;
  return JSON.parse(fs.readFileSync(file, 'utf8')) as RecordedRun;
}

/** Operations that got slower than the recorded run allows. */
export function regressions(previous: RecordedRun, current: SampleRun[]): string[] {
  const failures: string[] = [];
  const key = (run: SampleRun, op: OpTiming, index: number) => `${run.format}/${run.sample}#${index}:${op.op}`;
  const now = new Map<string, number>();
  for (const run of current) run.ops.forEach((op, index) => now.set(key(run, op, index), op.e2eMs));
  for (const run of previous.samples) {
    run.ops.forEach((before, index) => {
      const id = key(run, before, index);
      const after = now.get(id);
      if (after === undefined) {
        failures.push(`${id}: missing from the current run`);
        return;
      }
      const delta = after - before.e2eMs;
      const percent = before.e2eMs === 0 ? Infinity : (delta / before.e2eMs) * 100;
      if (delta > REGRESSION.minimumMs && percent > REGRESSION.percent) {
        failures.push(`${id}: ${before.e2eMs.toFixed(2)}ms -> ${after.toFixed(2)}ms (+${percent.toFixed(0)}%)`);
      }
    });
  }
  return failures;
}

/** One line per operation, the way a run reads in a terminal. */
export function describeRun(run: SampleRun): string {
  const lines = [`${run.format} ${run.sample}  load ${run.loadMs.toFixed(1)}ms`];
  for (const op of run.ops) {
    const stages = op.internal
      ? '  {' + Object.entries(op.internal).map(([k, v]) => `${k} ${v.toFixed(2)}`).join(', ') + '}'
      : '';
    lines.push(`  ${op.op.padEnd(34)} ${op.e2eMs.toFixed(2).padStart(8)}ms${stages}`);
  }
  return lines.join('\n');
}

/** Whether this process should run the corpus-backed suites at all. */
export function e2eEnabled(): boolean {
  return process.env.BETTEROFFICE_E2E === '1' || process.env.BETTEROFFICE_E2E === 'record';
}

export function shouldRecord(): boolean {
  return process.env.BETTEROFFICE_E2E === 'record';
}

export function writeRecordedRun(format: Format, samples: SampleRun[]): void {
  const run: RecordedRun = { schemaVersion: 1, commit: currentCommit(), samples };
  fs.mkdirSync(path.dirname(resultsPath(format)), { recursive: true });
  fs.writeFileSync(resultsPath(format), JSON.stringify(run, null, 2) + '\n');
}

function currentCommit(): string {
  const result = Bun.spawnSync(['git', 'rev-parse', 'HEAD'], { cwd: import.meta.dir });
  return result.success ? result.stdout.toString().trim() : 'unknown';
}

/**
 * Close one format's run: print every sample, then either record it or hold it
 * against the recorded run. Throws with the offending operations on a regression.
 */
export function finishFormat(format: Format, samples: SampleRun[]): void {
  for (const sample of samples) console.log(describeRun(sample));
  if (shouldRecord()) {
    writeRecordedRun(format, samples);
    return;
  }
  const previous = readRecordedRun(format);
  if (!previous) return;
  const slower = regressions(previous, samples);
  if (slower.length > 0) {
    throw new Error(`${format} e2e regressions against ${previous.commit}:\n  ${slower.join('\n  ')}`);
  }
}
