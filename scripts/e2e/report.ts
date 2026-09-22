/**
 * Reads recorded or exported e2e runs and prints them for people or programs:
 * `bun scripts/e2e/report.ts [dir]` renders markdown, `--json` flattens every
 * op into one row, `--diff <before> <after>` lists what got slower or faster.
 */

import * as fs from 'node:fs';
import * as path from 'node:path';

import type { RecordedRun, ScenarioRun } from './harness';
import { regressions } from './harness';

const RESULTS = path.resolve(import.meta.dir, 'results');

export interface OpRow {
  format: string;
  scenario: string;
  sample: string;
  participants: string;
  op: string;
  count: number;
  meanMs: number;
  p50Ms: number;
  p95Ms: number;
  maxMs: number;
  stagesMs?: Record<string, number>;
}

export function readRuns(dir = RESULTS): RecordedRun[] {
  if (!fs.existsSync(dir)) return [];
  return fs
    .readdirSync(dir)
    .filter((name) => name.endsWith('.json'))
    .sort()
    .map((name) => JSON.parse(fs.readFileSync(path.join(dir, name), 'utf8')) as RecordedRun)
    .filter((run) => run.schemaVersion === 2);
}

export function rows(runs: RecordedRun[]): OpRow[] {
  const out: OpRow[] = [];
  for (const run of runs) {
    for (const scenario of run.scenarios) {
      if (scenario.status !== 'passed') continue;
      for (const [op, stats] of Object.entries(scenario.summary.byOp)) {
        out.push({
          format: scenario.format,
          scenario: scenario.scenario,
          sample: scenario.sample,
          participants: scenario.participants.join('+'),
          op,
          count: stats.count,
          meanMs: stats.meanMs,
          p50Ms: stats.p50Ms,
          p95Ms: stats.p95Ms,
          maxMs: stats.maxMs,
          stagesMs: stats.stagesMs,
        });
      }
    }
  }
  return out;
}

function slowestOp(run: ScenarioRun): string {
  const entries = Object.entries(run.summary.byOp);
  if (entries.length === 0) return '';
  const [op, stats] = entries.reduce((best, entry) => (entry[1].p50Ms > best[1].p50Ms ? entry : best));
  return `${op} (${stats.p50Ms.toFixed(1)} ms)`;
}

export function markdown(runs: RecordedRun[]): string {
  const lines: string[] = [];
  for (const run of runs) {
    const format = run.scenarios[0]?.format ?? '?';
    lines.push(`## ${format} (${run.commit.slice(0, 8)}, ${run.recordedAt}, ${run.environment.cpu} x${run.environment.cpus}, bun ${run.environment.bun})`, '');
    lines.push('| scenario | sample | participants | status | ops | load ms | total ms | slowest op (p50) |', '| --- | --- | --- | --- | ---: | ---: | ---: | --- |');
    for (const scenario of run.scenarios) {
      lines.push(
        `| ${scenario.scenario} | ${scenario.sample} | ${scenario.participants.join('+')} | ${scenario.status}${scenario.reason ? `: ${scenario.reason}` : ''} | ${scenario.summary.opCount} | ${scenario.loadMs.toFixed(1)} | ${scenario.summary.totalMs.toFixed(1)} | ${slowestOp(scenario)} |`
      );
    }
    lines.push('');
  }
  const slowest = rows(runs)
    .sort((a, b) => b.p50Ms - a.p50Ms)
    .slice(0, 15);
  lines.push('## Slowest operations across formats (p50)', '', '| format | scenario | sample | op | n | p50 ms | p95 ms | stages (mean ms) |', '| --- | --- | --- | --- | ---: | ---: | ---: | --- |');
  for (const row of slowest) {
    const stages = row.stagesMs ? Object.entries(row.stagesMs).map(([k, v]) => `${k} ${v.toFixed(1)}`).join(', ') : '';
    lines.push(`| ${row.format} | ${row.scenario} | ${row.sample} | ${row.op} | ${row.count} | ${row.p50Ms.toFixed(2)} | ${row.p95Ms.toFixed(2)} | ${stages} |`);
  }
  lines.push('');
  return lines.join('\n');
}

/** Ops whose median moved by more than the harness thresholds, both ways. */
export function diff(before: RecordedRun[], after: RecordedRun[]): { slower: string[]; faster: string[] } {
  const slower: string[] = [];
  const faster: string[] = [];
  for (const previous of before) {
    const format = previous.scenarios[0]?.format;
    const current = after.find((run) => run.scenarios[0]?.format === format);
    if (!current) continue;
    slower.push(...regressions(previous, current.scenarios));
    faster.push(...regressions(current, previous.scenarios).filter((line) => !line.endsWith('missing from the current run')));
  }
  return { slower, faster };
}

if (import.meta.main) {
  const args = process.argv.slice(2);
  if (args[0] === '--diff') {
    const { slower, faster } = diff(readRuns(args[1]), readRuns(args[2]));
    console.log(`slower (${slower.length}):\n  ${slower.join('\n  ') || 'none'}`);
    console.log(`faster (${faster.length}):\n  ${faster.join('\n  ') || 'none'}`);
    process.exit(slower.length > 0 ? 1 : 0);
  }
  const json = args.includes('--json');
  const dir = args.find((arg) => !arg.startsWith('--'));
  const runs = readRuns(dir);
  console.log(json ? JSON.stringify(rows(runs), null, 2) : markdown(runs));
}
