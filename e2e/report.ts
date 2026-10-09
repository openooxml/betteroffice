import * as fs from 'node:fs';
import * as path from 'node:path';

import {
  actorStats,
  baselineDir,
  parseRecordedRun,
  regressions,
} from './harness';
import type { RecordedRun } from './harness';

export function readRuns(dir = baselineDir()): RecordedRun[] {
  if (!fs.existsSync(dir))
    throw new Error('results directory does not exist: ' + dir);
  const files = fs
    .readdirSync(dir)
    .filter((name) => name.endsWith('.json'))
    .sort();
  if (!files.length) throw new Error('no results in ' + dir);
  const runs = files.map((name) => {
    try {
      return parseRecordedRun(
        JSON.parse(fs.readFileSync(path.join(dir, name), 'utf8'))
      );
    } catch (error) {
      throw new Error(path.join(dir, name) + ': ' + String(error));
    }
  });
  if (new Set(runs.map((run) => run.format)).size !== runs.length)
    throw new Error('duplicate format in ' + dir);
  return runs;
}

export function rows(runs: RecordedRun[]) {
  return runs.flatMap((run) =>
    run.scenarios.flatMap((scenario) =>
      [...actorStats(scenario).values()].map(
        ({ actor, op, stats, errors }) => ({
          format: run.format,
          scenario: scenario.scenario,
          sample: scenario.sample,
          status: scenario.status,
          actor,
          op,
          errors,
          ...stats,
        })
      )
    )
  );
}

const cell = (value: unknown) =>
  String(value).replaceAll('|', '\\|').replaceAll('\n', ' ');

export function markdown(runs: RecordedRun[]): string {
  const lines: string[] = [];
  for (const run of runs) {
    lines.push(
      '## ' +
        run.format +
        ' (' +
        run.commit.slice(0, 8) +
        ', ' +
        cell(run.environment.cpu) +
        ', Bun ' +
        run.environment.bun +
        ')',
      ''
    );
    lines.push(
      '| scenario | sample | status | operations | reason |',
      '| --- | --- | --- | ---: | --- |'
    );
    for (const scenario of run.scenarios)
      lines.push(
        '| ' +
          [
            scenario.scenario,
            scenario.sample,
            scenario.status,
            scenario.ops.length,
            scenario.reason ?? '',
          ]
            .map(cell)
            .join(' | ') +
          ' |'
      );
    lines.push('');
  }
  lines.push(
    '## Operations by actor',
    '',
    '| format | scenario | sample | actor | operation | n | p50 ms | p95 ms |',
    '| --- | --- | --- | --- | --- | ---: | ---: | ---: |'
  );
  for (const row of rows(runs))
    lines.push(
      '| ' +
        [
          row.format,
          row.scenario,
          row.sample,
          row.actor,
          row.op,
          row.count,
          row.p50Ms.toFixed(2),
          row.p95Ms.toFixed(2),
        ]
          .map(cell)
          .join(' | ') +
        ' |'
    );
  return lines.join('\n') + '\n';
}

export function diff(
  before: RecordedRun[],
  after: RecordedRun[]
): { slower: string[]; faster: string[] } {
  const formats = (runs: RecordedRun[]) =>
    runs
      .map((run) => run.format)
      .sort()
      .join(',');
  if (!before.length || !after.length || formats(before) !== formats(after))
    throw new Error('missing or mismatched result formats');
  if (
    new Set(before.map((run) => run.format)).size !== before.length ||
    new Set(after.map((run) => run.format)).size !== after.length
  )
    throw new Error('duplicate result formats');
  const slower: string[] = [];
  const faster: string[] = [];
  for (const previous of before) {
    const current = after.find((run) => run.format === previous.format)!;
    slower.push(...regressions(previous, current));
    faster.push(...regressions(current, previous));
  }
  return { slower, faster };
}

if (import.meta.main) {
  const args = process.argv.slice(2);
  if (args[0] === '--diff') {
    if (args.length !== 3)
      throw new Error('usage: report.ts --diff <before> <after>');
    const { slower, faster } = diff(readRuns(args[1]), readRuns(args[2]));
    console.log(
      'slower (' + slower.length + '):\n  ' + (slower.join('\n  ') || 'none')
    );
    console.log(
      'faster (' + faster.length + '):\n  ' + (faster.join('\n  ') || 'none')
    );
    process.exitCode = slower.length > 0 ? 1 : 0;
  } else {
    if (
      args.some((arg) => arg.startsWith('--') && arg !== '--json') ||
      args.filter((arg) => arg !== '--json').length > 1
    )
      throw new Error('usage: report.ts [dir] [--json]');
    const runs = readRuns(args.find((arg) => arg !== '--json'));
    console.log(
      args.includes('--json')
        ? JSON.stringify(rows(runs), null, 2)
        : markdown(runs)
    );
  }
}
