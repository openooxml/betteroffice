import { readFileSync, writeFileSync } from 'node:fs';
import { relative, resolve } from 'node:path';

import { browserFiles, shardWeights } from '../e2e/browser/shards';

type Suite = {
  file: string;
  specs?: { tests: { results: { duration: number }[] }[] }[];
  suites?: Suite[];
};
type Report = { config: { rootDir: string }; suites: Suite[] };

const root = resolve(import.meta.dirname, '..');

export function reportDurations(report: Report): Map<string, number> {
  const seconds = new Map<string, number>();
  const visit = (suite: Suite) => {
    const file = relative(root, resolve(report.config.rootDir, suite.file));
    for (const spec of suite.specs ?? [])
      for (const test of spec.tests)
        for (const result of test.results)
          seconds.set(file, (seconds.get(file) ?? 0) + result.duration / 1000);
    suite.suites?.forEach(visit);
  };
  report.suites.forEach(visit);
  return seconds;
}

if (import.meta.main) {
  const reports = process.argv.slice(2);
  if (!reports.length) {
    console.error('usage: bun scripts/browser-shard-weights.ts <playwright report.json>...');
    process.exit(2);
  }
  const measured = new Map<string, number>();
  for (const path of reports)
    for (const [file, seconds] of reportDurations(JSON.parse(readFileSync(path, 'utf8'))))
      measured.set(file, (measured.get(file) ?? 0) + seconds);
  const previous = shardWeights();
  const weights = Object.fromEntries(
    browserFiles()
      .filter((file) => measured.has(file) || file in previous)
      .map((file) => [file, Math.round(measured.get(file) ?? previous[file])]),
  );
  writeFileSync(resolve(root, 'e2e/browser/shard-weights.json'), JSON.stringify(weights, null, 2) + '\n');
}
