import * as fs from 'node:fs';
import * as path from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';

type FixtureResult = {
  name: string;
  parseMs: number;
  resolveMs: number;
  evaluateMs: number;
  renderMs: number;
  saveMs: number;
};

type RunResult = {
  samples: number;
  fixtures: FixtureResult[];
};

type RecordedRun = RunResult & {
  schemaVersion: 1;
  commit: string;
  regression: {
    percent: number;
    minimumMs: number;
  };
};

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const recordedPath = path.join(root, 'scripts', 'vsdx-bench-results.json');
const stageKeys = ['parseMs', 'resolveMs', 'evaluateMs', 'renderMs', 'saveMs'] as const;
const regression = { percent: 12, minimumMs: 3 };
const record = process.argv.includes('--record');

function run(command: string[], purpose: string): string {
  const result = Bun.spawnSync(command, { cwd: root, stdout: 'pipe', stderr: 'pipe' });
  if (result.exitCode !== 0) {
    throw new Error(`${purpose} failed:\n${new TextDecoder().decode(result.stderr)}`);
  }
  return new TextDecoder().decode(result.stdout);
}

function currentCommit(): string {
  return run(['git', 'rev-parse', 'HEAD'], 'reading the current commit').trim();
}

function readRecordedRun(): RecordedRun | undefined {
  if (!fs.existsSync(recordedPath)) return undefined;
  return JSON.parse(fs.readFileSync(recordedPath, 'utf8')) as RecordedRun;
}

function checkRegression(previous: RecordedRun, current: RunResult): string[] {
  const failures: string[] = [];
  const currentByName = new Map(current.fixtures.map((fixture) => [fixture.name, fixture]));
  for (const before of previous.fixtures) {
    const after = currentByName.get(before.name);
    if (!after) {
      failures.push(`${before.name}: fixture is missing from the current run`);
      continue;
    }
    for (const stage of stageKeys) {
      const delta = after[stage] - before[stage];
      const percent = before[stage] === 0 ? Infinity : (delta / before[stage]) * 100;
      if (isRegression(delta, percent)) {
        failures.push(`${before.name} ${stage}: ${before[stage].toFixed(3)}ms -> ${after[stage].toFixed(3)}ms (+${percent.toFixed(1)}%, +${delta.toFixed(3)}ms)`);
      }
    }
  }
  return failures;
}

function isRegression(delta: number, percent: number): boolean {
  return delta > regression.minimumMs && percent > regression.percent;
}

function signed(value: number, digits: number): string {
  return `${value >= 0 ? '+' : ''}${value.toFixed(digits)}`;
}

function stageResult(stage: typeof stageKeys[number], current: FixtureResult, previous?: FixtureResult): string {
  const name = stage.slice(0, -2);
  const value = current[stage];
  if (!previous) return `${name} ${value.toFixed(3)}ms (no baseline)`;

  const delta = value - previous[stage];
  const percent = previous[stage] === 0 ? Infinity : (delta / previous[stage]) * 100;
  const regressionMarker = isRegression(delta, percent) ? ' !REGRESSION' : '';
  const percentage = Number.isFinite(percent) ? `${signed(percent, 1)}%` : '+∞%';
  return `${name} ${value.toFixed(3)}ms (${signed(delta, 3)}ms, ${percentage})${regressionMarker}`;
}

function printResults(commit: string, result: RunResult, previous?: RecordedRun): void {
  const baseline = previous ? previous.commit : 'none';
  console.log(`VSDX benchmark commit ${commit}; baseline ${baseline}; median of ${result.samples} samples after one warm-up`);
  console.log("Synthetic fixtures; render uses fontless fallback layout, without font shaping.");
  const previousByName = new Map((previous ? previous.fixtures : []).map((fixture) => [fixture.name, fixture]));
  for (const fixture of result.fixtures) {
    const before = previousByName.get(fixture.name);
    console.log(`${fixture.name}: ${stageKeys.map((stage) => stageResult(stage, fixture, before)).join(' | ')}`);
  }
}

function cleanupScratch(scratch: string): void {
  fs.rmSync(scratch, { recursive: true, force: true });
}

const commit = currentCommit();
const dirty = run(['git', 'status', '--porcelain', '--untracked-files=normal'], 'checking benchmark source state').trim() !== '';
if (record && dirty) throw new Error('Commit source changes before recording a benchmark baseline.');
const scratch = fs.mkdtempSync(path.join(tmpdir(), 'betteroffice-vsdx-bench-'));
try {
  run(['bun', 'scripts/create-vsdx-fixture.ts', `--benchmark-dir=${scratch}`], 'generating synthetic fixtures');
  const fixtures = fs.readdirSync(scratch)
    .filter((entry) => entry.endsWith('.vsdx'))
    .sort()
    .map((entry) => `${path.basename(entry, '.vsdx')}=${path.join(scratch, entry)}`);
  const result = JSON.parse(run(['cargo', 'run', '--quiet', '--release', '--offline', '-p', 'betteroffice-vsdx-bench', '--', ...fixtures], 'running staged benchmark')) as RunResult;
  const previous = readRecordedRun();
  printResults(dirty ? `${commit} (working tree changes)` : commit, result, previous);
  if (previous && !record) {
    const failures = checkRegression(previous, result);
    if (failures.length > 0) throw new Error(`VSDX benchmark regression against ${previous.commit}:\n${failures.join('\n')}`);
  }
  if (record) {
    const recorded: RecordedRun = { schemaVersion: 1, commit, regression, ...result };
    fs.writeFileSync(recordedPath, `${JSON.stringify(recorded, null, 2)}\n`);
    console.log(`Recorded baseline in ${path.relative(root, recordedPath)}.`);
  }
} finally {
  cleanupScratch(scratch);
}
