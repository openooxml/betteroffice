import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';

import { assertComplete, baselineDir, parseRecordedRun } from './harness';
import type { RecordedRun } from './harness';
import { diff, readRuns } from './report';

export const FORMATS = ['docx', 'xlsx', 'pptx'] as const;
const ROOT = path.resolve(import.meta.dir, '..');

export function promoteBaseline(
  runs: RecordedRun[],
  destination = baselineDir()
): void {
  if (
    runs
      .map((run) => run.format)
      .sort()
      .join(',') !== [...FORMATS].sort().join(',')
  )
    throw new Error('baseline requires every format exactly once');
  for (const run of runs) assertComplete(parseRecordedRun(run));
  if (fs.existsSync(destination) && !fs.lstatSync(destination).isDirectory())
    throw new Error('baseline destination must be a directory');
  fs.mkdirSync(path.dirname(destination), { recursive: true });
  const stage = fs.mkdtempSync(
    path.join(path.dirname(destination), '.e2e-record-')
  );
  const backup = stage + '-previous';
  try {
    for (const run of runs)
      fs.writeFileSync(
        path.join(stage, run.format + '.json'),
        JSON.stringify(run, null, 2) + '\n'
      );
    if (fs.existsSync(destination)) fs.renameSync(destination, backup);
    try {
      fs.renameSync(stage, destination);
    } catch (error) {
      if (fs.existsSync(backup)) fs.renameSync(backup, destination);
      throw error;
    }
    fs.rmSync(backup, { recursive: true, force: true });
  } finally {
    fs.rmSync(stage, { recursive: true, force: true });
  }
}

if (import.meta.main) {
  const mode = process.argv[2] ?? 'run';
  if (!['run', 'record', 'compare'].includes(mode) || process.argv.length > 3)
    throw new Error('usage: run.ts [run|record|compare]');
  if (mode !== 'run' && process.env.BETTEROFFICE_E2E_ALLOW_SKIP === '1')
    throw new Error('record/compare require every scenario');
  const output = fs.mkdtempSync(path.join(os.tmpdir(), 'betteroffice-corpus-'));
  try {
    const child = Bun.spawn(
      [
        process.execPath,
        'test',
        ...FORMATS.map((format) =>
          path.join(import.meta.dir, format + '.e2e.test.ts')
        ),
      ],
      {
        cwd: ROOT,
        stdout: 'inherit',
        stderr: 'inherit',
        env: {
          ...process.env,
          BETTEROFFICE_E2E: '1',
          BETTEROFFICE_E2E_OUTPUT: output,
        },
      }
    );
    const code = await child.exited;
    const exported = process.env.BETTEROFFICE_E2E_OUTPUT;
    if (exported) {
      fs.mkdirSync(exported, { recursive: true });
      for (const format of FORMATS) {
        const target = path.join(exported, format + '.json');
        fs.rmSync(target, { force: true });
        const source = path.join(output, format + '.json');
        if (fs.existsSync(source)) fs.copyFileSync(source, target);
      }
    }
    if (code !== 0)
      throw new Error('corpus scenarios failed; baseline unchanged');
    const runs = readRuns(output);
    if (
      runs
        .map((run) => run.format)
        .sort()
        .join(',') !== [...FORMATS].sort().join(',')
    )
      throw new Error('missing format results; baseline unchanged');
    if (mode === 'record') {
      promoteBaseline(runs);
      console.log('Recorded complete baseline at ' + baselineDir());
    } else if (mode === 'compare') {
      const result = diff(readRuns(baselineDir()), runs);
      if (result.slower.length) throw new Error(result.slower.join('\n'));
    }
  } finally {
    fs.rmSync(output, { recursive: true, force: true });
  }
}
