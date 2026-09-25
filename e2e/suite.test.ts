import { expect, test } from 'bun:test';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { parseRecordedRun } from './harness';

for (const fault of ['body', 'context', 'setup', 'dispose', 'filter', 'skip']) {
  test('preserves the full case inventory after ' + fault + ' failure', () => {
    const dir = mkdtempSync(join(tmpdir(), 'e2e-suite-test-'));
    const file = join(dir, 'injected.test.ts');
    writeFileSync(
      file,
      `
import { mock } from 'bun:test';
mock.module(${JSON.stringify(join(import.meta.dir, 'corpus.ts'))}, () => ({
  samplesFor: () => [{ id: 'fixture', format: 'xlsx', sha256: '${'a'.repeat(
    64
  )}' }],
  loadSample: async () => new Uint8Array(),
}));
const { defineSuite } = await import(${JSON.stringify(
        join(import.meta.dir, 'suite.ts')
      )});
const fault = ${JSON.stringify(fault)};
defineSuite('xlsx', ['passing', 'failing'].map((name) => ({
  name, description: name, participants: ['wasm'],
  requires: () => name === 'failing' && fault === 'skip' ? 'missing SDK' : undefined,
  run() { if (name === 'failing' && fault === 'body') throw new Error('body failure'); },
})), {
  setup() { if (fault === 'setup') throw new Error('setup failure'); },
  context(_sample, _bytes, recorder) {
    if (recorder.meta.scenario === 'failing' && fault === 'context') throw new Error('context failure');
    return { dispose() { if (recorder.meta.scenario === 'failing' && fault === 'dispose') throw new Error('dispose failure'); } };
  },
});
`
    );
    try {
      const child = Bun.spawnSync(
        [
          process.execPath,
          'test',
          file,
          ...(fault === 'filter' ? ['--test-name-pattern', 'passing'] : []),
        ],
        {
          env: {
            ...process.env,
            BETTEROFFICE_E2E: '1',
            BETTEROFFICE_E2E_OUTPUT: dir,
            BETTEROFFICE_E2E_ALLOW_SKIP: '0',
            GITHUB_STEP_SUMMARY: '',
          },
        }
      );
      expect(child.exitCode).not.toBe(0);
      const run = parseRecordedRun(
        JSON.parse(readFileSync(join(dir, 'xlsx.json'), 'utf8'))
      );
      expect(run.scenarios.map((s) => s.scenario)).toEqual([
        'passing',
        'failing',
      ]);
      expect(run.scenarios[1].status).toBe(
        fault === 'skip' ? 'skipped' : 'failed'
      );
      if (fault !== 'filter')
        expect(run.scenarios[1].reason).toContain(
          fault === 'skip' ? 'missing SDK' : fault + ' failure'
        );
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
}
