import { expect, test } from 'bun:test';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

test('an act left pending by an ended test fails the next test once and does not block the rest', () => {
  const dir = mkdtempSync(join(tmpdir(), 'test-scope-'));
  try {
    writeFileSync(join(dir, 'pending.test.ts'), `
import { test } from 'bun:test';
import { testScope } from ${JSON.stringify(resolve(import.meta.dir, '__fixtures__/testScope.ts'))};
const { act } = testScope((callback) => callback(), { drainMs: 100 });
test('leaves an act pending', () => { void act(() => new Promise<void>(() => {})); });
test('runs after the abandoned act', () => {});
test('runs after that too', async () => { await act(async () => {}); });
`);
    const run = Bun.spawnSync([process.execPath, 'test', './pending.test.ts'], {
      cwd: dir,
      stdout: 'pipe',
      stderr: 'pipe',
      timeout: 30_000,
    });
    const output = `${run.stdout}\n${run.stderr}`;
    expect(output).toContain('test #1 of this file left 1 unfinished act() or poll call(s)');
    expect(output).toContain('(pass) runs after that too');
    expect(output).toMatch(/\b2 pass\b/);
    expect(output).toMatch(/\b1 fail\b/);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}, 40_000);
