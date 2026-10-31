import { expect, test } from 'bun:test';
import { resolve } from 'node:path';
import { cliEnv } from './cli-env';

test('stdio children run the engine this test run links', () => {
  const engine = resolve(import.meta.dir, '../../xlsx/src/wasm/engine.ts');
  const child = Bun.spawnSync(
    [process.execPath, '-e', `console.log((await import(${JSON.stringify(engine)})).engineKind)`],
    { env: cliEnv() }
  );
  expect(child.stdout.toString().trim()).toBe(
    String((globalThis as Record<symbol, unknown>)[Symbol.for('betteroffice.testEngine')] ?? 'wasm')
  );
});
