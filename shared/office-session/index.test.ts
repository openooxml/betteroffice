import { expect, it } from 'bun:test';
import { spawnSync } from 'node:child_process';

it('imports in a fresh process without browser globals', () => {
  const entry = new URL('./index.ts', import.meta.url).href;
  const result = spawnSync(process.execPath, ['--eval', `
    for (const name of ['window', 'document', 'self', 'Worker']) {
      if (!Reflect.deleteProperty(globalThis, name) || name in globalThis) {
        throw new Error('Cannot remove global ' + name);
      }
    }
    const session = await import(${JSON.stringify(entry)});
    if (typeof session.createSessionClient !== 'function' ||
        typeof session.createSessionHost !== 'function' || 'createInProcessPair' in session) {
      throw new Error('Unexpected session exports');
    }
  `], { encoding: 'utf8', timeout: 5000 });
  expect(result.stderr).toBe('');
  expect(result.status).toBe(0);
});
