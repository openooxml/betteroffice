import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { version } from '../package.json';
import { packageVersion } from './wasm/engine';

test('the engine version import equals package.json', () => {
  const manifest = JSON.parse(
    readFileSync(resolve(import.meta.dir, '../package.json'), 'utf8')
  ) as { version: string };
  expect(packageVersion).toBe(manifest.version);
  expect(version).toBe(manifest.version);
});
