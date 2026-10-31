import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { engineKind as pptxEngine } from '../packages/pptx/src/wasm/engine';
import { engineKind as xlsxEngine } from '../packages/xlsx/src/wasm/engine';
import { NATIVE_ENGINE_FILTER, NATIVE_ENGINES, nativeEngineSource } from './native-engine';

const expected =
  (globalThis as Record<symbol, unknown>)[Symbol.for('betteroffice.testEngine')] ?? 'wasm';
const root = resolve(import.meta.dir, '..');

test('the engine modules match the engine this run links', () => {
  expect(xlsxEngine).toBe(expected);
  expect(pptxEngine).toBe(expected);
});

test('the native preload relinks only the wasm engine modules', () => {
  for (const { engine, native } of Object.values(NATIVE_ENGINES)) {
    expect(NATIVE_ENGINE_FILTER.test(resolve(root, engine))).toBe(true);
    expect(nativeEngineSource(resolve(root, engine))).toBe(
      `export * from ${JSON.stringify(resolve(root, native))};\nexport { default } from ${JSON.stringify(resolve(root, native))};\n`
    );
  }
  expect(nativeEngineSource(resolve(root, 'packages/docx/src/wasm/engine.ts'))).toBeUndefined();
  expect(nativeEngineSource(resolve(root, 'packages/xlsx/src/session/engine.ts'))).toBeUndefined();
});

test('the wasm packages never reference their native packages', () => {
  for (const [format, { engine, native }] of Object.entries(NATIVE_ENGINES)) {
    const source = readFileSync(resolve(root, engine), 'utf8');
    expect(source).not.toContain(`${format}-native`);
    expect(source).not.toContain('bindings/');
    expect(source).not.toContain('createRequire');
    const manifest = JSON.parse(readFileSync(resolve(root, `packages/${format}/package.json`), 'utf8'));
    expect(JSON.stringify(manifest)).not.toContain(`@betteroffice/${format}-native`);
    expect(readFileSync(resolve(root, native), 'utf8')).toContain("engineKind: 'wasm' | 'native' = 'native'");
  }
});
