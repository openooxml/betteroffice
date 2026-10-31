import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { engineKind as pptxEngine } from '../packages/pptx/src/wasm/engine';
import { engineKind as xlsxEngine } from '../packages/xlsx/src/wasm/engine';
import { engineKind as docxEditEngine } from '../packages/docx/src/wasm/engineEdit';
import { engineKind as docxParseEngine } from '../packages/docx/src/wasm/engineParse';
import { engineKind as docxQueryEngine } from '../packages/docx/src/wasm/engineQuery';
import { engineKind as docxOpcEngine } from '../packages/docx/src/wasm/engineOpc';
import { engineKind as docxPreflightEngine } from '../packages/docx/src/wasm/enginePreflight';
import { NATIVE_ENGINE_FILTER, NATIVE_ENGINES, NATIVE_ASSETS, NATIVE_ENTRIES, nativeBuildSource, nativeEngineSource } from './native-engine';

const expected =
  (globalThis as Record<symbol, unknown>)[Symbol.for('betteroffice.testEngine')] ?? 'wasm';
const root = resolve(import.meta.dir, '..');

test('the engine modules match the engine this run links', () => {
  for (const engine of [xlsxEngine, pptxEngine, docxEditEngine, docxParseEngine, docxQueryEngine, docxOpcEngine, docxPreflightEngine]) {
    expect(engine).toBe(expected);
  }
});

test('the native preload relinks each engine and asset separately', () => {
  for (const [format, value] of Object.entries(NATIVE_ENGINES)) {
    for (const { engine, native } of (Array.isArray(value) ? value : [value]) as readonly { engine: string; native: string }[]) {
      const path = resolve(root, engine);
      const source = `export * from ${JSON.stringify(resolve(root, native))};\nexport { default } from ${JSON.stringify(resolve(root, native))};\n`;
      expect(NATIVE_ENGINE_FILTER.test(path)).toBe(true);
      expect(nativeEngineSource(path)).toBe(source);
      expect(nativeBuildSource(format as keyof typeof NATIVE_ENGINES, path)).toBe(source);
      if (format !== 'docx') expect(nativeBuildSource('docx', path)).toBeUndefined();
    }
  }
  for (const [format, value] of Object.entries(NATIVE_ASSETS)) {
    for (const { asset, native } of (Array.isArray(value) ? value : [value]) as readonly { asset: string; native: string }[]) {
      const path = resolve(root, asset);
      const source = `export * from ${JSON.stringify(resolve(root, native))};\n`;
      expect(NATIVE_ENGINE_FILTER.test(path)).toBe(true);
      expect(nativeEngineSource(path)).toBe(Array.isArray(value) ? source : undefined);
      expect(nativeBuildSource(format as keyof typeof NATIVE_ENGINES, path)).toBe(source);
      const otherFormat = format === 'docx' ? 'xlsx' : 'docx';
      expect(nativeBuildSource(otherFormat, path)).toBeUndefined();
    }
  }
  expect(nativeEngineSource(resolve(root, 'packages/docx/src/wasm/engine.ts'))).toBeUndefined();
  expect(nativeEngineSource(resolve(root, 'packages/xlsx/src/session/engine.ts'))).toBeUndefined();
});

test('the wasm packages never reference their native packages', () => {
  for (const [format, value] of Object.entries(NATIVE_ENGINES)) {
    for (const { engine, native } of (Array.isArray(value) ? value : [value]) as readonly { engine: string; native: string }[]) {
      const source = readFileSync(resolve(root, engine), 'utf8');
      expect(source).not.toContain(`${format}-native`);
      expect(source).not.toContain('bindings/');
      expect(source).not.toContain('createRequire');
      expect(readFileSync(resolve(root, native), 'utf8')).toContain("engineKind: 'wasm' | 'native' = 'native'");
    }
    const manifest = JSON.parse(readFileSync(resolve(root, `packages/${format}/package.json`), 'utf8'));
    expect(JSON.stringify(manifest)).not.toContain(`@betteroffice/${format}-native`);
  }
});

test('docx export sources cover exactly its public subpaths', () => {
  const manifest = JSON.parse(readFileSync(resolve(root, 'packages/docx/package.json'), 'utf8'));
  expect(Object.keys(NATIVE_ENTRIES.docx)).toEqual(Object.keys(manifest.exports));
  expect(Object.keys(NATIVE_ENTRIES.docx)).toHaveLength(28);
  const config = readFileSync(resolve(root, 'packages/docx/tsup.config.ts'), 'utf8');
  for (const entry of Object.values(NATIVE_ENTRIES.docx)) expect(config).toContain(`'${entry}'`);
});
